#!/usr/bin/env node
/**
 * check-num.js —— 数据表里的数字自洽机检。
 *
 * 用法：node check-num.js <文件.docx> [--strict]
 *       稿源脚本由 writeDoc 写盘后自动调用；--strict 时非零退出。
 *
 * 为什么有它：check-doc 只查格式（单位空格、小数 0、日期写法），不查数字本身
 * 对不对。表内合计、百分比构成、同一指标在两张表里的取值——全靠手算，改一个
 * 参数要人工复算全部关联表。这条把它机检化。
 *
 * **拦截只收无歧义的四类，宁漏勿滥**（检查项自己制造假阳性比没有检查更糟）：
 *   1. 合计行／合计列：标了「合计／总计／小计」的行列，等于其余数值格之和
 *   2. 百分比构成：标了百分号的合计，且该行／列是占比，和应为 100
 *   3. 同指标跨表同值：同一文字标签在两张数据表里出现不同数值
 *   4. 计数与列举：正文「以下三点／如下四个方面」与其后实际列出的条数
 *      （下一级标题，或「一是、二是」行内列举）对不上
 *
 * **提示不拦稿**（要靠标签匹配，有误报可能，交给人判断）：
 *   5. 正文同指标异值：「××为／达 N 单位」在正文多处出现、数值不同
 *      （标签＋单位＋同句年份作键，2025 年与 2026 年的收入不会被当成同一项）
 *   6. 正文与表格异值：正文「××为 N 单位」与表格同名行的值不同；
 *      只比单值列、且表头写明了同一单位的表
 *
 * 一格不是干净数字（空格、中文单位、括号负数、破折号）就不参与这一列，不当 0，
 * 也不报错——那是 check-doc 与人工的活。数值容差见 TOL。
 *
 * 合并格（w:gridSpan）按真实列位展开——读取在 table-read.js，与 check-doc 共用。
 * 读取走 jszip，不调 unzip（Windows 上没有后者）。
 */
"use strict";
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { textOf, rowCells } = require("../lib/table-read");

// 相对容差 0.5%，绝对下限 0.02——印刷稿的舍入误差不会比这更大。
const TOL_REL = 0.005;
const TOL_ABS = 0.02;
const PCT_TOL = 0.5;          // 百分比构成和与 100 的容差
const TOTAL_LABEL = /^(合计|总计|小计|合\s*计|总\s*计|小\s*计)$/;

// 去掉某个标签的整块，按嵌套深度配对（非贪婪正则遇到表中表、文本框会提前收尾）
function stripBlocks(xml, tag) {
  const open = new RegExp(`<${tag}[ >]`, "g");
  const close = `</${tag}>`;
  let out = "", i = 0;
  for (;;) {
    open.lastIndex = i;
    const m = open.exec(xml);
    if (!m) return out + xml.slice(i);
    out += xml.slice(i, m.index);
    let depth = 1, j = m.index + m[0].length;
    while (depth && j < xml.length) {
      open.lastIndex = j;
      const o = open.exec(xml);
      const c = xml.indexOf(close, j);
      if (c < 0) return out;
      if (o && o.index < c) { depth++; j = o.index + o[0].length; } else { depth--; j = c + close.length; }
    }
    i = j;
  }
}

async function loadParts(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file("word/document.xml").async("string");
  const st = zip.file("word/styles.xml");
  return { xml, styles: st ? await st.async("string") : "" };
}

// 不要求首行声明 tblHeader：本工具生成的表都有，别人手写的 Word 表大多没有，
// 体检现成文件时一律跳过等于没查。版式表格（联合行文标志等）只有一行，或者
// 没有「合计」行和干净数字，判据本身不会命中。
// rows.idx 是它在全文表格里的序号（修复工具挂批注用）。
// 比率类标签：以「率」「单价」等结尾（去掉末尾括注的单位），或以「平均」开头。
// 只认词尾：「功率模块」「效率提升工程」是普通明细，不是比率。
function isRateLabel(s) {
  const t = (s || "").replace(/[（(][^（()）]*[）)]\s*$/, "").trim();
  return /(?:率|增速|增幅|单价|均价|倍数)$/.test(t) || /^平均/.test(t);
}

function readTables(xml) {
  return (xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) || [])
    .map((tb, idx) => {
      const trs = tb.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
      const rows = trs.map((tr) => rowCells(tr, textOf));
      // 表尾的无边框定高空行不是数据行（展开后是一串 null／空串）
      if (rows.length && rows[rows.length - 1].every((c) => !c)) rows.pop();
      rows.idx = idx;
      rows.declared = /<w:tblHeader/.test(tb);
      return rows;
    })
    .filter((rows) => rows.length >= 2);
}

// 干净数字才返回对象：可带千分位、正负号、百分号。带中文单位、括号、范围一律 null。
function num(s) {
  if (typeof s !== "string") return null;
  const t = s.replace(/[,\s]/g, "");
  const m = t.match(/^(-?\d+(?:\.\d+)?)([%％]?)$/);
  if (!m) return null;
  const decimals = m[1].includes(".") ? m[1].split(".")[1].length : 0;
  return { v: Number(m[1]), pct: !!m[2], decimals, raw: s };
}

const round = (n) => Math.round(n * 1000) / 1000;
const close = (a, b) => Math.abs(a - b) <= Math.max(TOL_ABS, TOL_REL * Math.abs(b));

async function check(file) {
  const { xml, styles } = await loadParts(file);
  const tables = readTables(xml);
  const hits = [];
  const hints = [];
  const add = (rule, detail, anchors = []) => hits.push({ rule, detail, anchors });
  const hint = (rule, detail, anchors = []) => hints.push({ rule, detail, anchors });

  tables.forEach((rows, ti) => {
    const head = rows[0];
    // 手写表（首行没声明表头）只提示不拦：跨页拆开的表、只列前几名的表，
    // 合计本来就不等于本页各项之和，工具分不出来（真实稿回放见过这两种）
    const report = rows.declared ? add : hint;
    // 没声明表头、首行又有数字的，首行也是数据行
    const first = !rows.declared && rows[0].some((c) => num(c)) ? 0 : 1;

    // ── 1 合计行 ──────────────────────────────────────────
    // 分段求和：小计只加它上面那一段（上一个合计类行之后）的数据行；合计、总计
    // 前面有小计的，对各小计之和，没有小计的对上面那一段。合计行写在表首的，
    // 对它下面那一段。真实稿回放（手写的财务表）里一张表常有好几段小计，整列
    // 一起加必然对不上。
    // 比率列（毛利率、增长率、单价……）的合计行是加权平均，不是加总，不查；
    // 占比、比例、比重这类构成列照查。
    const isTotalRow = (cells) => cells.some((c) => TOTAL_LABEL.test(c || ""));
    const isRateCol = (h) => isRateLabel(h) && !/占比|比例|比重|构成/.test(h || "");
    // 比率行（「预期增长率」「单价」）夹在明细里，不进求和
    const isRateRow = (cells) => cells.some((c) => c && !num(c) && isRateLabel(c));
    rows.forEach((cells, ri) => {
      if (ri < first) return;
      const labelIdx = cells.findIndex((c) => TOTAL_LABEL.test(c || ""));
      if (labelIdx < 0) return;
      const label = cells[labelIdx].replace(/\s/g, "");
      // 上一段：往上找到上一个合计类行（小计找任意合计类行，合计／总计只找合计／总计）
      const stop = (r) => (label === "小计" ? isTotalRow(rows[r]) : rows[r].some((c) => /^(合|总)\s*计$/.test(c || "")));
      let from = ri - 1;
      while (from >= first && !stop(from)) from--;
      let seg = [];
      for (let r = from + 1; r < ri; r++) if (!isRateRow(rows[r])) seg.push(r);
      if (label !== "小计") {
        const subs = seg.filter((r) => rows[r].some((c) => /^小\s*计$/.test(c || "")));
        if (subs.length) seg = subs;
        else seg = seg.filter((r) => !isTotalRow(rows[r]));
      }
      if (!seg.length) {   // 合计写在表首：对下面那一段
        for (let r = ri + 1; r < rows.length && !isTotalRow(rows[r]); r++) if (!isRateRow(rows[r])) seg.push(r);
      }
      for (let ci = 0; ci < cells.length; ci++) {
        if (ci === labelIdx || isRateCol(head[ci])) continue;
        const total = num(cells[ci]);
        if (!total) continue;
        // 这一段的数据格都必须是干净数字，才判得了
        const vals = [];
        let clean = true;
        for (const r of seg) {
          const n = num(rows[r][ci]);
          if (!n) { clean = false; break; }
          // 合计是绝对数，分项里的百分数是比率行（有的连标签都没写），不加
          if (n.pct && !total.pct) continue;
          vals.push(n.v);
        }
        // 手写表只有一项的段不查：「合计」常是另一个口径（跨页续表、只列一项的明细）
        if (!clean || vals.length < (rows.declared ? 1 : 2)) continue;
        const sum = vals.reduce((x, y) => x + y, 0);
        if (!close(sum, total.v)) {
          report("合计不符", `第 ${ti + 1} 表「${cells[labelIdx]}」行第 ${ci + 1} 列：上面 ${vals.length} 项之和 ${round(sum)} ≠ 表内${cells[labelIdx]} ${total.raw}`, [{ text: total.raw, table: rows.idx }]);
        } else if (total.pct && Math.abs(sum - 100) > PCT_TOL) {
          // 合计与各项对得上、只是不到 100%：「前五名合计占比」这类表本来就是部分，只提示
          hint("百分比构成", `第 ${ti + 1} 表「${cells[labelIdx]}」行第 ${ci + 1} 列：占比合计 ${round(sum)}%，不是 100%；是部分构成就不用改`, [{ text: total.raw, table: rows.idx }]);
        }
      }
    });

    // ── 2 合计列（表头标了合计的那一列，等于本行其余列之和）────
    head.forEach((h, ci) => {
      if (!TOTAL_LABEL.test(h)) return;
      if (head.some((x, k) => k > 0 && k !== ci && isRateCol(x))) return;   // 混着比率列，行内加总无意义
      for (let r = 1; r < rows.length; r++) {
        // 合计行已在第 1 段查过，别从列的方向再报一遍；余额、单价、比率这类行按行加总没有意义
        if (TOTAL_LABEL.test(rows[r][0] || "") || /余额|期初|期末|累计/.test(rows[r][0] || "") || isRateLabel(rows[r][0])) continue;
        const total = num(rows[r][ci]);
        if (!total) continue;
        const vals = [];
        let clean = true;
        for (let k = 1; k < head.length; k++) {
          if (k === ci) continue;
          const n = num(rows[r][k]);
          if (!n) { clean = false; break; }
          vals.push(n.v);
        }
        if (!clean || !vals.length) continue;
        const sum = vals.reduce((a, b) => a + b, 0);
        if (!close(sum, total.v)) {
          report("合计不符", `第 ${ti + 1} 表第 ${r + 1} 行「${head[ci]}」列：其余列之和 ${round(sum)} ≠ ${total.raw}`, [{ text: total.raw, table: rows.idx }]);
        }
      }
    });
  });

  // ── 3 同指标跨表同值 ────────────────────────────────────
  // **只在表头完全相同时才比**：历史表与预测表都有「营业收入」，但年份区间不同，
  // 数值本就不同——拿整行比会满屏误报。表头一致还出现同标签不同值，才是真的对不上。
  // **只比首行声明了表头的表**（本工具生成的表）。手写的财务表里同名行常是
  // 不同口径（管理费用表与销售费用表都有「办公费」），真实稿回放 105 条全是这类。
  const seen = new Map();
  tables.forEach((rows, ti) => {
    if (!rows.declared) return;
    const head = rows[0];
    const schema = head.join("｜");
    rows.slice(1).forEach((cells) => {
      const label = (cells[0] || "").trim();
      if (!label || label.length < 2 || TOTAL_LABEL.test(label)) return;
      if (!/[\u4e00-\u9fa5]/.test(label)) return;
      const vals = cells.slice(1).map(num).filter(Boolean).map((n) => n.v);
      if (!vals.length) return;
      const key = schema + "::" + label;
      const sig = vals.map(round).join("｜");
      if (seen.has(key)) {
        const prev = seen.get(key);
        if (prev.sig !== sig) {
          add("跨表同值", `「${label}」（表头 ${head.join("/")}）在第 ${prev.ti + 1} 表是 ${prev.sig}，第 ${ti + 1} 表是 ${sig}`, [{ text: label, table: rows.idx }]);
        }
      } else {
        seen.set(key, { ti, sig });
      }
    });
  });

  // ── 4～6 正文 ──────────────────────────────────────────
  const { paras, unitTables } = readBody(xml, styles, tables);
  checkEnumeration(paras, add);
  checkProseValues(paras, unitTables, hint);
  hits.hints = hints;
  return hits;
}

// ── 正文读取 ──────────────────────────────────────────────
// 标题层级的来源，按可靠程度依次认：
//   1. 段落样式在 styles.xml 里的名字是 heading N，或样式带 outlineLvl
//      （中文版 Word 的标题样式 ID 是 "1"／"2"，不是 Heading1，只认 ID 会全部漏掉）
//   2. 段落自己带 outlineLvl
//   3. 手打的层次序数：短段落以「一、」「（一）」开头、不以句号结尾
function headingLevels(styles) {
  const map = new Map();
  for (const st of styles.match(/<w:style [^>]*w:type="paragraph"[\s\S]*?<\/w:style>/g) || []) {
    const id = (st.match(/w:styleId="([^"]+)"/) || [])[1];
    const name = (st.match(/<w:name w:val="([^"]+)"/) || [])[1] || "";
    const byName = name.match(/^heading\s*(\d)$/i);
    const byOutline = st.match(/<w:outlineLvl w:val="(\d)"/);
    if (id && byName) map.set(id, Number(byName[1]));
    else if (id && byOutline) map.set(id, Number(byOutline[1]) + 1);
  }
  return map;
}
const MANUAL_H = [[/^[一二三四五六七八九十]+、/, 1], [/^（[一二三四五六七八九十]+）/, 2]];

function readBody(xml, styles, tables) {
  const levels = headingLevels(styles);
  let body = stripBlocks(xml, "w:tbl");
  body = stripBlocks(body, "w:txbxContent");
  const paras = (body.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map((p) => {
    const text = textOf(p);
    const st = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || "";
    let level = (st.match(/^Heading(\d)$/) || [])[1] ? Number(st.slice(7)) : levels.get(st) || 0;
    const ol = p.match(/<w:pPr>[\s\S]*?<w:outlineLvl w:val="(\d)"/);
    if (!level && ol && Number(ol[1]) < 9) level = Number(ol[1]) + 1;
    if (!level && text.length <= 40 && !/[。；]$/.test(text)) {
      for (const [re, lv] of MANUAL_H) if (re.test(text)) { level = lv; break; }
    }
    return { level, text };
  }).filter((p) => p.text);
  // 表格里能和正文对照的：两列（标签｜数值）、表头写了单位的表
  const unitTables = tables.map((rows) => {
    const head = rows[0].join("");
    const unit = (head.match(/[（(]([^）)]+)[）)]/) || [])[1];
    if (!unit || rows[0].length !== 2) return null;
    const vals = new Map();
    rows.slice(1).forEach((c) => { const n = num(c[1]); if (c[0] && n) vals.set(c[0].trim(), n.v); });
    return { unit: unit.trim(), vals };
  }).filter(Boolean);
  return { paras, unitTables };
}

const CN = { 一: 1, 二: 2, 两: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
const cnNum = (t) => (/^\d+$/.test(t) ? Number(t)
  : t.length === 1 ? CN[t]
    : t[0] === "十" ? 10 + (CN[t[1]] || 0)
      : (CN[t[0]] || 0) * 10 + (CN[t[2]] || 0));
const ENUM_RE = /(?:以下|如下|下列)([一二两三四五六七八九十]{1,3}|\d{1,2})(?:个)?(?:方面|点|条|项|类|件事|项工作|个问题)/;

// 4 计数与列举
function checkEnumeration(paras, add) {
  let section = 0;   // 当前所在章节的标题层级（0＝还没进任何章节）
  paras.forEach((p, i) => {
    if (p.level) { section = p.level; return; }
    const m = p.text.match(ENUM_RE);
    if (!m) return;
    const want = cnNum(m[1]);
    const tail = p.text.slice(m.index + m[0].length);

    // a. 行内列举：只看本段剩余部分；本段在计数句处就结束了，再看紧接的下一段。
    //    「一是」必须在句首或标点之后（「统一是」「唯一是」不算）
    let inline = tail;
    if (!/[一二三四五六七八九十]是/.test(tail) && paras[i + 1] && !paras[i + 1].level) inline = paras[i + 1].text;
    const marks = [...inline.matchAll(/(?:^|[，。；：、\s])([一二三四五六七八九十]{1,2})是/g)].map((x) => cnNum(x[1]));
    if (marks[0] === 1 && marks.length >= 2) {
      let n = 1;
      while (marks.includes(n + 1)) n++;
      if (n !== want) add("计数与列举", `「${m[0]}」其后行内列了 ${n} 条（一是…${n === 1 ? "" : "、" + "一二三四五六七八九十"[n - 1] + "是"}）`, [m[0]]);
      return;
    }
    // b. 下一级标题：只在章节之内判（全文开头的计数句，名下有几条判不清）
    if (!section) return;
    const next = paras.slice(i + 1).find((q) => q.level);
    if (!next || next.level !== section + 1) return;
    let n = 0;
    for (let k = i + 1; k < paras.length; k++) {
      if (paras[k].level && paras[k].level <= section) break;
      if (paras[k].level === section + 1) n++;
    }
    if (n !== want) add("计数与列举", `「${m[0]}」其后列了 ${n} 个下一级标题`, [m[0]]);
  });
}

// 5、6 正文指标取值（提示）
const NUM = "(-?\\d{1,3}(?:,\\d{3})+(?:\\.\\d+)?|-?\\d+(?:\\.\\d+)?)";
const VAL_RE = new RegExp("([\\u4e00-\\u9fa5A-Za-z]{2,16}?)(?:为|达到|达|共计|合计|共有|共)\\s*" + NUM
  + "\\s*(%|％|万亿元|亿元|万元|元|万人|个|项|家|人|名|亩|平方米|公里|吨|台|套)", "g");
// 正文标签与表格行名之间允许的修饰前缀（「全年营业收入」对「营业收入」）
const TABLE_PREFIX = /^(全年|本年度?|当年|年度|累计|实现|全市|全省|全区|全县|全国)*$/;

function checkProseValues(paras, unitTables, hint) {
  const seen = new Map();      // key → { label, unit, years, vals:Set }
  const tableHits = new Map(); // 去重：同一处不一致只报一次
  for (const p of paras) {
    if (p.level) continue;
    for (const sent of p.text.split(/[。；！？]/)) {
      for (const m of sent.matchAll(VAL_RE)) {
        // 年份取这个数之前最近的一个：「2025 年收入为 120，2026 年收入为 135」是两项
        const before = sent.slice(0, m.index);
        const ys = before.match(/\d{4}\s*年/g);
        const years = ys ? ys[ys.length - 1].replace(/\s/g, "") : "";
        const raw = m[1].replace(/^[年月日]/, "");
        // 真实稿回放（60 份业务文件）里的三类误报，逐类排除：
        //   「分别为 90、105」本来就是多个值；「由 500 万元变更为 450 万元」的「为」
        //   跟在变动动词后面，不是「指标为」；「复合增长率约 12%」没有主语，
        //   各细分市场各有各的增长率
        if (/分别$/.test(raw) || /(变更|调整|改|成|增加|减少|增至|降至|提高|降低|上升|下降|增长|扩大|缩减)$/.test(raw)) continue;
        if (/[万亿元%％]/.test(raw)) continue;
        const core = raw.replace(/(约|累计|预计|预计将|将|已)$/, "");
        if (!core || core.length < 2 || /^(其中|合计|总计|共|约)$/.test(core)) continue;
        if (/^(年均)?(复合)?(增长率|增速|增幅|占比|比例|比重|规模|市场规模|同比增长|环比增长|毛利率|净利率)$/.test(core)) continue;
        // 主语：同一分句里、标签之前被空格隔开的那一截（「A 公司营业收入」的 A）
        const clause = before.split(/[，、：；（）]/).pop().replace(/\d{4}\s*年/g, "").replace(/\s+/g, "");
        const label = clause + core;
        const unit = m[3].replace("％", "%");
        const v = Number(m[2].replace(/,/g, ""));
        const key = `${label}｜${unit}｜${years}`;
        if (!seen.has(key)) seen.set(key, { label, unit, years, vals: new Set(), at: [] });
        seen.get(key).vals.add(v);
        seen.get(key).at.push(m[0].trim());
        for (const t of unitTables) {
          if (t.unit !== unit) continue;
          for (const [tl, tv] of t.vals) {
            if (tl.length < 2 || !label.endsWith(tl) || !TABLE_PREFIX.test(label.slice(0, -tl.length))) continue;
            if (!close(v, tv)) tableHits.set(`${tl}|${v}|${tv}`, [`「${tl}」正文是 ${v}${unit}，表格里是 ${tv}${unit}`, m[0].trim()]);
          }
        }
      }
    }
  }
  for (const { label, unit, years, vals, at } of seen.values()) {
    if (vals.size > 1) hint("正文同指标异值", `「${label}」${years ? "（" + years + "）" : ""}在正文里出现了 ${[...vals].map((x) => x + unit).join("、")}`, at);
  }
  for (const [msg, at] of tableHits.values()) hint("正文与表格异值", msg, [at]);
}

module.exports = { check };

// ── 入口 ─────────────────────────────────────────────────
if (require.main === module) (async () => {
  const file = process.argv[2];
  if (!file) {
    console.error("用法：node check-num.js <文件.docx> [--strict] [--json]");
    process.exit(2);
  }
  const strict = process.argv.includes("--strict");
  let hits;
  try {
    hits = await check(file);
  } catch (e) {
    console.error("机检读不进这份文件（工具故障，不是稿子的问题）：", e.message);
    process.exit(2);
  }
  const hints = hits.hints || [];
  if (process.argv.includes("--json")) {
    const row = (blocking) => (h) => ({ rule: h.rule, detail: h.detail, blocking, anchors: h.anchors || [] });
    console.log(JSON.stringify({
      check: "数字自洽", file: path.basename(file),
      hits: [...hits.map(row(true)), ...hints.map(row(false))],
    }));
    if (hits.length && strict) process.exit(1);
    return;
  }
  const showHints = () => {
    if (!hints.length) return;
    console.log(`  提示 ${hints.length} 处（不拦稿，请人工核对）：`);
    hints.forEach((h) => console.log(`    · ${h.rule}：${h.detail}`));
  };
  if (!hits.length) {
    console.log(`数字自洽通过　${path.basename(file)}`);
    showHints();
  } else {
    console.log(`数字自检 ${hits.length} 处　${path.basename(file)}`);
    hits.forEach((h) => console.log(`  · ${h.rule}：${h.detail}`));
    showHints();
    if (strict) process.exit(1);
  }
})();
