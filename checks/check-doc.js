#!/usr/bin/env node
/**
 * check-doc.js —— 体例与语言机检（公文体）。
 *
 * 用法：node check-doc.js <文件.docx> [--strict] [--json]
 *       --json 时只在 stdout 输出一个 JSON：{check, file, hits:[{rule, detail, blocking}]}，退出码不变。
 *       稿源脚本由 writeDoc 写盘后自动调用；--strict 时命中即非零退出。
 *
 * 查四类：
 *   版式与要素（GB/T 9704）：版心与行字数网格、层次序数、成文日期、发文字号
 *   数字（GB/T 15835）：八条
 *   标点与表格（GB/T 15834、GB/T 1.1）：七条
 *   结构：空文档、序列化坏掉、标题重复编号
 *
 * 两类判据：规则能写成「命中即错」的 blocking；依赖标签/对齐/语境、可能
 * 误报的只做 hint（提示），不拦稿。2026-10-23 起 `自称/称对方/表头夹英文/
 * 斜体/数值列非数值` 属 hint——把风格偏好包装成规范是这套体系最不能犯的错。
 *
 * 读取用 jszip，不调 unzip —— 后者在 Windows 上不存在。
 * 词表类文风检查在 check-style（扫描器 style-scan.js）。
 */
"use strict";
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { textOf, rowCells, tableRanges } = require("../lib/table-read");

// ── 取 docx 的结构与文本 ──────────────────────────────────
async function readDocx(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file("word/document.xml").async("string");

  // 表格区间。段落在不在表里，有几条规则要分开看——表里的「1.2」是个
  // 数值，正文里的「1.2 」是个编号，同一个形状两种意思。开标签带属性
  // （w:rsidR 等）也要认，否则整张表会对机检静默消失。
  const tblRanges = tableRanges(xml);
  const inTbl = (pos) => tblRanges.some(([a, b]) => pos >= a && pos < b);

  // 段落级：每个 w:p 拼成一条记录，保留编号与样式信息。
  // at 必须指向 `<w:p` 本身，不能指向 split 后那段的开头：表后第一段的
  // 段首在上一张表的 `</w:p>` 之后、`</w:tbl>` 之前，指错就会把正文段
  // 当成表格内容，从而漏掉 1.1 式编号等规则。
  let cursor = 0;
  const paras = xml.split("</w:p>").map((p) => {
    const rel = p.indexOf("<w:p");
    const at = cursor + (rel >= 0 ? rel : 0);
    cursor += p.length + "</w:p>".length;
    return {
      text: textOf(p),
      numbered: /<w:numPr>/.test(p),
      style: (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || "",
      inTable: inTbl(at),
    };
  }).filter((p) => p.text);

  // 表格：全部表格都读，首行声明 tblHeader 的才 blocking；手写表只提示。
  // 单元格 jc 另存一份（数值列口径检查用）；合并格（gridSpan）按真实列位
  // 展开、合并位为 null；纵向合并续行也是 null，不是空格。
  const cellAlign = (tc) => ((tc.match(/<w:jc w:val="([^"]+)"/) || [])[1]) || "left";
  const tables = [];
  const tableAligns = [];
  for (const [idx, m] of [...xml.matchAll(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g)].entries()) {
    const tb = m[0];
    const trs = tb.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
    const rows = trs.map((tr) => rowCells(tr, textOf));
    const aligns = trs.map((tr) => rowCells(tr, cellAlign));
    // 构件在表尾补了一行无边框定高空行（表与后续段落的间距），不是数据行
    if (rows.length && rows[rows.length - 1].every((c) => !c)) {
      rows.pop(); aligns.pop();
    }
    if (rows.length < 2) continue;
    rows.idx = idx;                        // 全文表格里的序号（修复工具挂批注用）
    rows.declared = /<w:tblHeader/.test(tb);
    tables.push(rows);
    tableAligns.push(aligns);
  }

  return { paras, tables, tableAligns, xml };
}

// 抠掉成对引号/书名号内的内容，只留一个占位空格，供文本类规则使用。
// 书名号里是文件名/书名，里面出现「二〇三五年」「筹画」是引用，不是错用。
const unquote = (s) => s.replace(/「[^」]*」/g, "　").replace(/“[^”]*”/g, "　")
  .replace(/《[^》]*》/g, "　");

// 数值格 vs 叙述格：含阿拉伯数字、且不含中文句读。判据同 check-num 的观察面。
const looksNumeric = (s) => /\d/.test(s) && !/[，、；。]/.test(s);

// ── 规则 ─────────────────────────────────────────────────
function check(file, { paras, tables, tableAligns, xml }) {
  const raw = paras.map((p) => p.text).join("\n");
  const text = unquote(raw);
  const hits = [];
  const hints = [];
  // anchors：命中处的原文片段，修复工具（tools/fix-doc.js）按它把批注挂到对应段落
  const add = (rule, detail, anchors = []) => hits.push({ rule, detail, anchors, blocking: true });
  const hint = (rule, detail, anchors = []) => hints.push({ rule, detail, anchors, blocking: false });

  // 0 有没有正文。放在所有规则之前——内容全丢的稿子后面每条都查不到东西，
  //   于是逐条通过、打印「通过」。这种坏法出现过：序列化坏掉、内容没了、
  //   机检说没问题。产出还在、内容没了，是最危险的一种坏。
  if (!paras.length) {
    add("文档无正文", "一个有文字的段落都没有。稿子是空的，或者序列化坏了——先看 word/document.xml 里 w:body 的内容");
    hits.hints = hints;
    return hits;
  }
  if (/<rootKey>/.test(xml)) {
    add("段落序列化坏了", "document.xml 里出现 <rootKey> 标签。多半是同一进程里加载了两份 docx 模块");
  }

  // 1 自称 / 称对方：企业材料里常见，属于口径偏好，不是国标硬性规定，只提示。
  if (/我司/.test(text)) hint("自称", "出现「我司」。一般用「我方」；非常正式、以机关名义署名的文件用「我单位」", ["我司"]);
  if (/贵司/.test(text)) hint("称对方", "出现「贵司」，改成「贵单位」", ["贵司"]);

  // 2 数字与单位之间缺空格（只认常见物理单位，避免误伤「3 个」「5 年」）
  //   ℃、%、°、′、″ 按 GB 3101 与数值之间不留空，不在此列。
  const unitRe = /\d(?:mAh\/g|GWh|kWh|MWh|kg|kW|MW|mm|cm|km|ppm|t\/h|g\/cm|Wh)/g;
  const noSpace = text.match(unitRe);
  if (noSpace) add("数字与单位", `${[...new Set(noSpace)].slice(0, 4).join("、")} 等处缺半角空格`, noSpace);

  // 3 数值范围省略后位单位（GB/T 15835 5.1.3）
  const rangeBad = text.match(/\d+(?:\.\d+)?\s*[~～—–－-]\s*\d+(?:\.\d+)?\s*(?:%|万元|亿元|万|亿)/g) || [];
  rangeBad.forEach((m) => {
    const head2 = m.split(/[~～—–－-]/)[0];
    if (!/%|万|亿|元/.test(head2)) {
      add("数值范围", `「${m}」前一个数缺单位，应写成 15%~30%、13 万元~17 万元`, [m]);
    }
  });

  // 4 纯小数缺整数位 0（GB/T 15835 5.1.2）
  const bareDec = text.match(/[^\d\w.]\.\d+/g);
  if (bareDec) add("纯小数", `${bareDec.slice(0, 3).join("、")} 缺小数点前的 0`, bareDec.map((x) => x.slice(1)));

  // 5 四位年份简写（GB/T 15835 5.1.4）。命中条件是「两位数字 + 年 + 时点词」——
  //   只有跟着时点词，那个「年」才一定是年份而不是年限（「运营 20 年」是时长）。
  const sy = [...new Set(text.match(/(?<![第满近共达约]\s?)(?<!\d)\d{2}\s*年(?=初|中|底|末|上半年|下半年)/g) || [])];
  // 日期里的两位年份：「26年10月」「26年10月23日」。带「月」才判，避免误伤
  // 「运营 20 年」这类年限。
  const syDate = [...new Set(text.match(/(?<!\d)\d{2}\s*年\s*\d{1,2}\s*月(?:\s*\d{1,2}\s*日)?/g) || [])]
    .filter((m) => !/^\d{4}/.test(m));
  if (sy.length || syDate.length) {
    add("年份简写", `${[...sy, ...syDate].slice(0, 3).join("、")}，四位年份不简写为两位`, [...sy, ...syDate]);
  }

  // 6 中文里的半角括号（GB/T 15834）
  const halfParen = text.match(/[一-龥][()]|[()][一-龥]/g);
  if (halfParen) add("半角括号", `${halfParen.slice(0, 3).join("、")}，中文用全角（）`, halfParen);

  // 7 斜体（GB/T 15835 5.1.7.1 数字用正体；衬线中文没有真正的意大利体，
  //   Word 会做倾斜变形，印出来是歪的）。中文正文用正体，但斜体也可能是
  //   英文术语/变量名，属提示不拦稿。格式修订（w:rPrChange）里存的是改之前
  //   的格式，不算。
  if (/<w:i\/>|<w:i w:val="(?:1|true|on)"/.test(xml.replace(/<w:rPrChange\b[\s\S]*?<\/w:rPrChange>/g, ""))) {
    hint("斜体", "文中有斜体 run。中文正文一般用正体，强调用加粗；英文变量名等语境请自行判断");
  }

  // 8 全角数字（GB/T 15835 5.1.7.1 阿拉伯数字用二分字身）
  const fullNum = text.match(/[０-９]+/g);
  if (fullNum) add("全角数字", `${[...new Set(fullNum)].slice(0, 3).join("、")}，阿拉伯数字用半角`, fullNum);

  // 9 阿拉伯数字与汉字数位混用（GB/T 15835 5.3）。只收无歧义的一档：
  //   数字后接「十／百／千」再接「万／亿／零」，宽形会误伤「2018十大」「5千瓦」。
  const mixNum = text.match(/\d+\s*[十百千]\s*[万亿零]/g);
  if (mixNum) add("数位混用", `${[...new Set(mixNum)].slice(0, 3).join("、")}，写成 300 万、100 零 8`, mixNum);

  // 10 年月日简写不完整（GB/T 15835 5.1.4：年月日不完整时不能用「-」替代）。
  //    网址里的数字串不是日期，先剔掉 URL 再查。
  const badYmd = text.replace(/https?:\/\/\S+/g, "").match(/\d{4}-\d{1,2}(?!\d)(?!-)/g);
  if (badYmd) add("日期简写", `${[...new Set(badYmd)].slice(0, 3).join("、")}，年月日不全时写成 2008 年 8 月`, badYmd);

  // 11 年月日用顿号分隔（GB/T 15834 A.4.2：用短横线连接号，不用顿号）
  const dunYmd = text.match(/\d{4}\s*、\s*\d{1,2}\s*、\s*\d{1,2}/g);
  if (dunYmd) add("日期顿号", `${dunYmd.slice(0, 3).join("、")}，写成 2010-03-02`, dunYmd);

  // ── 表格内部（GB/T 1.1 10.4.2.2、GB/T 15834 A.1）────────
  tables.forEach((rows, ti) => {
    // 手写表（首行没声明 tblHeader）只提示：可能是跨页续表、只列前几名的表，
    // 也可能根本不是数据表；blocking 只收本工具生成/显式声明表头的表。
    const report = rows.declared ? add : hint;
    rows.forEach((cells, ri) => {
      cells.forEach((c, ci) => {
        const at = `第 ${ti + 1} 表第 ${ri + 1} 行第 ${ci + 1} 格`;
        if (c === null) return;   // 合并位/纵向合并续行不是格
        // 12 相同内容不写「同上」
        if (/^同[上左前]$/.test(c)) report("表格同上", `${at}「${c}」，写具体值或通栏`, [{ text: c, table: rows.idx }]);
        // 13 空单元格应写一字线（空串才是格空了；null 是合并位）
        if (ri > 0 && c === "") report("表格空格", `${at}是空的，没有内容的格写「—」`, [{ text: (cells[0] || ""), table: rows.idx }]);
        // 14 表内段后不加句号
        if (c.endsWith("。")) report("表格句号", `${at}「${c.slice(0, 16)}…」以句号收尾，表内不加句号`, [{ text: c, table: rows.idx }]);
        // 15 表头夹英文：连续两个以上大写字母即判为夹了英文词。GDP/CPI 这类
        //    标准缩写也会命中，属提示不拦稿；型号 XC-21 先抠掉再判。
        if (ri === 0 && /[A-Z]{2,}/.test(c.replace(/\b[A-Z]{1,5}-?\d[\w-]*/g, ""))) {
          hint("表头夹英文", `${at}「${c}」，表头一般用中文；GDP/CPI 等标准缩写可保留`, [{ text: c, table: rows.idx }]);
        }
      });
    });
  });

  // 16 数值列非数值：右对齐只该给数值列。机检不追参数，只认「右对齐的列里
  //    多数格不是数值」——下标错位、标错列在 docx 里长得一模一样。对齐位置
  //    只是启发式，属提示不拦稿。
  tables.forEach((rows, ti) => {
    if (rows.length < 2 || !rows[0]) return;
    for (let ci = 0; ci < rows[0].length; ci++) {
      if ((tableAligns[ti]?.[1]?.[ci]) !== "right") continue;
      const body = rows.slice(1).map((r) => r[ci] ?? "").filter(Boolean);
      if (!body.length) continue;
      const prose = body.filter((c) => !looksNumeric(c));
      if (prose.length * 2 > body.length) {
        hint("数值列非数值", `第 ${ti + 1} 表第 ${ci + 1} 列右对齐（数值列），但 ${prose.length}/${body.length} 格是叙述文字，如「${prose[0].slice(0, 16)}」`, [{ text: prose[0], table: rows.idx }]);
      }
    }
  });

  // 17 标题重复编号：h1～h4 自带「一、（一）1.（1）」编号，标题文字里再手写
  //    一套就叠成两个。小数（1.5）不算——点号后跟数字的排除。
  const dup = paras.filter((p) => /^Heading[1-4]$/.test(p.style) && p.numbered
    && /^(?:[一二三四五六七八九十]+[、．.]|（[一二三四五六七八九十]+）|\d{1,2}[.．、](?!\d))/.test(p.text));
  if (dup.length) {
    add("标题重复编号", `${dup.length} 个标题手写了编号，与自动编号叠成两个：「${dup[0].text.slice(0, 20)}」`, dup.map((p) => p.text));
  }

  // （原「自然段顶格」一条没有移植：私有编辑体例里正文段的缩进直接写在段落上，
  //   查 w:firstLine 查得到；公文体例的缩进由 Normal 样式统一给，段落上没有
  //   w:firstLine，照原判据会全量误报。缩进由样式底座保证，不需要这条。）

  // ── 公文体专属（GB/T 9704）────────────────────────────────
  // 19 层次序数。公文只有 一、／（一）／1.／（1） 四层，不用 1.1／1.1.1。
  //    表内不查（那是数值），后面跟计量单位的也不查（「1.2 亿元」不是编号）。
  {
    const UNIT = "亿元|万元|亿|万|元|吨|亩|人|户|个|年|月|日|%|km|kg|mm|m";
    const dotted = paras.filter((p) => !p.inTable
      && new RegExp(`^\\d+\\.\\d+(?!\\s*(?:${UNIT}))`).test(p.text));
    if (dotted.length) {
      add("公文层次序数", `${dotted.length} 段用了 1.1 式编号：「${dotted[0].text.slice(0, 24)}…」。公文用 一、（一）1.（1） 四层（GB/T 9704 7.3.3）`, dotted.map((p) => p.text));
    }
  }

  // 20 中文数字日期。公文成文日期用阿拉伯数字（GB/T 9704 7.3.5.3）。
  //    只认「〇」这个字符——它在中文里专用于年份写法，形状干净，不会误伤。
  {
    const linZero = (text.match(/[一二三四五六七八九〇]*〇[一二三四五六七八九〇]*\s*年/g) || []);
    if (linZero.length) {
      add("中文数字日期", `${[...new Set(linZero)].slice(0, 3).join("、")}，成文日期与正文日期一律阿拉伯数字（GB/T 9704 7.3.5.3）`, linZero);
    }
  }

  // 21 月日编虚位（GB/T 9704 7.3.5.4）：月、日不编虚位（即 1 不编为 01）。
  {
    const padded = text.match(/年\s*0\d\s*月|月\s*0\d\s*日/g) || [];
    if (padded.length) {
      add("月日编虚位", `${[...new Set(padded)].slice(0, 3).join("、")}，月、日不编虚位（1 不写 01）`, padded);
    }
  }

  // 21b 成文日期不完整：单独成段的「2026年10月」缺日，不是完整日期。
  //     只认整段就是年月，避免误伤正文里「2026年10月，……」这类叙述。
  {
    const monthOnly = paras
      .filter((p) => !p.inTable && /^\s*\d{4}\s*年\s*\d{1,2}\s*月\s*$/.test(p.text))
      .map((p) => p.text.trim());
    if (monthOnly.length) {
      add("成文日期", `${monthOnly.slice(0, 3).join("、")} 缺日；成文日期应年、月、日标全（GB/T 9704 7.3.5.3）`, monthOnly);
    }
  }

  // 22 发文字号（GB/T 9704 7.2.5）：年份应标全称、用六角括号〔〕括入；
  //    发文顺序号不加「第」字，不编虚位，在阿拉伯数字后加「号」字。
  {
    const fnBad = [];
    const wrongBracket = text.match(/【\d{1,4}】|\[\d{4}\]/g) || [];
    if (wrongBracket.length) fnBad.push(`年份用六角括号〔〕：${[...new Set(wrongBracket)].join("、")}`);
    const shortYear = text.match(/〔\d{1,3}〕(?=\s*(?:第\s*)?\d+\s*号)/g) || [];
    if (shortYear.length) fnBad.push(`年份应标全称：${[...new Set(shortYear)].join("、")}`);
    const withDi = text.match(/〔\d{4}〕\s*第\s*\d+\s*号/g) || [];
    if (withDi.length) fnBad.push(`顺序号不加「第」字：${[...new Set(withDi)].join("、")}`);
    const withZero = text.match(/〔\d{4}〕\s*0+\d+\s*号/g) || [];
    if (withZero.length) fnBad.push(`顺序号不编虚位：${[...new Set(withZero)].join("、")}`);
    if (fnBad.length) add("发文字号", fnBad.join("；"), [...wrongBracket, ...shortYear, ...withDi, ...withZero]);
  }

  // 23 版式。版心与行字数是国标硬指标，被人在 Word 里改过就不再是公文版式。
  //    每一节都查：修复工具会给所有节改版式，只查第一节会放过后面的节。
  //    sectPrChange 里存的是旧版式，不是当前版式，先摘掉。
  {
    const clean = xml.replace(/<w:sectPrChange\b[\s\S]*?<\/w:sectPrChange>/g, "");
    const sects = clean.match(/<w:sectPr(?:\s[^>]*)?>[\s\S]*?<\/w:sectPr>/g) || [];
    const want = { top: "2098", bottom: "1984", left: "1587", right: "1474" };
    const where = (i) => (sects.length > 1 ? `第 ${i + 1} 节` : "");
    if (!sects.length) {
      add("公文版式", "没找到 sectPr，版心未设（GB/T 9704 8.1：上 37 下 35 左 28 右 26 mm）");
    }
    sects.forEach((sect, i) => {
      const mar = sect.match(/<w:pgMar[^>]*>/);
      if (!mar) {
        add("公文版式", `${where(i)}没找到 pgMar，版心未设（GB/T 9704 8.1：上 37 下 35 左 28 右 26 mm）`);
      } else {
        // ±2 twips 容差：毫米换 twips 取整方式不同，差一两个不算问题。
        const bad = Object.entries(want).filter(([k, v]) => {
          const got = (mar[0].match(new RegExp(`w:${k}="(-?\\d+)"`)) || [])[1];
          return got === undefined || Math.abs(Number(got) - Number(v)) > 2;
        });
        if (bad.length) {
          add("公文版式", `${where(i)}版心 ${bad.map(([k]) => k).join("、")} 不合 GB/T 9704 8.1（上 37 下 35 左 28 右 26 mm）`);
        }
      }
      if (!/<w:docGrid[^>]*w:linePitch="579"/.test(sect)) {
        add("公文版式", `${where(i)}docGrid linePitch 不是 579：每页 22 行的行距被改过（GB/T 9704 8.1）`);
      }
    });
  }

  hits.hints = hints;
  return hits;
}

module.exports = { readDocx, check };

// ── 入口 ─────────────────────────────────────────────────
if (require.main === module) (async () => {
  const file = process.argv[2];
  if (!file) {
    console.error("用法：node check-doc.js <文件.docx> [--strict] [--json]");
    process.exit(2);
  }
  const strict = process.argv.includes("--strict");
  let doc;
  try {
    doc = await readDocx(file);
  } catch (e) {
    console.error("机检读不进这份文件（工具故障，不是稿子的问题）：", e.message);
    process.exit(2);
  }
  const hits = check(file, doc);
  const hints = hits.hints || [];
  if (process.argv.includes("--json")) {
    console.log(JSON.stringify({
      check: "体例", file: path.basename(file),
      hits: [...hits, ...hints].map((h) => ({ rule: h.rule, detail: h.detail, blocking: h.blocking !== false, anchors: h.anchors || [] })),
    }));
    // 不调 process.exit：stdout 是管道时，输出较大会在刷完之前被截断
    if (hits.length && strict) process.exitCode = 1;
    return;
  }
  if (!hits.length) {
    console.log(`体例机检通过　${path.basename(file)}`);
  } else {
    console.log(`体例机检 ${hits.length} 处　${path.basename(file)}`);
    hits.forEach((h) => console.log(`  · ${h.rule}：${h.detail}`));
    if (strict) process.exit(1);
  }
  if (hints.length) {
    console.log(`  提示 ${hints.length} 处（不拦稿，请人工核对）：`);
    hints.forEach((h) => console.log(`    · ${h.rule}：${h.detail}`));
  }
})();
