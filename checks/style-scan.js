"use strict";
/**
 * style-scan.js —— 文风扫描器（出稿链用）。
 *
 * 两层判据，分法是「这个检查能不能写成命中即错，不看语境」：
 *   规范层——依据国标，命中即错：异形词（GF 1001-2001）、中文里的半角标点（GB/T 15834）。
 *   风格层——依据偏好，只报不判：词表（AI 套话、商业黑话、空词、表演腔）
 *            与几种形状（否定平行句成堆、AI 工具指纹、拉丁缩写、情态堆叠、虚动词结构）。
 *
 * 词表在 data/ 下，每次调用读盘：加一行下次生效，不用改这个文件。
 * 负面词表**故意不进提示词**（进了模型就会写「本段未使用某某」来呼应它）。
 * 传入的正文须已抠掉引文——「」“”里是被讨论的词，不算在用。
 */
const fs = require("fs");
const path = require("path");
const { protectedRanges, inRanges } = require("../lib/text-guard");

const DATA = path.join(__dirname, "..", "data");
const CJK = "\\u4e00-\\u9fff\\u3000-\\u303f";
const GROUP_NAME = { ai: "AI 套话", biz: "商业黑话", empty: "空词与对冲", perform: "表演腔" };

function loadTsv(name, cols, { required = false } = {}) {
  const rows = [];
  let text;
  try {
    text = fs.readFileSync(path.join(DATA, name), "utf8");
  } catch (e) {
    // 风格词表缺失只让这一项失效；规范层词表（异形词）缺失是工具故障，
    // 不能静默放行——否则 --json 会给出一个「文风通过」的假阴性。
    if (required) throw new Error(`词表缺失：data/${name}（${e.code || e.message}）`);
    return rows;
  }
  for (const ln of text.split("\n")) {
    if (!ln.trim() || ln.trimStart().startsWith("#")) continue;
    const p = ln.split("\t").map((s) => s.trim());
    if (p.length < cols || p[cols - 1] !== "strict") continue;
    rows.push(p);
  }
  return rows;
}

// 某个词是否在保护范围之外出现（URL/引号/书名号/文件名/域名里的不算）。
function hasUnprotected(body, word, guard) {
  let i = body.indexOf(word);
  while (i >= 0) {
    if (!inRanges(guard, i)) return true;
    i = body.indexOf(word, i + word.length);
  }
  return false;
}

// 标准新增三类（判据本体在私有体系的《AI 输出执行标准》，这里只实现）：
const LATIN_ABBR = /\be\.g\.|\bi\.e\.|\betc\.|[（(]?et al\./gi;
const MODAL_STACK = /(?:可能|也许|或许|大概|似乎|恐怕|应该|应当)[、]?(?:可能|也许|或许|大概|似乎|恐怕|应该|应当)/g;
const NOMINALIZE = /(?:进行|予以|加以)(?:了|一个|一场|一次|一番)?[^，。；：！？、,.:;!?()（）「」\n]{0,8}?(?:调整|修改|优化|处理|检查|确认|评估|分析|核实|更新|变更|操作|梳理|排查|整理|验证|说明|补充|升级|改造|讨论|盘点|清理|迁移|部署|发布|采集|统计|核算)/g;

/** 扫一段已抠引文的正文，返回 { norm, style, anchors, normHits, styleHits }：
 *  normHits、styleHits 是 {rule, text} 数组，规则名与提示行出自同一次 push，不会错位；
 *  norm、style 是其中的 text（两组提示行，旧调用方用）。 */
function scan(body) {
  const normHits = [];
  const styleHits = [];
  const normHit = (rule, text) => normHits.push({ rule, text });
  const styleHit = (rule, text) => styleHits.push({ rule, text });
  const anchors = [];   // 规范层命中的原文片段（修复工具挂批注用）
  const guard = protectedRanges(body);

  // ── 规范层 ─────────────────────────────────────────────
  const yx = [];
  for (const row of loadTsv("yixingci.tsv", 4, { required: true })) {
    for (const bad of row[1].split("｜")) {
      const w = bad.trim();
      if (w && hasUnprotected(body, w, guard)) { yx.push(`${w} → ${row[0]}`); anchors.push(w); }
    }
  }
  if (yx.length) { normHit("异形词", `异形词 ${yx.length} 处（GF 1001-2001）：${yx.join("、")}`); }

  // 半角标点：只算紧贴中文的那些。例外一档——阿拉伯数字序号的半角下脚点
  // （「附件：1.××××」的「1.」）是 GB/T 15834 附录 B 的正当写法，
  // 不能因为后面紧跟中文就报（我们的附件说明构件自己就这么排版）。
  // 文件名/域名/网址/引号/书名号里的标点不算（2026-10-23 审计）。
  const halfRe = new RegExp(`[${CJK}][,.!?;:()]|[,.!?;:()][${CJK}]`, "g");
  const half = [];
  let hm;
  while ((hm = halfRe.exec(body))) {
    if (hm[0][0] === "." && hm.index > 0 && /\d/.test(body[hm.index - 1])) continue;
    const punct = hm[0].search(/[,.!?;:()]/);
    if (inRanges(guard, hm.index + punct)) continue;
    half.push(hm[0]);
    anchors.push(hm[0]);
  }
  if (half.length) {
    const seen = [...new Set(half.map((h) => h.trim()))];
    normHit("中文里的半角标点", `中文里的半角标点 ${half.length} 处（${seen.slice(0, 6).join("、")}）。一律改全角（GB/T 15834）。`);
  }

  // ── 风格层：词表 ───────────────────────────────────────
  const byGroup = {};
  for (const row of loadTsv("style-words.tsv", 4)) {
    const [bad, good, group] = row;
    if (hasUnprotected(body, bad, guard)) (byGroup[group] = byGroup[group] || []).push(`${bad} → ${good}`);
  }
  for (const [g, items] of Object.entries(byGroup)) {
    styleHit(GROUP_NAME[g] || g, `${GROUP_NAME[g] || g}：${items.join("；")}`);
  }

  // ── 风格层：形状 ───────────────────────────────────────
  // 否定平行句：句型本身正当（定义＋区分），病在滥用——同一篇 ≥3 次才报。
  const neg = body.match(/不是[^。！？\n]{0,30}而是|并非[^。！？\n]{0,30}而是/g) || [];
  if (neg.length >= 3) {
    styleHit("否定平行句", `否定平行句 ${neg.length} 处（不是…而是／并非…而是）。句型正当，短段里连着用是审美疲劳，留真正需要区分的那处。`);
  }

  // AI 工具指纹：从别处粘来时留下的标记，出现即删。
  const fp = body.match(/turn\d+search\d+|oai_citation|oaicite|contentReference|utm_source=(?:chatgpt|copilot)\.com/g) || [];
  if (fp.length) { styleHit("AI 工具指纹", `AI 工具指纹 ${fp.length} 处：${[...new Set(fp)].slice(0, 4).join("、")}。整段删掉。`); }

  const latin = body.match(LATIN_ABBR) || [];
  if (latin.length) {
    styleHit("拉丁缩写", `拉丁缩写 ${latin.length} 处（${[...new Set(latin.map((x) => x.toLowerCase()))].join("、")}）：中文文本写「例如／即／等等」。`);
  }
  const modal = body.match(MODAL_STACK) || [];
  if (modal.length) { styleHit("情态词堆叠", `情态词堆叠 ${modal.length} 处（${[...new Set(modal)].join("、")}）：留一个。`); }
  const nm = body.match(NOMINALIZE) || [];
  if (nm.length) { styleHit("虚动词结构", `虚动词结构 ${nm.length} 处（${nm.slice(0, 4).join("、")}）：动作直接用动词。`); }

  return { norm: normHits.map((h) => h.text), style: styleHits.map((h) => h.text), anchors, normHits, styleHits };
}

module.exports = { scan, loadTsv };
