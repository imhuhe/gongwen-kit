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

const DATA = path.join(__dirname, "..", "data");
const CJK = "\\u4e00-\\u9fff\\u3000-\\u303f";
const GROUP_NAME = { ai: "AI 套话", biz: "商业黑话", empty: "空词与对冲", perform: "表演腔" };

function loadTsv(name, cols) {
  const rows = [];
  let text;
  try {
    text = fs.readFileSync(path.join(DATA, name), "utf8");
  } catch {
    return rows;   // 词表缺失只让这一项失效，不崩整个机检
  }
  for (const ln of text.split("\n")) {
    if (!ln.trim() || ln.trimStart().startsWith("#")) continue;
    const p = ln.split("\t").map((s) => s.trim());
    if (p.length < cols || p[cols - 1] !== "strict") continue;
    rows.push(p);
  }
  return rows;
}

// 标准新增三类（判据本体在私有体系的《AI 输出执行标准》，这里只实现）：
const LATIN_ABBR = /\be\.g\.|\bi\.e\.|\betc\.|[（(]?et al\./gi;
const MODAL_STACK = /(?:可能|也许|或许|大概|似乎|恐怕|应该|应当)[、]?(?:可能|也许|或许|大概|似乎|恐怕|应该|应当)/g;
const NOMINALIZE = /(?:进行|予以|加以)(?:了|一个|一场|一次|一番)?[^，。；：！？、,.:;!?()（）「」\n]{0,8}?(?:调整|修改|优化|处理|检查|确认|评估|分析|核实|更新|变更|操作|梳理|排查|整理|验证|说明|补充|升级|改造|讨论|盘点|清理|迁移|部署|发布|采集|统计|核算)/g;

/** 扫一段已抠引文的正文，返回 { norm, style, anchors, normRules, styleRules }：
 *  norm、style 是两组提示行；normRules、styleRules 与它们逐项对应，给出各行的规则名
 *  （供 --json 使用，不要从提示行文本里反切）。 */
function scan(body) {
  const norm = [];
  const normRules = [];
  const styleRules = [];
  const anchors = [];   // 规范层命中的原文片段（修复工具挂批注用）
  const style = [];

  // ── 规范层 ─────────────────────────────────────────────
  const yx = [];
  for (const row of loadTsv("yixingci.tsv", 4)) {
    for (const bad of row[1].split("｜")) {
      const w = bad.trim();
      if (w && body.includes(w)) { yx.push(`${w} → ${row[0]}`); anchors.push(w); }
    }
  }
  if (yx.length) { norm.push(`异形词 ${yx.length} 处（GF 1001-2001）：${yx.join("、")}`); normRules.push("异形词"); }

  // 半角标点：只算紧贴中文的那些。例外一档——阿拉伯数字序号的半角下脚点
  // （「附件：1.××××」的「1.」）是 GB/T 15834 附录 B 的正当写法，
  // 不能因为后面紧跟中文就报（我们的附件说明构件自己就这么排版）。
  const halfRe = new RegExp(`[${CJK}][,.!?;:()]|[,.!?;:()][${CJK}]`, "g");
  const half = [];
  let hm;
  while ((hm = halfRe.exec(body))) {
    if (hm[0][0] === "." && hm.index > 0 && /\d/.test(body[hm.index - 1])) continue;
    half.push(hm[0]);
    anchors.push(hm[0]);
  }
  if (half.length) {
    const seen = [...new Set(half.map((h) => h.trim()))];
    norm.push(`中文里的半角标点 ${half.length} 处（${seen.slice(0, 6).join("、")}）。一律改全角（GB/T 15834）。`);
    normRules.push("中文里的半角标点");
  }

  // ── 风格层：词表 ───────────────────────────────────────
  const byGroup = {};
  for (const row of loadTsv("style-words.tsv", 4)) {
    const [bad, good, group] = row;
    if (body.includes(bad)) (byGroup[group] = byGroup[group] || []).push(`${bad} → ${good}`);
  }
  for (const [g, items] of Object.entries(byGroup)) {
    style.push(`${GROUP_NAME[g] || g}：${items.join("；")}`);
    styleRules.push(GROUP_NAME[g] || g);
  }

  // ── 风格层：形状 ───────────────────────────────────────
  // 否定平行句：句型本身正当（定义＋区分），病在滥用——同一篇 ≥3 次才报。
  const neg = body.match(/不是[^。！？\n]{0,30}而是|并非[^。！？\n]{0,30}而是/g) || [];
  if (neg.length >= 3) {
    style.push(`否定平行句 ${neg.length} 处（不是…而是／并非…而是）。句型正当，短段里连着用是审美疲劳，留真正需要区分的那处。`);
    styleRules.push("否定平行句");
  }

  // AI 工具指纹：从别处粘来时留下的标记，出现即删。
  const fp = body.match(/turn\d+search\d+|oai_citation|oaicite|contentReference|utm_source=(?:chatgpt|copilot)\.com/g) || [];
  if (fp.length) { style.push(`AI 工具指纹 ${fp.length} 处：${[...new Set(fp)].slice(0, 4).join("、")}。整段删掉。`); styleRules.push("AI 工具指纹"); }

  const latin = body.match(LATIN_ABBR) || [];
  if (latin.length) {
    style.push(`拉丁缩写 ${latin.length} 处（${[...new Set(latin.map((x) => x.toLowerCase()))].join("、")}）：中文文本写「例如／即／等等」。`);
    styleRules.push("拉丁缩写");
  }
  const modal = body.match(MODAL_STACK) || [];
  if (modal.length) { style.push(`情态词堆叠 ${modal.length} 处（${[...new Set(modal)].join("、")}）：留一个。`); styleRules.push("情态词堆叠"); }
  const nm = body.match(NOMINALIZE) || [];
  if (nm.length) { style.push(`虚动词结构 ${nm.length} 处（${nm.slice(0, 4).join("、")}）：动作直接用动词。`); styleRules.push("虚动词结构"); }

  return { norm, style, anchors, normRules, styleRules };
}

module.exports = { scan, loadTsv };
