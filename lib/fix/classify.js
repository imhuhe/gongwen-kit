"use strict";
/**
 * classify.js —— 修复工具的两类判定：版头发文字号 vs 条文体（2026-10-23）。
 *
 * 这两类判定原先在 layout.js 与 fix-doc.js 各写一份，文本口径不同：
 * layout 的 p.text 只拼 <w:t>，不含 <w:tab/>，于是「字号 + 制表位 + 签发人」
 * 在文本修订侧能保护、在版式侧却保护不了。这里只留一份，两边共用。
 *
 * 发文字号保护只在版头（前 12 段）且整段是字号，或「字号 + 制表位 + 签发人」；
 * 正文「根据国办发〔2026〕5号要求」不能整段跳过。
 *
 * 条文体判定：正文同时出现「第X章」「第X条」；或 Heading1 实际挂的那一级
 * lvlText 是「第%N章」且同一编号定义里还有「第%N条」。不能只在整个
 * numbering.xml 里找「章」或「条」——不分章的规定/细则只有「第X条」，
 * 没被段落用到的 abstractNum 也不该触发。
 */

const DOCNO_ONLY = /^\s*[^，。；：]{0,20}〔\d{4}〕\s*第?\s*\d+\s*号\s*$/;
const DOCNO_TAB = /^\s*[^，。；：]{0,20}〔\d{4}〕\s*第?\s*\d+\s*号\s+(?:签发人[：:])?[^，。；：\s]{1,20}(?:\s+[^，。；：\s]{1,20})*\s*$/;

/** text 应已把 <w:tab/> 还原成空格；hasTab 表示原段落里确有制表位。 */
function isDocNoLine(text, hasTab) {
  return DOCNO_ONLY.test(text) || (hasTab && DOCNO_TAB.test(text));
}

// numId -> abstractNumId；abstractNumId -> Map(ilvl -> lvlText)
function numberingLevels(numbering) {
  const nums = new Map();
  for (const m of (numbering || "").matchAll(/<w:num w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g)) {
    const ab = (m[2].match(/<w:abstractNumId w:val="(\d+)"/) || [])[1];
    if (ab) nums.set(m[1], ab);
  }
  const abs = new Map();
  for (const m of (numbering || "").matchAll(/<w:abstractNum w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g)) {
    const levels = new Map();
    for (const l of m[2].matchAll(/<w:lvl w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/g)) {
      const t = (l[2].match(/<w:lvlText w:val="([^"]*)"/) || [])[1];
      if (t != null) levels.set(l[1], t);
    }
    abs.set(m[1], levels);
  }
  return { nums, abs };
}

function numPrOf(block) {
  const m = (block || "").match(/<w:numPr>([\s\S]*?)<\/w:numPr>/);
  if (!m) return null;
  const numId = (m[1].match(/<w:numId w:val="(\d+)"/) || [])[1];
  const ilvl = (m[1].match(/<w:ilvl w:val="(\d+)"/) || [])[1] || "0";
  return numId ? { numId, ilvl } : null;
}

/** Heading1 的 numPr：样式优先，其次带 Heading1 样式的段落；兼容中文样式 ID。 */
function headingNumPr(styles, document) {
  let styleId = "";
  let styleBody = "";
  for (const m of (styles || "").matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    if (!/w:type="paragraph"/.test(m[1])) continue;
    const name = ((m[2].match(/<w:name w:val="([^"]+)"/) || [])[1] || "").toLowerCase();
    if (name === "heading 1") {
      styleId = (m[1].match(/w:styleId="([^"]+)"/) || [])[1] || "";
      styleBody = m[2];
      break;
    }
  }
  if (!styleId) styleId = "Heading1";
  // 段落上的 numPr 才是「实际挂的级别」；样式里的 numPr 可能指向另一套编号。
  for (const m of (document || "").matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g)) {
    if (new RegExp(`<w:pStyle w:val="${styleId.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`).test(m[0])) {
      const n = numPrOf(m[0]);
      if (n) return n;
    }
  }
  if (styleBody) {
    const n = numPrOf(styleBody);
    if (n) return n;
  }
  return null;
}

/**
 * text：document.xml 抽出的正文文字；styles/document/numbering 是对应部件。
 * 正文同时有章、条；或 Heading1 实际级别的 lvlText 是「第%N章」且同一
 * 编号定义里还有「第%N条」→ 条文体。
 */
function isZhangDoc({ text = "", styles = "", document = "", numbering = "" } = {}) {
  if (/第[一二三四五六七八九十百]+章/.test(text) && /第[一二三四五六七八九十百]+条/.test(text)) return true;
  const n = headingNumPr(styles, document);
  if (!n) return false;
  const { nums, abs } = numberingLevels(numbering);
  const abstractId = nums.get(n.numId);
  if (!abstractId) return false;
  const levels = abs.get(abstractId);
  if (!levels) return false;
  const current = levels.get(n.ilvl) || "";
  if (!/^第%\d+章/.test(current)) return false;
  return [...levels.values()].some((t) => /^第%\d+条/.test(t));
}

module.exports = { DOCNO_ONLY, DOCNO_TAB, isDocNoLine, isZhangDoc, numberingLevels, headingNumPr };
