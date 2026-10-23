"use strict";
/**
 * semantic.js —— 版式修复前后的「实际效果」比对（2026-10-23）。
 *
 * 目的：自家出稿跑 fix 时，不再靠「找有没有已知的改乱痕迹」来决定跳过，
 * 而是照常算一遍版式结果，和输入比 Normal/Title/Heading 的有效格式、正文
 * run 的直接格式、节属性与页脚内容。完全相同才保留原件并报「版式（未改）」。
 *
 * 只比显示效果，不比 XML 写法：cstheme、szCs、rsid、footer 的 rId 名字、
 * w:orient 这些不影响呈现的差异会被忽略；字体大小按中文 w:sz / docDefaults
 * 计算，颜色、加粗、对齐、缩进、段距、编号、页边距、页脚 PAGE 域都参与比较。
 */
const KEY_STYLES = ["Normal", "Title", "Heading1", "Heading2", "Heading3", "Heading4",
  "TitleChar", "Heading1Char", "Heading2Char", "Heading3Char", "Heading4Char"];

const ANCHOR = {};

function attr(xml, name) {
  return (xml.match(new RegExp(`w:${name}="([^"]*)"`)) || [])[1] || "";
}

function attrs(xml, tag, names) {
  const m = (xml.match(new RegExp(`<${tag}\\b[^>]*>`)) || [])[0] || "";
  const out = {};
  for (const n of names) out[n] = attr(m, n);
  return out;
}

function lastDocDefaults(styles) {
  const blocks = [...(styles || "").matchAll(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/g)];
  return blocks.length ? blocks[blocks.length - 1][0] : "";
}

function rPrOf(block) {
  return ((block || "").match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0];
}
function pPrOf(block) {
  return ((block || "").match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [""])[0];
}

function fontSig(rpr, defaults) {
  const rf = (rpr.match(/<w:rFonts\b[^>]*>/) || [""])[0];
  const df = (defaults.match(/<w:rFonts\b[^>]*>/) || [""])[0];
  return ["ascii", "eastAsia", "hAnsi"].map((n) => attr(rf, n) || attr(df, n) || "");
}

function szSig(rpr, defaults, fallback = 32) {
  return Number(attr((rpr.match(/<w:sz\b[^>]*>/) || [""])[0], "val")
    || attr((defaults.match(/<w:sz\b[^>]*>/) || [""])[0], "val") || fallback);
}

function rPrSig(rpr, defaults) {
  const b = /<w:b(?:\s+w:val="(?:1|true)")?\s*\/>/.test(rpr) && !/<w:b w:val="(?:0|false)"\s*\/>/.test(rpr);
  const color = attr((rpr.match(/<w:color\b[^>]*>/) || [""])[0], "val");
  return { font: fontSig(rpr, defaults), sz: szSig(rpr, defaults), b, color };
}

function pPrSig(ppr) {
  const jc = attr((ppr.match(/<w:jc\b[^>]*>/) || [""])[0], "val");
  const ind = attrs(ppr, "w:ind", ["firstLine", "firstLineChars", "left", "right", "leftChars", "rightChars", "hanging", "hangingChars"]);
  const spacing = attrs(ppr, "w:spacing", ["before", "after", "line", "lineRule", "beforeLines", "afterLines"]);
  const numPr = ppr.match(/<w:numPr>[\s\S]*?<\/w:numPr>/);
  const num = numPr ? {
    ilvl: attr((numPr[0].match(/<w:ilvl\b[^>]*>/) || [""])[0], "val"),
    numId: attr((numPr[0].match(/<w:numId\b[^>]*>/) || [""])[0], "val"),
  } : null;
  return {
    jc, ind, spacing, num,
    keepNext: /<w:keepNext\b/.test(ppr),
    outlineLvl: attr((ppr.match(/<w:outlineLvl\b[^>]*>/) || [""])[0], "val"),
    widowControl: attr((ppr.match(/<w:widowControl\b[^>]*>/) || [""])[0], "val"),
  };
}

function styleMap(styles) {
  const map = new Map();
  for (const m of (styles || "").matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    if (!/w:type="paragraph"/.test(m[1]) && !/w:type="character"/.test(m[1])) continue;
    const id = (m[1].match(/w:styleId="([^"]+)"/) || [])[1];
    if (id) map.set(id, m[2]);
  }
  return map;
}

function stylesSig(styles) {
  const defaults = rPrOf(lastDocDefaults(styles));
  const map = styleMap(styles);
  const out = { default: rPrSig(defaults, defaults), styles: {} };
  for (const id of KEY_STYLES) {
    const block = map.get(id);
    out.styles[id] = block ? {
      rPr: rPrSig(rPrOf(block), defaults),
      pPr: pPrSig(pPrOf(block)),
    } : null;
  }
  return out;
}

function bodySig(doc) {
  const out = [];
  for (const m of (doc || "").matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g)) {
    const p = m[0];
    const runs = [...p.matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)]
      .map((r) => {
        const rpr = (r[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0];
        return rpr ? rPrSig(rpr, "") : null;
      })
      .filter(Boolean);
    out.push({ pPr: pPrSig(pPrOf(p)), runs });
  }
  return out;
}

function relsMap(rels) {
  const map = new Map();
  for (const m of (rels || "").matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], "Id");
    const target = attr(m[0], "Target");
    if (id && target) map.set(id, target);
  }
  return map;
}

async function footerSig(zip, target) {
  const f = zip.file(`word/${target.replace(/^\.?\//, "")}`);
  if (!f) return null;
  const xml = await f.async("string");
  const texts = [];
  for (const m of xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>|<w:fldChar\b[^>]*w:fldCharType="([^"]+)"/g)) {
    texts.push(m[1] != null ? m[1] : m[2] != null ? `FIELD:${m[2]}` : `CHAR:${m[3]}`);
  }
  const jc = attr((xml.match(/<w:jc\b[^>]*>/) || [""])[0], "val");
  const ind = attrs(xml, "w:ind", ["left", "right", "firstLine", "firstLineChars"]);
  const font = fontSig(rPrOf(xml), "");
  const sz = szSig(rPrOf(xml), "", 0);
  return { texts, jc, ind, font, sz };
}

async function sectionSig(zip, doc) {
  const rels = relsMap(await (zip.file("word/_rels/document.xml.rels") ? zip.file("word/_rels/document.xml.rels").async("string") : ""));
  const out = [];
  for (const m of (doc || "").matchAll(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)) {
    const s = m[0];
    const ftrs = [];
    for (const f of s.matchAll(/<w:footerReference\b[^>]*>/g)) {
      const type = attr(f[0], "type");
      const id = (f[0].match(/r:id="([^"]+)"/) || [])[1];
      const target = rels.get(id);
      ftrs.push({ type, footer: target ? await footerSig(zip, target) : null });
    }
    ftrs.sort((a, b) => (a.type || "").localeCompare(b.type || ""));
    const hdrs = [];
    for (const f of s.matchAll(/<w:headerReference\b[^>]*>/g)) {
      const type = attr(f[0], "type");
      const id = (f[0].match(/r:id="([^"]+)"/) || [])[1];
      const target = rels.get(id);
      hdrs.push({ type, header: target ? await footerSig(zip, target) : null });
    }
    hdrs.sort((a, b) => (a.type || "").localeCompare(b.type || ""));
    out.push({
      pgSz: attrs(s, "w:pgSz", ["w", "h"]),
      pgMar: attrs(s, "w:pgMar", ["top", "right", "bottom", "left", "header", "footer", "gutter"]),
      docGrid: attrs(s, "w:docGrid", ["type", "linePitch", "charSpace"]),
      footers: ftrs,
      headers: hdrs,
    });
  }
  return out;
}

async function settingsSig(zip) {
  const f = zip.file("word/settings.xml");
  if (!f) return null;
  const x = await f.async("string");
  return {
    evenAndOddHeaders: /<w:evenAndOddHeaders\b/.test(x),
    characterSpacingControl: attr((x.match(/<w:characterSpacingControl\b[^>]*>/) || [""])[0], "val"),
  };
}

async function layoutParts(zip, doc, styles) {
  return {
    styles: stylesSig(styles),
    body: bodySig(doc),
    sections: await sectionSig(zip, doc),
    settings: await settingsSig(zip),
  };
}

async function layoutSignature(zip, doc, styles) {
  return JSON.stringify(await layoutParts(zip, doc, styles));
}

const STYLE_LABEL = {
  Normal: "正文样式（Normal）",
  Title: "公文标题样式（Title）",
  Heading1: "层次标题样式（Heading1）",
  Heading2: "层次标题样式（Heading2）",
  Heading3: "层次标题样式（Heading3）",
  Heading4: "层次标题样式（Heading4）",
  TitleChar: "标题字符样式（TitleChar）",
  Heading1Char: "层次标题字符样式（Heading1Char）",
  Heading2Char: "层次标题字符样式（Heading2Char）",
  Heading3Char: "层次标题字符样式（Heading3Char）",
  Heading4Char: "层次标题字符样式（Heading4Char）",
};

// 行文体与条文体的标题叫法不同（行文体的 Heading1 是「一、」，条文体的 Heading1
// 是「第X章」），层级名按 form 取。字体名与是否居中直接从修复结果的签名里读，不写死：
// 换了标题字体（GONGWEN_FONT_TITLE）或换了文体，报告也不会说错。
const HEADING_ROLE = {
  wen: { 1: "一级", 2: "二级", 3: "三四级", 4: "三四级" },
  zhang: { 1: "章标题", 2: "条与项", 3: "条与项", 4: "条与项" },
};
const FONT_CN = { SimHei: "黑体", KaiTi: "楷体", FangSong: "仿宋", SimSun: "宋体" };

function headingDetail(style) {
  if (!style) return "";
  const [ascii, eastAsia] = style.rPr.font;
  const name = FONT_CN[eastAsia] || FONT_CN[ascii] || eastAsia || ascii || "";
  const jc = style.pPr.jc === "center" ? "居中" : "";
  return [name, jc].filter(Boolean).join("、");
}

function styleNote(id, afterStyle, form) {
  const label = STYLE_LABEL[id] || id;
  const lv = (id.match(/^Heading([1-4])$/) || [])[1];
  if (!lv) return `${label}已校准`;
  const role = (HEADING_ROLE[form] || HEADING_ROLE.wen)[lv];
  const detail = headingDetail(afterStyle);
  // 行文体保持原来的「一级黑体」「二级楷体」；条文体的 Heading1 是章标题，
  // 字体与居中从修复结果读，写成「章标题方正小标宋简体、居中」。
  return detail ? `${label}：${role}${detail}` : `${label}：${role}已校准`;
}

// 正文差异分三类：run 上的直接字体字号、段落上的手设行距、补的「首行不缩进」。
// 只改了其中一类时，措辞不能把另外两类也写进去。
function bodyNote(before, after) {
  const a = Array.isArray(before.body) ? before.body : [];
  const b = Array.isArray(after.body) ? after.body : [];
  let direct = false;
  let spacing = false;
  let flush = false;
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    const x = a[i];
    const y = b[i];
    if (!x || !y) { direct = true; continue; }
    if (JSON.stringify(x.runs) !== JSON.stringify(y.runs)) direct = true;
    if (JSON.stringify(x.pPr.spacing) !== JSON.stringify(y.pPr.spacing)) spacing = true;
    if (JSON.stringify(x.pPr.ind) !== JSON.stringify(y.pPr.ind)) flush = true;
  }
  const parts = [];
  if (direct && spacing) parts.push("直接字体、字号或手设行距交还样式");
  else if (direct) parts.push("直接字体、字号交还样式");
  else if (spacing) parts.push("手设行距交还网格");
  if (flush) parts.push("原不缩进的段落补「首行不缩进」");
  return parts.join("；");
}

function describeLayoutDiff(before, after, { form = "wen" } = {}) {
  const notes = [];
  if (JSON.stringify(before.styles.default) !== JSON.stringify(after.styles.default)) {
    notes.push("正文默认字体：仿宋三号");
  }
  for (const id of KEY_STYLES) {
    if (JSON.stringify(before.styles.styles[id]) !== JSON.stringify(after.styles.styles[id])) {
      notes.push(styleNote(id, after.styles.styles[id], form));
    }
  }
  if (JSON.stringify(before.body) !== JSON.stringify(after.body)) {
    notes.push(`正文直接格式：${bodyNote(before, after) || "直接格式交还样式"}`);
  }
  const page = (parts) => parts.sections.map((s) => ({ pgSz: s.pgSz, pgMar: s.pgMar, docGrid: s.docGrid }));
  const furniture = (parts) => parts.sections.map((s) => ({ footers: s.footers, headers: s.headers }));
  if (JSON.stringify(page(before)) !== JSON.stringify(page(after))) {
    notes.push("版心与网格：A4，上 37 下 35 左 28 右 26 mm，每页 22 行、每行 28 字");
  }
  if (JSON.stringify(furniture(before)) !== JSON.stringify(furniture(after))) {
    notes.push("页码/页眉页脚：单页居右、双页居左（原有的单双页页脚已替换）");
  }
  if (JSON.stringify(before.settings) !== JSON.stringify(after.settings)) {
    if (before.settings?.evenAndOddHeaders !== after.settings?.evenAndOddHeaders) {
      notes.push("页码：奇偶页不同（evenAndOddHeaders）");
    }
    if (before.settings?.characterSpacingControl !== after.settings?.characterSpacingControl) {
      notes.push("标点压缩：settings 补 compressPunctuation");
    }
  }
  return notes;
}

/**
 * 版式（直接改）的报告条目。语义比对算得出结果，就只列真实变化项；
 * 算不出来（原件的部分读不出，origParts 为空）时保留 fixStyles/fixDocument
 * 的逐项说明，不能只剩一句笼统的总述、把修复细节丢掉。
 */
function layoutChangeNotes({ own = false, before = null, after = null, form = "wen", titleNote = "", fallback = [] } = {}) {
  const notes = [own
    ? "本工具出稿，但检测到版式/正文格式已被改动，重新校准"
    : "检测到版式/正文格式需要校准，已直接改"];
  if (titleNote) notes.push(titleNote);
  const diff = before && after ? describeLayoutDiff(before, after, { form }) : null;
  if (diff && diff.length) notes.push(...diff);
  else notes.push(...fallback);
  return notes;
}

module.exports = { layoutSignature, layoutParts, describeLayoutDiff, layoutChangeNotes };
