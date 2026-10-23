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
  Heading1: "层次标题样式（Heading1）：一级黑体",
  Heading2: "层次标题样式（Heading2）：二级楷体",
  Heading3: "层次标题样式（Heading3）：三四级仿宋",
  Heading4: "层次标题样式（Heading4）：三四级仿宋",
  TitleChar: "标题字符样式（TitleChar）",
  Heading1Char: "层次标题字符样式（Heading1Char）",
  Heading2Char: "层次标题字符样式（Heading2Char）",
  Heading3Char: "层次标题字符样式（Heading3Char）",
  Heading4Char: "层次标题字符样式（Heading4Char）",
};

function describeLayoutDiff(before, after) {
  const notes = [];
  if (JSON.stringify(before.styles.default) !== JSON.stringify(after.styles.default)) {
    notes.push("正文默认字体：仿宋三号");
  }
  for (const id of KEY_STYLES) {
    if (JSON.stringify(before.styles.styles[id]) !== JSON.stringify(after.styles.styles[id])) {
      notes.push(`${STYLE_LABEL[id] || id}已校准`);
    }
  }
  if (JSON.stringify(before.body) !== JSON.stringify(after.body)) {
    notes.push("正文直接格式：直接字体、字号或手设行距交还样式");
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

module.exports = { layoutSignature, layoutParts, describeLayoutDiff };
