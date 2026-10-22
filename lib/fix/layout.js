"use strict";
/**
 * layout.js —— 版式修复（tools/fix-doc.js 用）。直接改，不留修订痕迹：
 * 版式不是内容，逐条接受拒绝没有意义；原文件不动，另存新文件。
 *
 * 改什么（数值与 lib/styles.js、lib/gongwen-core.js 的样式底座同源）：
 *   页面      A4，版心 上 37 下 35 左 28 右 26 mm，每页 22 行每行 28 字
 *   正文      默认字体仿宋三号；Normal 首行缩进二字、两端对齐
 *   层次标题  按样式名 heading 1～4 认：黑体／楷体／仿宋／仿宋，三号，不加粗，去掉段前段后
 *   公文标题  首页前几段里居中、非红色的那一段（及紧接的回行段）：小标宋二号
 *   正文直接格式  段落外、表格外的 run 上手设的字体与字号去掉，交给样式；
 *              段落上手设的行距去掉，交给网格。加粗、下划线保留
 *   页码      「— 1 —」宋体四号，单页居右空一字、双页居左空一字
 *   标点压缩  settings 补 characterSpacingControl
 *
 * 不改：红头（红色段落）、表格内部、页眉、首页页脚、图片与文本框。
 * Normal 加了首行缩进二字，原来不缩进的段落（主送机关、居中右对齐的段落、表格里的段落、
 * 用空格手动缩进的段落）各自补一个「首行不缩进」，保持原样。
 * 版头的位置尺寸（35mm、4mm）在别人的文件里没有可靠的锚点，不动，交给批注。
 */

const { editRPr, editPPr, children, findElement, editParaPPr, editParaRuns } = require("./ooxml");
const { tableRanges } = require("../table-read");

const MARGIN = '<w:pgMar w:top="2098" w:right="1474" w:bottom="1984" w:left="1587" w:header="850" w:footer="1392" w:gutter="0"/>';
const PGSZ = '<w:pgSz w:w="11906" w:h="16838"/>';
const GRID = '<w:docGrid w:type="linesAndChars" w:linePitch="579" w:charSpace="-842"/>';
const FONT = (f) => `<w:rFonts w:ascii="${f}" w:eastAsia="${f}" w:hAnsi="${f}"/>`;
const HEAD_FONT = { 1: "SimHei", 2: "KaiTi", 3: "FangSong", 4: "FangSong" };
const ZHANG_HEAD_FONT = (titleFont) => ({ 1: titleFont, 2: "FangSong", 3: "FangSong", 4: "FangSong" });

// 段落样式表：id → { basedOn, jc }。标题的居中可能是样式给的（Title 样式），
// 只认段落上的 <w:jc> 会把本工具自己的标题漏掉。
function styleMap(styles) {
  const map = new Map();
  for (const m of (styles || "").matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const attr = m[1];
    if (!/w:type="paragraph"/.test(attr)) continue;
    const id = (attr.match(/w:styleId="([^"]+)"/) || [])[1];
    if (!id) continue;
    const body = m[2];
    map.set(id, {
      basedOn: (body.match(/<w:basedOn w:val="([^"]+)"/) || [])[1] || "",
      jc: (body.match(/<w:jc w:val="([^"]+)"/) || [])[1] || "",
    });
  }
  return map;
}

// 段落的有效对齐：直接格式优先，其次样式链。
function resolveJc(pXml, styles) {
  const direct = (pXml.match(/<w:jc w:val="([^"]+)"/) || [])[1];
  if (direct) return direct;
  let id = (pXml.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || "";
  const seen = new Set();
  while (id && styles.has(id) && !seen.has(id)) {
    seen.add(id);
    const s = styles.get(id);
    if (s.jc) return s.jc;
    id = s.basedOn;
  }
  return "";
}

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const footerXml = (side) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
  + `<w:ftr ${W_NS}><w:p><w:pPr><w:ind w:${side === "right" ? "right" : "left"}="280" w:firstLine="0" w:firstLineChars="0"/>`
  + `<w:jc w:val="${side}"/></w:pPr>`
  + ["— ", null, " —"].map((t) => {
    const rPr = `<w:rPr>${FONT("SimSun")}<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`;
    return t === null
      ? `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve">PAGE</w:instrText></w:r>`
        + `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>1</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
      : `<w:r>${rPr}<w:t xml:space="preserve">${t}</w:t></w:r>`;
  }).join("") + `</w:p></w:ftr>`;

// OOXML 的子元素有固定顺序，插错位置 Word 会报文件损坏。插入时找排在它后面的
// 第一个已有元素，插到它前面；一个都没有就插到父元素末尾。
function insertBefore(xml, closeTag, value, followers) {
  for (const f of followers) {
    const m = new RegExp(`<${f}(?=[\\s>/])`).exec(xml);
    if (m) return xml.slice(0, m.index) + value + xml.slice(m.index);
  }
  return xml.replace(new RegExp(`</${closeTag}>$`), value + `</${closeTag}>`);
}

// sectPr 的子元素顺序（ECMA-376 CT_SectPr），取 pgSz 之后的部分
const SECT_ORDER = ["w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols",
  "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter",
  "w:docGrid", "w:printerSettings", "w:sectPrChange"];

// 把某个元素换成新值；没有就按 schema 顺序插入
function setChild(sect, tag, value) {
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?/>|<${tag}(?:\\s[^>]*)?>[\\s\\S]*?</${tag}>`);
  if (re.test(sect)) return sect.replace(re, value);
  return insertBefore(sect, "w:sectPr", value, SECT_ORDER.slice(SECT_ORDER.indexOf(tag) + 1));
}

// 顶层的每个 name 元素交给 fn 改（按深度配对；嵌在里面的同名元素随外层一起交出）
function eachElement(xml, name, fn) {
  let out = "", i = 0, f;
  while ((f = findElement(xml, name, i))) {
    out += xml.slice(i, f.start) + fn(xml.slice(f.start, f.end));
    i = f.end;
  }
  return out + xml.slice(i);
}

// settings 里排在 evenAndOddHeaders、characterSpacingControl 之后的元素（ECMA-376 CT_Settings）
const SETTINGS_AFTER_CSC = ["w:printTwoOnOne", "w:strictFirstAndLastChars", "w:noLineBreaksAfter",
  "w:noLineBreaksBefore", "w:savePreviewPicture", "w:doNotValidateAgainstSchema", "w:saveInvalidXml",
  "w:ignoreMixedContent", "w:alwaysShowPlaceholderText", "w:doNotDemarcateInvalidXml", "w:saveXmlDataOnly",
  "w:useXSLTWhenSaving", "w:saveThroughXslt", "w:showXMLTags", "w:alwaysMergeEmptyNamespace", "w:updateFields",
  "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars", "w:rsids", "m:mathPr",
  "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats",
  "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions", "w:readModeInkLockDown", "w:smartTagType",
  "sl:schemaLibrary", "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator"];
const SETTINGS_AFTER_EOH = ["w:bookFoldRevPrinting", "w:bookFoldPrinting", "w:bookFoldPrintingSheets",
  "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing", "w:displayHorizontalDrawingGridEvery",
  "w:displayVerticalDrawingGridEvery", "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin",
  "w:drawingGridVerticalOrigin", "w:doNotShadeFormData", "w:noPunctuationKerning", "w:characterSpacingControl",
  ...SETTINGS_AFTER_CSC];


const STYLE_ORDER = ["w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine", "w:hidden",
  "w:uiPriority", "w:semiHidden", "w:unhideWhenUsed", "w:qFormat", "w:locked", "w:personal",
  "w:personalCompose", "w:personalReply", "w:rsid", "w:pPr", "w:rPr", "w:tblPr", "w:trPr", "w:tcPr", "w:tblStylePr"];

const innerOf = (xml) => xml.replace(/^<[^>]*?\/>$/, "").replace(/^<[^>]+>/, "").replace(/<\/[^>]+>$/, "");

// 改一个容器元素（样式、docDefaults 的 rPrDefault 等）里的 pPr／rPr，没有就建
function editContainer(xml, order, { ppr, rpr }) {
  const open = xml.match(/^<[^>]+>/)[0];
  const tag = open.match(/^<([\w:]+)/)[1];
  let kids = children(innerOf(xml));
  const apply = (name, op, edit) => {
    if (!op) return;
    const k = kids.find((c) => c.name === name);
    const inner = edit(k ? innerOf(k.xml) : "", op);
    const xmlNew = inner ? `<${name}>${inner}</${name}>` : "";
    if (k) k.xml = xmlNew; else if (xmlNew) kids.push({ name, xml: xmlNew });
  };
  apply("w:pPr", ppr, editPPr);
  apply("w:rPr", rpr, editRPr);
  const rank = (n) => { const r = order.indexOf(n); return r < 0 ? order.length : r; };
  kids = kids.map((k, i) => ({ ...k, i })).sort((a, b) => rank(a.name) - rank(b.name) || a.i - b.i);
  return `${open.replace(/\/>$/, ">")}${kids.map((k) => k.xml).join("")}</${tag}>`;
}

const SZ = (n) => ({ "w:sz": `<w:sz w:val="${n}"/>`, "w:szCs": `<w:szCs w:val="${n}"/>` });

/** 样式表：docDefaults、Normal、标题 1～4。返回 { xml, headingIds:{id:level} }。 */
function fixStyles(styles, { form = "wen", titleFont = "方正小标宋简体" } = {}, notes) {
  let xml = styles;
  const info = styleMap(styles);
  const defRpr = { set: { "w:rFonts": FONT("FangSong"), ...SZ(32) } };
  if (/<w:rPrDefault>/.test(xml)) {
    xml = xml.replace(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>|<w:rPrDefault\/>/, (blk) => editContainer(blk, ["w:rPr"], { rpr: defRpr }));
  } else if (/<w:docDefaults>/.test(xml)) {
    xml = xml.replace(/<w:docDefaults>/, `<w:docDefaults><w:rPrDefault><w:rPr>${editRPr("", defRpr)}</w:rPr></w:rPrDefault>`);
  } else {
    xml = xml.replace(/(<w:styles[^>]*>)/, `$1<w:docDefaults><w:rPrDefault><w:rPr>${editRPr("", defRpr)}</w:rPr></w:rPrDefault></w:docDefaults>`);
  }
  const headingIds = {};
  const headFont = form === "zhang" ? ZHANG_HEAD_FONT(titleFont) : HEAD_FONT;
  xml = xml.replace(/<w:style\b[^>]*w:type="paragraph"[^>]*>[\s\S]*?<\/w:style>/g, (st) => {
    const id = (st.match(/w:styleId="([^"]+)"/) || [])[1];
    const name = ((st.match(/<w:name w:val="([^"]+)"/) || [])[1] || "").toLowerCase();
    const lv = (name.match(/^heading ([1-4])$/) || [])[1];
    if (id === "Title") {
      return editContainer(st, STYLE_ORDER, {
        ppr: { set: { "w:ind": '<w:ind w:firstLineChars="0" w:firstLine="0"/>', "w:jc": '<w:jc w:val="center"/>' } },
        rpr: { set: { "w:rFonts": FONT(titleFont), ...SZ(44) }, remove: ["w:b"] },
      });
    }
    if (/w:default="1"/.test(st)) {
      notes.push("正文样式：仿宋三号、首行缩进二字、两端对齐、行距交给网格");
      return editContainer(st, STYLE_ORDER, {
        ppr: { set: { "w:ind": '<w:ind w:firstLineChars="200" w:firstLine="632"/>', "w:jc": '<w:jc w:val="both"/>' },
          remove: ["w:spacing", "w:snapToGrid"] },
        rpr: { remove: ["w:rFonts", "w:sz", "w:szCs"] },
      });
    }
    if (!lv) return st;
    headingIds[id] = +lv;
    // 条文体：章标题是方正小标宋、居中，保留段前段后；条/项是仿宋，不居中。
    // 行文体：一级黑体、二级楷体、三四级仿宋，不居中，去掉段前段后。
    const zhangH1 = form === "zhang" && lv === "1";
    return editContainer(st, STYLE_ORDER, {
      ppr: zhangH1
        ? { set: { "w:jc": '<w:jc w:val="center"/>' }, remove: ["w:snapToGrid"] }
        : { remove: ["w:spacing", "w:jc", "w:snapToGrid"] },
      rpr: { set: { "w:rFonts": FONT(headFont[lv]), ...SZ(32) }, remove: ["w:b", "w:bCs", "w:color"] },
    });
  });
  const n = Object.keys(headingIds).length;
  if (n) {
    notes.push(form === "zhang"
      ? `层次标题样式 ${n} 个：章小标宋居中、条与项仿宋，三号，不加粗`
      : `层次标题样式 ${n} 个：一级黑体、二级楷体、三四级仿宋，三号，不加粗，去掉段前段后`);
  }
  return { xml, headingIds, styles: info };
}

const RED_RE = /<w:color w:val="(?:FF0000|C00000|E60012|D7261E|[C-F][0-9A-F]0000|[C-F][0-9A-F]1[0-9A-F]1[0-9A-F])"/i;
const isRed = (p) => RED_RE.test(p);

/** 样式表里直接定义成红色的样式 ID（红头常用一个红色段落样式或字符样式）。 */
function redStyleIds(styles) {
  const ids = new Set();
  for (const m of (styles || "").matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
    if (RED_RE.test(m[2])) ids.add(m[1]);
  }
  return ids;
}

/** 正文：版心、页码、标题、直接格式。 */
function fixDocument(xml, { headingIds, titleFont, redStyles = new Set(), styles = new Map() }, notes) {
  let out = xml;
  const redPara = (p) => isRed(p) || [...p.matchAll(/<w:[pr]Style w:val="([^"]+)"/g)].some((m) => redStyles.has(m[1]));

  let nSect = 0;
  out = eachElement(out, "w:sectPr", (whole) => {
    nSect++;
    // 页面设置的修订（sectPrChange）里嵌着旧的 sectPr，先摘出来，改完原样放回
    const ch = findElement(whole, "w:sectPrChange");
    const change = ch ? whole.slice(ch.start, ch.end) : "";
    const sect = ch ? whole.slice(0, ch.start) + whole.slice(ch.end) : whole;
    let s = setChild(sect, "w:pgSz", PGSZ);
    s = setChild(s, "w:pgMar", MARGIN);
    s = setChild(s, "w:docGrid", GRID);
    // 页码：去掉原有的单页、双页页脚引用，换成公文页码（首页页脚不动）
    s = s.replace(/<w:footerReference w:type="(?:default|even)"[^>]*\/>/g, "")
      .replace(/(<w:sectPr\b[^>]*>)/, '$1<w:footerReference w:type="default" r:id="rIdGwFtrOdd"/><w:footerReference w:type="even" r:id="rIdGwFtrEven"/>');
    // settings 打开了单双页不同：只有默认页眉的节，双页沿用同一个页眉，不然双页页眉会消失
    const hdr = s.match(/<w:headerReference w:type="default" r:id="([^"]+)"\/>/);
    if (hdr && !/<w:headerReference w:type="even"/.test(s)) {
      s = s.replace(hdr[0], `${hdr[0]}<w:headerReference w:type="even" r:id="${hdr[1]}"/>`);
    }
    return change ? s.replace(/<\/w:sectPr>$/, `${change}</w:sectPr>`) : s;
  });
  notes.push(`版心与网格（${nSect} 节）：A4，上 37 下 35 左 28 右 26 mm，每页 22 行、每行 28 字`);
  notes.push("页码：— 1 —，宋体四号，单页居右、双页居左（原有的单双页页脚已替换）");

  const tbl = tableRanges(out);
  const inTbl = (i) => tbl.some(([a, b]) => i >= a && i < b);
  const box = [];
  for (const m of out.matchAll(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g)) box.push([m.index, m.index + m[0].length]);
  const inBox = (i) => box.some(([a, b]) => i >= a && i < b);

  const paras = [];
  for (const m of out.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g)) {
    const text = (m[0].match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, "")).join("");
    const style = (m[0].match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || "";
    paras.push({ at: m.index, xml: m[0], text, style, table: inTbl(m.index), box: inBox(m.index), idx: paras.length });
  }

  const centered = (p) => resolveJc(p.xml, styles) === "center";
  // 这些段落不是「正文/层次标题」，它们的直接字体字号是标准的组成部分，不能交还样式。
  const nearTable = (pos) => tbl.some(([a]) => a > pos && a - pos < 4000);
  const isCaption = (p) => centered(p) && /^\s*(表|图|附表)\s*\d/.test(p.text)
    && (/w:keepNext/.test(p.xml) || nearTable(p.at));
  const isAttendee = (p) => /^(出席|请假|列席)[：:]/.test(p.text);
  const isBanji = (p) => /<w:framePr\b/.test(p.xml);
  const isDocNo = (p) => /〔\d{4}〕\s*第?\s*\d+\s*号/.test(p.text);
  const isSigner = (p) => /^签发人[：:]/.test(p.text.trim());
  // 多签发人时，续行只有姓名、右对齐；它们与首行同属版头。
  const isSignerRow = (p) => p.idx < 12 && resolveJc(p.xml, styles) === "right";
  // 落款署名同样按专名保护；右对齐的日期仍按正文交还样式（日期无直接格式时本来就不变）。
  const isBottomRight = (p) => p.idx > paras.length - 12 && resolveJc(p.xml, styles) === "right"
    && !/^\s*(?:\d{4}|[一二三四五六七八九〇]{4})\s*年/.test(p.text);
  const isTopField = (p, i) => i < 6 && (/^\d{6}$/.test(p.text.trim())
    || /^(秘密|机密|绝密|特密|内部)/.test(p.text.trim())
    || /^(特急|加急|平急)/.test(p.text.trim()));

  // 附件页：顶格「附件N」及其后空行/居中的附件标题，是独立要素，保留原字体。
  const fujianAt = new Set();
  {
    let inHead = false;
    for (const p of paras) {
      const t = p.text.trim();
      if (/^附件(?:\d+)?$/.test(t) || /〔\d{4}〕\d+号附件\d*$/.test(t)) { fujianAt.add(p.at); inHead = true; continue; }
      if (inHead) {
        if (!t) { fujianAt.add(p.at); continue; }
        if (centered(p)) { fujianAt.add(p.at); continue; }
        inHead = false;
      }
    }
  }

  // 公文标题：前 15 个有字段落里第一个居中、非红、不像发文字号或日期的段落，
  // 及紧接的居中段。居中要解析样式链（Title 样式就是居中），不能只看段落 jc。
  const titleAt = new Set();
  {
    const firsts = paras.filter((p) => p.text.trim() && !p.table && !p.box && !isTopField(p, p.idx)).slice(0, 15);
    const excluded = (p) => /[〔【[]\d{4}[〕】\]]|\d\s*号$|\d{4}\s*年\s*\d{1,2}\s*月|[〇零]\S*年/.test(p.text) || p.text.length > 80;
    const k = firsts.findIndex((p) => centered(p) && !/^Heading[1-4]$/.test(p.style) && !/<w:numPr>/.test(p.xml)
      && !redPara(p.xml) && !isCaption(p) && !isAttendee(p) && !isBanji(p) && !fujianAt.has(p.at) && !excluded(p));
    if (k >= 0) {
      titleAt.add(firsts[k].at);
      for (let j = k + 1; j < firsts.length && j <= k + 2; j++) {
        const q = firsts[j];
        if (!centered(q) || /^Heading[1-4]$/.test(q.style) || /<w:numPr>/.test(q.xml)
          || redPara(q.xml) || isCaption(q) || isAttendee(q) || isBanji(q) || fujianAt.has(q.at)) break;
        titleAt.add(q.at);
      }
    }
  }

  // 主送机关：标题之后第一个有字的段落，以冒号结尾；找不到标题时也在前段里找。
  let zhusongAt = -1;
  {
    const after = titleAt.size ? Math.max(...titleAt) : -1;
    const z = paras.find((p) => p.at > after && p.text.trim() && !p.table && !p.box
      && !redPara(p.xml) && !isCaption(p) && !isAttendee(p) && !fujianAt.has(p.at));
    if (z && /[：:]\s*$/.test(z.text) && z.text.length <= 80) zhusongAt = z.at;
  }

  // 原来不缩进、Normal 加了缩进以后会跟着缩进的段落：补「首行不缩进」保持原样
  const NO_IND = '<w:ind w:firstLine="0" w:firstLineChars="0"/>';
  const hasInd = (s) => {
    const f = findElement(s, "w:pPr");
    return !!f && children(s.slice(f.start, f.end).replace(/^<w:pPr>|<\/w:pPr>$/g, "")).some((c) => c.name === "w:ind");
  };
  const keepFlush = (p, s) => !(p.style && /^Heading[1-4]$/.test(p.style)) && !hasInd(s) && (p.table || p.box || p.at === zhusongAt
    || ["center", "right", "end"].includes(resolveJc(s, styles)) || /^[\s\u3000]/.test(p.text));

  const directRunFormat = (s) => [...s.matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)].some((r) => {
    const rpr = (r[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [])[0] || "";
    return /<w:rFonts\b|<w:sz\b/.test(rpr);
  });

  const protectedRole = (p) => !p.text.trim() || p.table || p.box || redPara(p.xml) || isBanji(p) || isCaption(p)
    || isAttendee(p) || isDocNo(p) || isSigner(p) || isSignerRow(p) || isBottomRight(p)
    || fujianAt.has(p.at) || isTopField(p, p.idx);

  let nRuns = 0;
  let title = "";
  let titleStyled = false;
  for (const p of [...paras].reverse()) {   // 从后往前替换，前面的位置不变
    if (protectedRole(p)) {
      // 红头、表格、文本框、表题、版记、附件页、纪要名单、份号密级：只补首行不缩进，
      // 不动 run 的字体字号与手设行距——这些是国标要素自己的格式。
      if (keepFlush(p, p.xml)) out = out.slice(0, p.at) + editParaPPr(p.xml, { set: { "w:ind": NO_IND } }) + out.slice(p.at + p.xml.length);
      continue;
    }
    let s = p.xml;
    if (!titleAt.has(p.at) && keepFlush(p, s)) s = editParaPPr(s, { set: { "w:ind": NO_IND } });
    if (titleAt.has(p.at)) {
      title = p.text + title;
      const sid = (s.match(/<w:pStyle w:val="([^"]+)"/) || [])[1];
      // 样式就是 Title 且没有直接字体字号：样式已校准，不必再给 run 加直接格式，
      // 否则合规件会被「改一遍」却看不出差别。
      if (!(sid === "Title" && !directRunFormat(s))) {
        s = editParaRuns(s, { set: { "w:rFonts": FONT(titleFont), ...SZ(44) }, remove: ["w:b", "w:bCs"] }, { create: true }).xml;
      } else {
        titleStyled = true;
      }
      // Title 样式已经给了不缩进；不要再给合规件加一条直接缩进
      if (!(sid === "Title" && !hasInd(s))) s = editParaPPr(s, { set: { "w:ind": NO_IND } });
      // 标题上手设的行距（如 1.5 倍）交还网格；已经是 579 exact 的合规标题不动。
      {
        const f = findElement(s, "w:pPr");
        const inner = f ? s.slice(f.start, f.end).replace(/^<w:pPr>|<\/w:pPr>$|^<w:pPr\/>$/g, "") : "";
        const sp = children(inner).find((c) => c.name === "w:spacing");
        if (sp && !(/w:line="579"/.test(sp.xml) && /w:lineRule="exact"/.test(sp.xml))) {
          const keep = sp.xml.replace(/\s*w:line(?:Rule)?="[^"]*"/g, "");
          s = editParaPPr(s, { set: { "w:spacing": /w:(?:before|after)/.test(keep) ? keep : "" }, remove: ["w:snapToGrid"] });
        }
      }
      out = out.slice(0, p.at) + s + out.slice(p.at + p.xml.length);
      continue;
    }
    {
      const sid = (s.match(/<w:pStyle w:val="([^"]+)"/) || [])[1];
      const r = editParaRuns(s, { remove: ["w:rFonts", "w:sz", "w:szCs", ...(sid && headingIds[sid] ? ["w:b", "w:bCs"] : [])] });
      s = r.xml; nRuns += r.changed;
    }
    // 段落上手设的行距交还网格（段前段后保留）
    {
      const f = findElement(s, "w:pPr");
      const inner = f ? s.slice(f.start, f.end).replace(/^<w:pPr>|<\/w:pPr>$|^<w:pPr\/>$/g, "") : "";
      const sp = children(inner).find((c) => c.name === "w:spacing");
      const keep = sp ? sp.xml.replace(/\s*w:line(?:Rule)?="[^"]*"/g, "") : "";
      const set = sp ? { "w:spacing": /w:(?:before|after)/.test(keep) ? keep : "" } : {};
      if (f) s = editParaPPr(s, { set, remove: ["w:snapToGrid"] });
    }
    out = out.slice(0, p.at) + s + out.slice(p.at + p.xml.length);
  }
  if (title) notes.push(`公文标题：「${title.slice(0, 30)}」${titleStyled ? "已按 Title 样式校准" : "改为"}${titleStyled ? "" : (titleFont === "方正小标宋简体" ? "方正小标宋" : titleFont) + "二号"}`);
  else notes.push("公文标题：前几段里没找到居中的标题段，未改");
  if (nRuns) notes.push(`正文直接格式：${nRuns} 个 run 上手设的字体字号交还样式；手设行距交还网格`);
  return out;
}

/** 补页脚部件、关系、内容类型、settings。parts：{ rels, types, settings } 字符串。 */
function fixPackage(zip, parts, notes) {
  zip.file("word/footerGwOdd.xml", footerXml("right"));
  zip.file("word/footerGwEven.xml", footerXml("left"));
  const FTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
  let rels = parts.rels.replace(/<Relationship [^>]*Id="rIdGwFtr(?:Odd|Even)"[^>]*\/>/g, "");
  rels = rels.replace(/<\/Relationships>/, `<Relationship Id="rIdGwFtrOdd" Type="${FTR}" Target="footerGwOdd.xml"/><Relationship Id="rIdGwFtrEven" Type="${FTR}" Target="footerGwEven.xml"/></Relationships>`);
  const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
  let types = parts.types;
  for (const f of ["footerGwOdd", "footerGwEven"]) {
    if (!types.includes(`/word/${f}.xml`)) types = types.replace(/<\/Types>/, `<Override PartName="/word/${f}.xml" ContentType="${CT}"/></Types>`);
  }
  let settings = parts.settings;
  if (!/<w:evenAndOddHeaders/.test(settings)) {
    settings = insertBefore(settings, "w:settings", "<w:evenAndOddHeaders/>", SETTINGS_AFTER_EOH);
  }
  if (!/<w:characterSpacingControl/.test(settings)) {
    settings = insertBefore(settings, "w:settings", '<w:characterSpacingControl w:val="compressPunctuation"/>', SETTINGS_AFTER_CSC);
    notes.push("标点压缩：settings 补 compressPunctuation");
  }
  return { rels, types, settings };
}

module.exports = { fixStyles, fixDocument, fixPackage, isRed, redStyleIds, styleMap, resolveJc };
