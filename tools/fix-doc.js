#!/usr/bin/env node
/**
 * fix-doc.js —— 修复现成的公文 docx：版式直接改，体例按修订痕迹改，剩下的插批注。
 *
 * 用法：node tools/fix-doc.js 稿子.docx [--out 输出.docx] [--no-layout] [--dry-run]
 *       默认输出到同目录「原名-修订.docx」，原文件不动。
 *
 * 三档：
 *   版式      直接改（lib/fix/layout.js）：版心网格、正文与层次标题字体字号、公文标题、页码
 *   体例与规范层文风  改成 Word 修订（lib/fix/text-rules.js）：写稿人逐条接受或拒绝；
 *            全部拒绝等于原文（tests/fix.js 断言）
 *   其余      不改，插批注：数字自洽（合计、跨表、计数与列举）、改法要靠判断的体例项、
 *            碰到域与超链接没法安全改的。数字不自动改——合计错还是分项错，工具判断不了
 *
 * 2026-10-23 起：
 *   · 角色识别：红头、表格、文本框、表题、版记、附件页、纪要名单、份号密级不参与
 *     「交还样式」，这些要素的字体字号是国标的一部分，不能被当成正文剥掉；
 *   · 信函/纪要格式自动跳过通用版式（--layout 可强制）；
 *   · 条文体（第X章/第X条）按章小标宋、条仿宋处理，不再套行文体的一级黑体/二级楷体；
 *   · --dry-run 只报告、不写文件；建议对要交付的稿子先跑一遍。
 *
 * 改完对输出件重跑三道机检，批注就是机检剩下的那些，所以批注与机检口径一致。
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const JSZip = require("jszip");
const { parseParagraph, applyEdits, unitalic } = require("../lib/fix/runs");
const { isDocNoLine, isZhangDoc } = require("../lib/fix/classify");
const { editsFor } = require("../lib/fix/text-rules");
const { fixStyles, fixDocument, fixPackage, redStyleIds, isRed, styleMap, resolveJc } = require("../lib/fix/layout");
const { addComments, registerComments } = require("../lib/fix/comments");
const checkDoc = require("../checks/check-doc");
const checkNum = require("../checks/check-num");
const { scan } = require("../checks/style-scan");
const { textOf, tableRanges } = require("../lib/table-read");

const AUTHOR = "gongwen-kit";
const TITLE_FONT = process.env.GONGWEN_FONT_TITLE || "方正小标宋简体";

function args() {
  const a = process.argv.slice(2);
  const o = { layout: true, forceLayout: false, dryRun: false };
  for (let i = 0; i < a.length; i++) {
    if (a[i] === "--out") {
      const v = a[++i];
      if (!v || v.startsWith("--")) {
        console.error("--out 后面要跟一个输出文件。");
        process.exit(2);
      }
      o.out = v;
    } else if (a[i] === "--no-layout") o.layout = false;
    else if (a[i] === "--layout") o.forceLayout = true;
    else if (a[i] === "--dry-run") o.dryRun = true;
    else if (!o.file) o.file = a[i];
  }
  if (!o.file) {
    console.error("用法：node tools/fix-doc.js 稿子.docx [--out 输出.docx] [--no-layout] [--dry-run]");
    process.exit(2);
  }
  o.out = o.out || path.join(path.dirname(o.file), path.basename(o.file).replace(/\.docx$/i, "") + "-修订.docx");
  if (path.resolve(o.out) === path.resolve(o.file)) {
    console.error("输出路径和原文件相同，会覆盖原稿。换一个 --out。");
    process.exit(2);
  }
  return o;
}

const read = async (zip, name) => (zip.file(name) ? zip.file(name).async("string") : "");

// 体例修复：逐段算改动、写成修订。角色保护与 layout 同一口径：
// 红头/表格/文本框/表题/版记/附件页/纪要名单/份号密级不改文本。
function textPass(docXml, styles, headingIds, rev, tally, { redStyles = new Set(), styleInfo = new Map() } = {}) {
  // 样式自带编号的段落（标题样式里的 numPr）也算有自动编号
  const numberedStyles = new Set();
  for (const m of styles.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
    if (/<w:numPr>/.test(m[2])) numberedStyles.add(m[1]);
  }
  const plain = docXml.split("</w:p>").map(textOf).join("\n");
  const spacedDates = /\d\s+年\s+\d/.test(plain);
  const tbl = tableRanges(docXml);
  const nearTable = (pos) => tbl.some(([a]) => a > pos && a - pos < 4000);
  const boxes = [];
  for (const m of docXml.matchAll(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g)) boxes.push([m.index, m.index + m[0].length]);
  const inRanges = (rs, pos) => rs.some(([a, b]) => pos >= a && pos < b);
  const redPara = (p) => isRed(p) || [...p.matchAll(/<w:[pr]Style w:val="([^"]+)"/g)].some((m) => redStyles.has(m[1]));
  const centered = (p) => resolveJc(p, styleInfo) === "center";

  let skipped = 0;
  let idx = 0;
  let fujianHead = false;
  const xml = docXml.replace(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g, (p, offset) => {
    const i = idx++;
    const text = textOf(p);
    const t = text.trim();
    if (!t) return p;   // 空段是版式间距，不能当成正文交还行距
    // 附件页：顶格「附件N」及其后的空行/居中附件标题不参与文本修订
    if (/^附件(?:\d+)?$/.test(t) || /〔\d{4}〕\d+号附件\d*$/.test(t)) { fujianHead = true; return p; }
    if (fujianHead) {
      if (!t || centered(p)) return p;
      fujianHead = false;
    }
    const jc = resolveJc(p, styleInfo);
    // 右对齐的署名/机关名是专名，不能按异形词改；右对齐的日期（阿拉伯或中文数字）仍允许修。
    const isDate = /^\s*(?:\d{4}|[一二三四五六七八九〇]{4})\s*年/.test(t);
    const rightName = jc === "right" && !isDate;
    const isDocNo = (p, text, idx) => idx < 12
      && isDocNoLine(text, /<w:tab\/>/.test(p) || /\t/.test(p));
    const protectedRole = inRanges(tbl, offset) || inRanges(boxes, offset) || redPara(p)
      || (i < 12 && jc === "right")
      || rightName
      || /<w:framePr\b/.test(p)
      || isDocNo(p, t, i)
      || /^签发人[：:]/.test(t)
      || (centered(p) && /^\s*(表|图|附表)\s*\d/.test(text)
        && (/w:keepNext/.test(p) || nearTable(offset)))
      || /^(出席|请假|列席)[：:]/.test(t)
      || (i < 6 && (/^\d{6}$/.test(t) || /^(秘密|机密|绝密|特密|内部)/.test(t) || /^(特急|加急|平急)/.test(t)));
    if (protectedRole) return p;

    let para;
    try { para = parseParagraph(p); } catch { return p; }   // 解析不了的段落原样保留
    const sid = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1];
    const ctx = {
      spacedDates,
      heading: !!(sid && headingIds[sid]) || /<w:outlineLvl w:val="[0-3]"/.test(p),
      numbered: /<w:numPr>/.test(p) || numberedStyles.has(sid),
    };
    let out = p;
    const edits = editsFor(para.text, ctx);
    if (edits.length) {
      const r = applyEdits(para, edits, rev);
      skipped += r.skipped.length;
      for (const e of r.applied) tally[e.rule] = (tally[e.rule] || 0) + 1;
      if (r.xml) out = r.xml;
    }
    const it = unitalic(out, rev);
    if (it.n) tally["斜体改正体"] = (tally["斜体改正体"] || 0) + it.n;
    return it.xml;
  });
  return { xml, skipped };
}

// 对一份 docx 跑三道机检，收集剩下的问题（带锚点）
async function remaining(file) {
  const items = [];
  const doc = await checkDoc.readDocx(file);
  const dh = checkDoc.check(file, doc);
  for (const h of dh) items.push({ kind: "体例", ...h });
  for (const h of dh.hints || []) items.push({ kind: "体例提示", ...h });
  const num = await checkNum.check(file);
  for (const h of num) items.push({ kind: "数字", ...h });
  for (const h of num.hints || []) items.push({ kind: "数字提示", ...h });
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const body = (await read(zip, "word/document.xml")).split("</w:p>").map(textOf).filter(Boolean).join("\n")
    .replace(/「[^」\n]*」/g, "　").replace(/“[^”\n]*”/g, "　").replace(/《[^》\n]*》/g, "　");
  const st = scan(body);
  st.norm.forEach((line) => items.push({ kind: "文风", rule: "规范", detail: line, anchors: st.anchors }));
  return { items, style: st.style };
}

// 一条命中拆成批注：每个不同的锚点一条，上限 20
function toComments(items) {
  const out = [];
  for (const it of items) {
    const label = it.kind === "数字提示" ? "数字·提示（不一定错，请核对）"
      : it.kind === "体例提示" ? "体例·提示（不一定错，请核对）" : it.kind;
    const text = `【${label}·${it.rule}】${it.detail}`;
    const seen = new Set();
    const anchors = (it.anchors || []).filter((a) => {
      const k = typeof a === "string" ? a : `${a.table}|${a.text}`;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
    if (it.kind === "文风" || anchors.length <= 1) { out.push({ text, anchors }); continue; }
    anchors.slice(0, 20).forEach((a) => out.push({ text, anchors: [a] }));
  }
  return out;
}

(async () => {
  const o = args();
  const zip = await JSZip.loadAsync(fs.readFileSync(o.file));
  let docXml = await read(zip, "word/document.xml");
  let styles = await read(zip, "word/styles.xml");
  let rels = await read(zip, "word/_rels/document.xml.rels");
  let types = await read(zip, "[Content_Types].xml");
  let settings = await read(zip, "word/settings.xml");
  const existingComments = await read(zip, "word/comments.xml");
  if (!docXml) { console.error("读不到 word/document.xml，不是 docx？"); process.exit(2); }

  const plain = docXml.split("</w:p>").map(textOf).join("\n");
  // 条文体可能只把「第X章」「第X条」放在自动编号里，正文只有「总则」。
  const numbering = await read(zip, "word/numbering.xml");
  const isZhang = isZhangDoc({ text: plain, styles, document: docXml, numbering });
  const form = isZhang ? "zhang" : "wen";
  const isXinhan = /<w:titlePg\s*\/>|<w:titlePg\b/.test(docXml)
    && /<w:headerReference w:type="first"/.test(docXml) && /<w:footerReference w:type="first"/.test(docXml);
  const isJiyao = (docXml.match(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g) || [])
    .some((p) => isRed(p) && /纪要/.test(textOf(p)));

  const notes = [];
  let headingIds = {};
  let redStyles = new Set();
  let styleInfo = new Map();
  if (o.layout) {
    if ((isXinhan || isJiyao) && !o.forceLayout) {
      o.layout = false;
      notes.push(`检测到${isXinhan ? "信函" : "纪要"}格式：版式未改（用 --layout 可强制通用版式）`);
    }
  }
  if (o.layout) {
    redStyles = redStyleIds(styles);
    const s = fixStyles(styles, { form, titleFont: TITLE_FONT }, notes);
    styles = s.xml; headingIds = s.headingIds; styleInfo = s.styles;
    docXml = fixDocument(docXml, { headingIds, titleFont: TITLE_FONT, redStyles, styles: styleInfo }, notes);
    const pk = fixPackage(zip, { rels, types, settings }, notes);
    ({ rels, types, settings } = pk);
  } else {
    styleInfo = styleMap(styles);
    for (const m of styles.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?<w:name w:val="heading ([1-4])"/gi)) headingIds[m[1]] = +m[2];
  }

  // 修订与批注共用 id 空间，从文中已有的最大 id 之后开始
  // 不用 Math.max(...ids)：id 上十万个的大文件会超出参数上限
  let next = 1;
  for (const m of (docXml + existingComments).matchAll(/w:id="(\d+)"/g)) if (+m[1] >= next) next = +m[1] + 1;
  const date = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const rev = { nextId: () => next++, author: AUTHOR, date };
  const tally = {};
  const tp = textPass(docXml, styles, headingIds, rev, tally, { redStyles, styleInfo });
  docXml = tp.xml;

  // 写中间件，跑机检
  const put = () => {
    zip.file("word/document.xml", docXml);
    if (styles) zip.file("word/styles.xml", styles);
    if (rels) zip.file("word/_rels/document.xml.rels", rels);
    if (types) zip.file("[Content_Types].xml", types);
    if (settings) zip.file("word/settings.xml", settings);
  };
  put();
  const tmp = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "gongwen-fix-")), "mid.docx");
  fs.writeFileSync(tmp, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  const left = await remaining(tmp);
  fs.rmSync(path.dirname(tmp), { recursive: true, force: true });

  const cm = addComments(docXml, toComments(left.items), { existing: existingComments, startId: next, author: AUTHOR, date });
  docXml = cm.xml;
  ({ rels, types } = registerComments(rels, types));
  zip.file("word/comments.xml", cm.comments);
  put();
  if (!o.dryRun) fs.writeFileSync(o.out, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));

  // 报告
  const nEdits = Object.values(tally).reduce((a, b) => a + b, 0);
  console.log(`修复完成　${o.dryRun ? "（dry-run，未写文件）" : path.basename(o.out)}（原文件未改动）`);
  if (o.layout) {
    console.log("  版式（直接改）：");
    notes.forEach((n) => console.log("    · " + n));
  } else if (notes.length) {
    console.log("  版式（未改）：");
    notes.forEach((n) => console.log("    · " + n));
  }
  console.log(`  体例（修订痕迹，${nEdits} 处，在 Word「审阅」里逐条接受或拒绝）：`);
  Object.entries(tally).forEach(([k, v]) => console.log(`    · ${k} ${v} 处`));
  if (tp.skipped) console.log(`    · 另有 ${tp.skipped} 处碰到域、超链接或已有修订，没有自动改，已转批注`);
  console.log(`  批注 ${cm.placed + cm.unplaced} 条（数字、提示与改法要靠判断的项，不自动改）${cm.unplaced ? `，其中 ${cm.unplaced} 条没定位到原文、挂在首段` : ""}`);
  const byKind = {};
  left.items.forEach((it) => { byKind[it.kind] = (byKind[it.kind] || 0) + 1; });
  Object.entries(byKind).forEach(([k, v]) => console.log(`    · ${k} ${v} 项`));
  if (left.style.length) {
    console.log("  文风风格层（只提示，没插批注，看语境定）：");
    left.style.forEach((l) => console.log("    · " + l));
  }
})().catch((e) => {
  console.error("修复失败（工具故障，原文件未改动）：", e.message);
  process.exit(2);
});
