#!/usr/bin/env node
/**
 * gongwen-core.js —— 对外·行文档（公文体）的构件库。
 *
 * 党政机关公文体（GB/T 9704）。无封面、无目录、无色彩，版心与行字数
 * 是硬指标。
 *
 * 两种体式，用哪种在稿源脚本第一行就声明：
 *
 *   form("wen")    行文体。报告·请示·函·纪要。层次 一、／（一）／1.／（1）
 *   form("zhang")  条文体。管理办法·工作条例。层次 第X章／第X条／（一）／1.
 *
 * 样式底座由 lib/styles.js 生成最小 styles.xml 当 externalStyles 用
 * （2026-10-12 起不再读 dotx）。这几项都是国标硬指标，值与原底座逐项核对过：
 *
 *   版心   上 37 / 下 35 / 左 28 / 右 26 mm            GB/T 9704 8.1
 *   行字   每页 22 行、每行 28 字                       GB/T 9704 8.1
 *          （docGrid linePitch 579 + charSpace -842）
 *   正文   仿宋 三号（sz 32）                           GB/T 9704 7.3.4
 *   标题   方正小标宋简体 二号（sz 44）                 GB/T 9704 7.3.1
 *   层次   四层，两种体式各一套                       GB/T 9704 7.3.3
 *   页码   — N — 居中，首页不同、单双页分设            GB/T 9704 8.4
 *
 * 编号不吃模板 styles 里那份 numPr（它指向模板自己的 numId=1，我们换了
 * numbering 部件之后那个号未必还在），一律在段落上显式声明，见 heading()。
 *
 * 这一档此前有一版手拼 word/document.xml 塞回 dotx 的做法，没有机检；
 * 现在正文全部走构件生成，写盘后自动跑机检。
 */
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Paragraph, TextRun, Tab, AlignmentType, HeadingLevel, LevelFormat,
  Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign,
  Header, Footer, PageNumber, SectionType, LevelSuffix,
} = require("docx");

// ── 样式底座 ──────────────────────────────────────────────
const ROOT = path.dirname(__dirname);   // 仓根：本文件在 lib/
const { stylesFor } = require("./styles");
const { runChecks } = require("./run-checks");
const { letterhead: letterheadOf } = require("./letterhead");
const { banji: banjiOf } = require("./banji");
const { xinhanSection: xinhanSectionOf, xinhanHead } = require("./xinhan");

// 标题与发文机关标志字体。GB/T 9704 7.3.1：标题用 2 号小标宋体字；
// 7.2.4：发文机关标志推荐使用小标宋体字。实际字体用「方正小标宋简体」——
// 公文环境的通行实现，政府/央国企的机器装了它。这里只声明字体名，不分发
// 字体文件（方正字库个人授权不含商业使用，买断要花钱；声明名称不涉分发）。
// 机器上没有这个字体时 Word 会替换；只想本机预览的用环境变量覆盖：
//   GONGWEN_FONT_TITLE="Songti SC Black" node 稿源.js
// （macOS 自带、宋体骨架、最接近小标宋；Windows 上会替换成宋体，仍是宋体骨架。）
// 字体策略详见 docs/字体.md。
const TITLE_FONT = process.env.GONGWEN_FONT_TITLE || "方正小标宋简体";

// externalStyles 不再从 dotx 抽取：由 styles.js 的 stylesFor(form, TITLE_FONT)
// 生成最小底座。numbering 与 sectPr 不吃样式底座，在下面自己声明。

// ── 版面 ──────────────────────────────────────────────────
const A4 = { width: 11906, height: 16838 };
// 版心直接抄模板 sectPr 的 twips 值，不走 convertMillimetersToTwip——毫米换算
// 取整会差 1 twip（37 mm 算出来是 2097，模板写的是 2098），差值没有实际意义，
// 但两个数不一样就得有人去判断哪个算数。抄权威值，省掉这个判断。
// footer=1392 是页码位置的量测校准值：让一字线上缘落在版心下边缘之下 7mm
// （GB/T 9704 7.5），反推自渲染 PDF 的着墨位置。
const MARGIN = { top: 2098, bottom: 1984, left: 1587, right: 1474, header: 850, footer: 1392 };
// 每页 22 行、每行 28 字。这两个数不是审美，是 GB/T 9704 8.1 的规定值，
// 抄自模板 sectPr——改它等于改成非公文版式。
const GRID = { type: "linesAndChars", linePitch: 579, charSpace: -842 };

const IND = 632;   // 首行缩进 2 字（三号字 2 × 316 twips）

// ── 层次序数：一、／（一）／1.／（1） ─────────────────────
// ── 两种体式 ─────────────────────────────────────────────
// 编号与标题之间不留制表位（suff nothing，抄模板）。缺了它 Word 会跳到
// 下一个制表位，出样是「一、␣␣要素与顺序」，空出两个字。
// run 参数给编号层单独指定字体：条文体「第X条」几个字要黑体，而条的内容
// 是仿宋——内容字体来自段落样式，编号字体只能挂在编号层的 rPr 上
// （2026-10-12 对照参考模板核实：这是 Word 多级列表的标准做法）。
const lvl = (level, format, text, run) => ({
  level, format, text, alignment: AlignmentType.LEFT, suffix: LevelSuffix.NOTHING,
  style: { paragraph: { indent: { left: 0, firstLine: IND } }, ...(run ? { run } : {}) },
});
const CN = LevelFormat.CHINESE_COUNTING_THOUSAND;
const NUM = LevelFormat.DECIMAL;
const HEI = { font: { eastAsia: "SimHei", ascii: "SimHei", hAnsi: "SimHei" } };

const FORMS = {
  // 行文体：报告·请示·函·纪要。GB/T 9704 7.3.3 的四层。
  wen: {
    levels: [lvl(0, CN, "%1、"), lvl(1, CN, "（%2）"), lvl(2, NUM, "%3."), lvl(3, NUM, "（%4）")],
  },
  // 条文体：管理办法·工作条例。章条项款。
  // 第一层是「第X章」，模板的 Heading1 样式给它居中，构件不用管。
  zhang: {
    levels: [lvl(0, CN, "第%1章  "), lvl(1, CN, "第%2条  ", HEI), lvl(2, CN, "（%3）"), lvl(3, NUM, "%4.")],
    // 章号与条号全文连续，不按上级重启——「第二章 第一条」是错的，管理办法
    // 的条号从第一条排到最后一条。项号「（一）」按条重启，那是对的。
    // docx 库的 level 选项里没有 lvlRestart，走写盘后的补丁，见 patchNumbering。
    noRestart: ["第%1章", "第%2条"],
  },
};

// ── 构件 ──────────────────────────────────────────────────
const esc = (s) => String(s);

/** 标题。多行时逐行给，回行处按意群断，不按宽度断（GB/T 9704 7.3.1）。 */
function title(lines) {
  const arr = Array.isArray(lines) ? lines : [lines];
  return arr.map((ln, i) => new Paragraph({
    style: "Title",
    ...(i < arr.length - 1 ? { spacing: { after: 0, line: 579, lineRule: "exact" } } : {}),
    children: [new TextRun(esc(ln))],
  }));
}

/** 主送机关。顶格，末尾冒号。函与纪要可以没有。 */
const zhusong = (name) => new Paragraph({
  indent: { firstLine: 0, firstLineChars: 0 },
  children: [new TextRun(name.endsWith("：") ? name : name + "：")],
});

/** 正文段。首行缩进 2 字由 Normal 样式给，这里不重复声明。 */
const p = (text) => new Paragraph({ children: [new TextRun(esc(text))] });

/** 顶格段（不缩进），用于附件说明、抄送这类要求左空的行。 */
const flush = (text, { align } = {}) => new Paragraph({
  indent: { firstLine: 0, firstLineChars: 0 },
  ...(align ? { alignment: align } : {}),
  children: [new TextRun(esc(text))],
});

const HEADING_STYLE = ["Heading1", "Heading2", "Heading3", "Heading4"];

/** 层次标题。level 0..3 的编号形式由体式决定，见 FORMS。
 *  字体由模板样式给，编号由本文件给。 */
function makeHeading(ref) {
  return function heading(level, text) {
    if (level < 0 || level > 3) {
      throw new Error(`公文层次只有四层（GB/T 9704 7.3.3），给的是第 ${level + 1} 层`);
    }
    return new Paragraph({
      style: HEADING_STYLE[level],
      numbering: { reference: ref, level },
      children: [new TextRun(esc(text))],
    });
  };
}

// ── 三线表 ────────────────────────────────────────────────
const NONE = { style: BorderStyle.NONE, size: 0, color: "auto" };
const RULE = (sz) => ({ style: BorderStyle.SINGLE, size: sz, color: "000000" });

/** 表题。四号黑体居中，置于表上方（GB/T 1.1 表的编号与名称在表上方）。
 *  keepNext 保证表题不孤悬页底、跟表走。 */
const caption = (text) => new Paragraph({
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0, firstLineChars: 0 },
  keepNext: true,
  children: [new TextRun({ text: esc(text), font: { eastAsia: "SimHei", ascii: "SimHei" }, size: 28 })],
});

function table({ head, rows, widths, align, caption: cap }) {
  const cell = (text, w, isHead, jc) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    ...(isHead ? { borders: { top: NONE, left: NONE, right: NONE, bottom: RULE(6) } } : {}),
    children: [new Paragraph({
      alignment: jc,
      indent: { firstLine: 0, firstLineChars: 0 },
      spacing: { line: 360, lineRule: "exact" },
      children: [new TextRun({
        text: esc(text), size: 28,
        ...(isHead ? { font: { eastAsia: "SimHei", ascii: "SimHei" } } : {}),
      })],
    })],
  });
  const jcs = align || widths.map(() => AlignmentType.CENTER);
  const tbl = new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    borders: { top: RULE(12), bottom: RULE(12), left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    rows: [
      new TableRow({
        tableHeader: true,
        children: head.map((t, j) => cell(t, widths[j], true, AlignmentType.CENTER)),
      }),
      ...rows.map((r) => new TableRow({
        children: r.map((t, j) => cell(t, widths[j], false, jcs[j])),
      })),
    ],
  });
  return [...(cap ? [caption(cap)] : []), tbl, new Paragraph({ children: [] })];
}

// ── 纪要名单 ──────────────────────────────────────────────
/** 纪要的出席／请假／列席名单（GB/T 9704 10.3）。
 *  在正文或附件说明下空一行、左空二字编排；「出席」等标签与冒号用 3 号黑体，
 *  名单用 3 号仿宋；回行与冒号后的首字对齐。
 *  rows：[["出席", "×××、×××"], ["请假", "×××"], ["列席", "×××"]] */
function attendees(rows) {
  const ONE = IND / 2;   // 三号字一个字宽
  const out = [new Paragraph({ children: [] })];   // 空一行
  for (const [label, names] of rows) {
    const lab = String(label) + "：";
    const w = [...lab].length;
    out.push(new Paragraph({
      // 首行落在 2 字位；回行落在 2＋标签字数位（冒号后的首字）
      indent: { left: ONE * (2 + w), hanging: ONE * w, firstLine: 0, firstLineChars: 0 },
      children: [
        new TextRun({ text: lab, font: { eastAsia: "SimHei", ascii: "SimHei", hAnsi: "SimHei" } }),
        new TextRun(esc(names)),
      ],
    }));
  }
  return out;
}

// ── 落款 ──────────────────────────────────────────────────
/** 发文机关署名与成文日期。两种形态（GB/T 9704 7.3.5.1／7.3.5.2）：
 *  默认加盖印章：成文日期右空四字；单一机关时，发文机关署名以成文日期为准
 *  居中（7.3.5.1）。联合行文时各署名按顺序右空四字整齐排列。
 *  无公章：署名右空二字；日期首字比署名首字右移二字——日期更长时按
 *  7.3.5.2 改为日期右空二字、相应增加署名右空。字宽按全角 1、半角 0.5 估。
 *  署名多行时逐行给（联合行文）。日期 "2026 年 9 月 10 日" 这种写法里
 *  的数字与空格是半角，估宽会低一点，按条款关系换算右空字数。 */
function luokuan({ unit, date, sealed = true }) {
  const units = Array.isArray(unit) ? unit : [unit];
  const FIRST = { firstLine: 0, firstLineChars: 0 };
  const w = (s) => [...String(s)].reduce((a, ch) => a + (/[\x00-\xff]/.test(ch) ? 0.5 : 1), 0);
  const dateW = w(date);
  if (sealed) {
    const datePad = { ...FIRST, right: 316 * 4 };
    if (units.length === 1) {
      // 署名与日期右缘不齐：署名以日期为中心，右空 = 4 + (日期宽 - 署名宽)/2。
      const unitRight = Math.max(0, 4 + (dateW - w(units[0])) / 2);
      const unitPad = { ...FIRST, right: Math.round(316 * unitRight) };
      return [
        new Paragraph({ children: [] }),
        new Paragraph({
          alignment: AlignmentType.RIGHT, indent: unitPad,
          children: [new TextRun(esc(units[0]))],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT, indent: datePad,
          children: [new TextRun(esc(date))],
        }),
      ];
    }
    return [
      new Paragraph({ children: [] }),
      ...units.map((u) => new Paragraph({
        alignment: AlignmentType.RIGHT, indent: datePad,
        children: [new TextRun(esc(u))],
      })),
      new Paragraph({
        alignment: AlignmentType.RIGHT, indent: datePad,
        children: [new TextRun(esc(date))],
      }),
    ];
  }
  const unitW = Math.max(...units.map(w));
  const unitRight = Math.max(2, 4 + dateW - unitW);          // 字
  const dateRight = unitRight + unitW - 2 - dateW;           // 字
  return [
    new Paragraph({ children: [] }),
    ...units.map((u) => new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { ...FIRST, right: Math.round(316 * unitRight) },
      children: [new TextRun(esc(u))],
    })),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { ...FIRST, right: Math.round(316 * Math.max(0, dateRight)) },
      children: [new TextRun(esc(date))],
    }),
  ];
}

/** 附件说明。正文下空一行，左空二字（GB/T 9704 7.3.5.4）。多个附件用阿拉伯
 *  数字标顺序号（「附件：1.××××」），其余各行的序号与第一行附件名称首字对齐；
 *  附件名称后不加标点。空行由本构件给，调用方不用自己补——补重了就是空两行。 */
function fujian(items) {
  const arr = Array.isArray(items) ? items : [items];
  const LABEL = 316 * 3;   // 「附件：」三个三号字宽
  return [
    new Paragraph({ children: [] }),
    ...arr.map((it, i) => new Paragraph({
      indent: i === 0
        ? { firstLine: IND, firstLineChars: 200 }
        : { firstLine: IND + LABEL, firstLineChars: 200 + 300 },
      children: [new TextRun(i === 0
        ? `附件：${arr.length > 1 ? "1." : ""}${it}`
        : `${i + 1}.${it}`)],
    })),
  ];
}

/** 附注。成文日期下一行，居左空二字加圆括号（GB/T 9704 7.3.6）。
 *  惯例写法：「（联系人：××，联系电话：×××。）」——括注连同括号由调用方给全。 */
const fuzhu = (text) => new Paragraph({
  indent: { firstLine: IND, firstLineChars: 200 },
  children: [new TextRun(esc(text))],
});

/** 附件页（GB/T 9704 7.3.7）：附件应当另面编排（正文之后、版记之前），与
 *  正文一起装订。「附件」二字及顺序号用 3 号黑体顶格编排在版心左上角第一行，
 *  附件标题居中编排在版心第三行（小标宋三号，比正文标题小），其余格式同正文；
 *  页码与正文连续（由页脚统一给，不用管）。只有一个附件时不编顺序号，只写
 *  「附件」二字；不与正文一起装订的附件，第一行是「文号附件X」，传 detached
 *  发文机关代字〔年份〕序号号。第一段自带分页（另面编排）。
 *  用法：...fujianPage({ n: 1, title: "标题", body: [p("内容")] }) */
function fujianPage(o = {}) {
  const { n, title: t, detached, body = [] } = o;
  if (!t) throw new Error("附件页要有附件标题（title）");
  const mark = (detached ? String(detached) : "") + "附件" + (n != null ? n : "");
  return [
    new Paragraph({
      pageBreakBefore: true,
      indent: { firstLine: 0, firstLineChars: 0 },
      spacing: { line: 579, lineRule: "exact" },
      children: [new TextRun({
        text: mark,
        font: { eastAsia: "SimHei", ascii: "SimHei", hAnsi: "SimHei" },
        size: 32,
      })],
    }),
    new Paragraph({ children: [] }),   // 版心第二行留空，标题落在第三行
    ...(Array.isArray(t) ? t : [t]).map((ln) => new Paragraph({
      alignment: AlignmentType.CENTER,
      indent: { firstLine: 0, firstLineChars: 0 },
      children: [new TextRun({
        text: esc(ln),
        font: { eastAsia: TITLE_FONT, ascii: TITLE_FONT, hAnsi: TITLE_FONT },
        size: 32,
      })],
    })),
    ...body,
  ];
}

// ── 页脚：— N —，单双页分设 ───────────────────────────────
// 页码（GB/T 9704 7.5）：4 号半角宋体阿拉伯数字，数字左右各放一条一字线；
// 单页码居右空一字、双页码居左空一字（一字＝4 号的一个全角宽，280tw）。
const PAGE_FONT = { eastAsia: "SimSun", ascii: "SimSun", hAnsi: "SimSun" };
const pageNo = (align, side) => new Footer({
  children: [new Paragraph({
    alignment: align,
    indent: { firstLine: 0, firstLineChars: 0, ...(side === "right" ? { right: 280 } : { left: 280 }) },
    children: [
      new TextRun({ text: "— ", size: 28, font: PAGE_FONT }),
      new TextRun({ children: [PageNumber.CURRENT], size: 28, font: PAGE_FONT }),
      new TextRun({ text: " —", size: 28, font: PAGE_FONT }),
    ],
  })],
});

/** 一节的属性。页码单页在右下、双页在左下（GB/T 9704 8.4），首页照印——
 *  模板的 first 页脚是一条「总页数 > 1 才印」的域，单页文件才不印；我们的
 *  行文都不止一页，直接印。 */
function sectionProps() {
  return {
    properties: { page: { size: A4, margin: MARGIN }, grid: GRID },
    footers: {
      default: pageNo(AlignmentType.RIGHT, "right"),
      even: pageNo(AlignmentType.LEFT, "left"),
    },
  };
}

// ── settings 补丁 ─────────────────────────────────────────
// docx 库不写 characterSpacingControl，也不写模板 compat 里那几个中日韩
// 排版项。缺了标点压缩，同样 28 字一行的断行位置会跟模板差一个字。
// 这几项是版式的一部分，写盘后补回去。补丁与重打包走 jszip，不调
// unzip/zip 命令——Windows 上没有那两个命令，这一步必须跨平台。
const SETTINGS_FE = "<w:characterSpacingControl w:val=\"compressPunctuation\"/>";
const COMPAT_FE = "<w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/>"
  + "<w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/>"
  + "<w:adjustLineHeightInTable/><w:useFELayout/>";

async function patchSettings(file) {
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  let x = await zip.file("word/settings.xml").async("string");
  // 锚点找不到就抛错，不静默跳过——静默的话版式差异会一直在，
  // 而产出看着一切正常。
  if (!/<w:compat>/.test(x)) throw new Error("settings.xml 里没有 w:compat，docx 库的结构变了");
  if (!x.includes("useFELayout")) x = x.replace("<w:compat>", "<w:compat>" + COMPAT_FE);
  if (!x.includes("characterSpacingControl")) x = x.replace("<w:compat>", SETTINGS_FE + "<w:compat>");
  zip.file("word/settings.xml", x);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

/** 给指定层次加 lvlRestart 0（编号不按上级重启）。docx 库的 level 选项里
 *  没有这一项，只能写盘后补。按 lvlText 前缀定位，不按 ilvl——reference
 *  映射到哪个 abstractNumId 由库决定，不该靠猜。补丁与重打包走 jszip。 */
async function patchNumbering(file, prefixes) {
  if (!prefixes || !prefixes.length) return;
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  let x = await zip.file("word/numbering.xml").async("string");
  let n = 0;
  x = x.replace(/<w:lvl [^>]*>[\s\S]*?<\/w:lvl>/g, (blk) => {
    const t = (blk.match(/<w:lvlText w:val="([^"]*)"/) || [])[1] || "";
    if (!prefixes.some((p) => t.startsWith(p))) return blk;
    if (/<w:lvlRestart/.test(blk)) return blk;
    n += 1;
    // lvlRestart 在 schema 里排在 numFmt 之后，位置不对 Word 会当整份
    // numbering 无效——不是报错，是编号整个不显示。
    return blk.replace(/(<w:numFmt[^>]*\/>)/, '$1<w:lvlRestart w:val="0"/>');
  });
  if (n !== prefixes.length) {
    throw new Error(`lvlRestart 该补 ${prefixes.length} 层，实际补了 ${n} 层——`
      + `lvlText 对不上（要找的是 ${prefixes.join("、")}）`);
  }
  zip.file("word/numbering.xml", x);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// ── 出稿：写盘后按公文档跑机检 ────────────────────────────
async function writeDoc(doc, out, { strict = true, quiet = false, noRestart } = {}) {
  // 结构检查两层：打包前的文档树、写盘后的 XML 部件（实现见 doc-integrity）。
  const { assertDocTree, checkXmlParts } = require("./doc-integrity");
  assertDocTree(doc);
  const { Packer } = require("docx");
  const buf = await Packer.toBuffer(doc);
  fs.mkdirSync(path.dirname(path.resolve(out)), { recursive: true });
  fs.writeFileSync(out, buf);
  await patchSettings(out);
  await patchNumbering(out, noRestart);
  if (!quiet) console.log("→", out, `(${(fs.statSync(out).size / 1024).toFixed(0)} KB)`);

  // 查补丁之后的最终文件：patchSettings／patchNumbering 自己改 XML，也在查的范围里。
  const xmlProblems = await checkXmlParts(fs.readFileSync(out));
  if (xmlProblems.length) {
    console.error("✗ XML 结构：以下部件不是 well-formed XML，Word 打不开这份文件：");
    for (const p of xmlProblems) console.error(`  ${p.part} ${p.at}：${p.message}\n    附近：…${p.near}…`);
    process.exitCode = 1;
    return { out, size: buf.length, checked: false, ok: false };
  }

  const r = runChecks(out, strict);
  if (r.fault) {
    process.exitCode = 2;
    return { out, size: buf.length, checked: false, ok: false };
  }
  if (!r.ok && strict) process.exitCode = 1;
  return { out, size: buf.length, checked: true, ok: r.ok };
}

/** 稿源脚本可能住在仓库之外的项目目录，那里没有 node_modules，各自
 *  require("docx") 会解析到别处，同一进程里就有了两份 docx 模块——段落
 *  序列化成 <rootKey>w:p</rootKey>，产出的 docx 打开是空的。
 *  一律从这里拿，只有一份实例。 */
function require_docx() {
  return require("docx");
}

// ── 体式工厂 ─────────────────────────────────────────────
/** 取一套构件。稿源脚本第一行就声明用哪种体式，往后拿到的 h1～h4
 *  自带对应的编号形式，不会拿错。
 *
 *    const G = require(".../gongwen-core").form("wen");
 *    const { externalStyles, numbering, h1, h2, p, writeDoc } = G;
 *
 *  只有一份 docx 模块实例：稿源脚本可能在仓库之外的项目目录，那里没有
 *  node_modules，自己 require("docx") 会解析到别处，两份实例混用会把段落
 *  序列化成 <rootKey>w:p</rootKey>，产出的 docx 打开是空的。用 G.require_docx()。 */
function form(name = "wen") {
  const f = FORMS[name];
  if (!f) throw new Error(`体式只有 ${Object.keys(FORMS).join("／")}，给的是「${name}」`);
  const ref = "gongwen-" + name;
  const heading = makeHeading(ref);
  return {
    form: name, require_docx, patchSettings,
    externalStyles: () => stylesFor(name, TITLE_FONT),
    numbering: { config: [{ reference: ref, levels: f.levels }] },
    sectionProps,
    // 体式自带的 noRestart 绑进去，稿源脚本不用知道有这回事
    writeDoc: (doc, out, opts = {}) => writeDoc(doc, out, { noRestart: f.noRestart, ...opts }),
    A4, MARGIN, GRID, IND, TITLE_FONT,
    title, zhusong, p, flush,
    heading,
    h1: (t) => heading(0, t), h2: (t) => heading(1, t),
    h3: (t) => heading(2, t), h4: (t) => heading(3, t),
    letterhead: (o = {}) => letterheadOf({ markFont: TITLE_FONT, ...o }),
    banji: (o = {}) => banjiOf(o),
    // 信函格式：节属性（首页页眉页脚）＋首页正文开头
    xinhanSection: (o = {}) => xinhanSectionOf(sectionProps(), { markFont: TITLE_FONT, ...o }),
    xinhanHead,
    table, caption, luokuan, fujian, fuzhu, fujianPage, attendees,
    // 纪要标志：「×××纪要」居中、红色小标宋、上边缘距版心 35mm——与发文机关标志同位，
    // 复用版头构件（不加「文件」）；期号放在发文字号的位置
    jiyaoHead: ({ mark, issue = "", ...rest } = {}) => {
      if (!mark) throw new Error("纪要要有标志：mark（如「××市人民政府常务会议纪要」）");
      return letterheadOf({ markFont: TITLE_FONT, ...rest, withFile: false, agency: mark, docNo: issue });
    },
    AlignmentType, SectionType,
  };
}

module.exports = { form, FORMS, require_docx, patchSettings, patchNumbering, TITLE_FONT, runChecks };
