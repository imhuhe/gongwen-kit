"use strict";
/**
 * styles.js —— 最小样式底座（2026-10-12 立）。
 *
 * 替代原先「读 dotx 的 styles.xml 当 externalStyles」的做法：只保留生成侧
 * 真正用到的样式定义（二十个上下），值从原底座逐项核对后重新声明，Word 的
 * 修订元数据（w:rsid）一律去掉。原底座 433 个样式、686KB，这里十几 KB。
 *
 * 两种体式各有自己的一套：行文体（wen）标题下接主送机关，层次序数走
 * 一、（一）1.（1）；条文体（zhang）章标题居中、条项各有别名（| 章／条／项）。
 * 共同的部分（docDefaults、Normal、脚注字符样式、TableNormal、NoList）只写一份。
 *
 * 标题字体在模板里留「方正小标宋简体」占位，stylesFor() 按传入字体替换；
 * 占位数量对不上就抛错——字体替换失灵必须当场看见，不能静默退化成默认字体。
 */
const PLACEHOLDER = "方正小标宋简体";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const ROOT = '<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
  + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  + ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  + ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
  + ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
  + ' xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"'
  + ' xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"'
  + ' xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"'
  + ' xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"'
  + ' xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"'
  + ' xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock"'
  + ' xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"'
  + ' mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du">';

const DOC_DEFAULTS = `
<w:docDefaults>
<w:rPrDefault>
<w:rPr>
<w:rFonts w:ascii="FangSong" w:eastAsia="FangSong" w:hAnsi="FangSong" w:cstheme="minorBidi"/>
<w:kern w:val="2"/>
<w:sz w:val="32"/>
<w:szCs w:val="32"/>
<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
</w:rPr>
</w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>`;

const NORMAL = `
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
<w:name w:val="Normal"/>
<w:qFormat/>
<w:pPr>
<w:widowControl w:val="0"/>
<w:ind w:firstLineChars="200" w:firstLine="632"/>
<w:jc w:val="both"/>
</w:pPr>
<w:rPr>
<w:kern w:val="0"/>
</w:rPr>
</w:style>`;

const DEFAULT_PARAGRAPH_FONT = `
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
<w:name w:val="Default Paragraph Font"/>
<w:uiPriority w:val="1"/>
<w:semiHidden/>
<w:unhideWhenUsed/>
</w:style>`;

const FOOTNOTE_REFERENCE = `
<w:style w:type="character" w:styleId="FootnoteReference">
<w:name w:val="footnote reference"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:uiPriority w:val="39"/>
<w:semiHidden/>
<w:rPr>
<w:vertAlign w:val="superscript"/>
</w:rPr>
</w:style>`;

const ENDNOTE_REFERENCE = `
<w:style w:type="character" w:styleId="EndnoteReference">
<w:name w:val="endnote reference"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:uiPriority w:val="39"/>
<w:semiHidden/>
<w:rPr>
<w:vertAlign w:val="superscript"/>
</w:rPr>
</w:style>`;

const TABLE_NORMAL = `
<w:style w:type="table" w:default="1" w:styleId="TableNormal">
<w:name w:val="Normal Table"/>
<w:uiPriority w:val="99"/>
<w:semiHidden/>
<w:unhideWhenUsed/>
<w:tblPr>
<w:tblInd w:w="0" w:type="dxa"/>
<w:tblCellMar>
<w:top w:w="0" w:type="dxa"/>
<w:left w:w="108" w:type="dxa"/>
<w:bottom w:w="0" w:type="dxa"/>
<w:right w:w="108" w:type="dxa"/>
</w:tblCellMar>
</w:tblPr>
</w:style>`;

const NO_LIST = `
<w:style w:type="numbering" w:default="1" w:styleId="NoList">
<w:name w:val="No List"/>
<w:uiPriority w:val="99"/>
<w:semiHidden/>
<w:unhideWhenUsed/>
</w:style>`;

// ── 行文体（报告·请示·函·纪要）────────────────────────────
// 标题回行与标题下接主送机关（a0）。层次 一、（一）1.（1），黑体／楷体／仿宋。
const WEN_TITLE = `
<w:style w:type="paragraph" w:styleId="Title">
<w:name w:val="Title"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="a0"/>
<w:link w:val="TitleChar"/>
<w:uiPriority w:val="5"/>
<w:qFormat/>
<w:pPr>
<w:spacing w:afterLines="100" w:after="579" w:line="579" w:lineRule="exact"/>
<w:ind w:firstLineChars="0" w:firstLine="0"/>
<w:contextualSpacing/>
<w:jc w:val="center"/>
<w:outlineLvl w:val="0"/>
</w:pPr>
<w:rPr>
<w:rFonts w:ascii="方正小标宋简体" w:eastAsia="方正小标宋简体" w:cstheme="majorBidi"/>
<w:bCs/>
<w:sz w:val="44"/>
</w:rPr>
</w:style>`;

const WEN_ZHUSONG = `
<w:style w:type="paragraph" w:customStyle="1" w:styleId="a0">
<w:name w:val="主送机关"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="a1"/>
<w:uiPriority w:val="6"/>
<w:qFormat/>
<w:pPr>
<w:ind w:firstLineChars="0" w:firstLine="0"/>
</w:pPr>
</w:style>`;

const WEN_ZHUSONG_CHAR = `
<w:style w:type="character" w:customStyle="1" w:styleId="a1">
<w:name w:val="主送机关 字符"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="a0"/>
<w:rPr>
<w:kern w:val="0"/>
</w:rPr>
</w:style>`;

const WEN_H1 = `
<w:style w:type="paragraph" w:styleId="Heading1">
<w:name w:val="heading 1"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading1Char"/>
<w:uiPriority w:val="1"/>
<w:qFormat/>
<w:pPr>
<w:keepNext/>
<w:keepLines/>
<w:numPr>
<w:ilvl w:val="0"/>
<w:numId w:val="1"/>
</w:numPr>
<w:ind w:firstLine="632"/>
<w:outlineLvl w:val="0"/>
</w:pPr>
<w:rPr>
<w:rFonts w:ascii="SimHei" w:eastAsia="SimHei" w:hAnsi="SimHei"/>
<w:bCs/>
<w:szCs w:val="44"/>
</w:rPr>
</w:style>`;

const WEN_H2 = `
<w:style w:type="paragraph" w:styleId="Heading2">
<w:name w:val="heading 2"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading2Char"/>
<w:uiPriority w:val="2"/>
<w:qFormat/>
<w:pPr>
<w:numPr>
<w:ilvl w:val="1"/>
<w:numId w:val="1"/>
</w:numPr>
<w:ind w:firstLine="632"/>
<w:outlineLvl w:val="1"/>
</w:pPr>
<w:rPr>
<w:rFonts w:ascii="KaiTi" w:eastAsia="KaiTi" w:hAnsi="KaiTi" w:cstheme="majorBidi"/>
<w:bCs/>
</w:rPr>
</w:style>`;

const WEN_H3 = `
<w:style w:type="paragraph" w:styleId="Heading3">
<w:name w:val="heading 3"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading3Char"/>
<w:uiPriority w:val="3"/>
<w:qFormat/>
<w:pPr>
<w:numPr>
<w:ilvl w:val="2"/>
<w:numId w:val="1"/>
</w:numPr>
<w:ind w:firstLine="632"/>
<w:outlineLvl w:val="2"/>
</w:pPr>
<w:rPr>
<w:bCs/>
</w:rPr>
</w:style>`;

const WEN_H4 = `
<w:style w:type="paragraph" w:styleId="Heading4">
<w:name w:val="heading 4"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading4Char"/>
<w:uiPriority w:val="4"/>
<w:qFormat/>
<w:pPr>
<w:numPr>
<w:ilvl w:val="3"/>
<w:numId w:val="1"/>
</w:numPr>
<w:ind w:firstLine="632"/>
<w:outlineLvl w:val="3"/>
</w:pPr>
<w:rPr>
<w:rFonts w:cstheme="majorBidi"/>
<w:bCs/>
<w:szCs w:val="28"/>
</w:rPr>
</w:style>`;

// ── 条文体（管理办法·条例）────────────────────────────────
// 章标题居中（方正小标宋），条与项各带别名；条号全文连续由 numbering 补丁管。
const ZHANG_TITLE = `
<w:style w:type="paragraph" w:styleId="Title">
<w:name w:val="Title"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Heading1"/>
<w:link w:val="TitleChar"/>
<w:uiPriority w:val="5"/>
<w:qFormat/>
<w:pPr>
<w:spacing w:afterLines="100" w:after="579" w:line="579" w:lineRule="exact"/>
<w:ind w:firstLineChars="0" w:firstLine="0"/>
<w:contextualSpacing/>
<w:jc w:val="center"/>
<w:outlineLvl w:val="0"/>
</w:pPr>
<w:rPr>
<w:rFonts w:ascii="方正小标宋简体" w:eastAsia="方正小标宋简体" w:cstheme="majorBidi"/>
<w:bCs/>
<w:sz w:val="44"/>
</w:rPr>
</w:style>`;

const ZHANG_H1 = `
<w:style w:type="paragraph" w:styleId="Heading1">
<w:name w:val="heading 1"/>
<w:aliases w:val="| 章"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Heading2"/>
<w:link w:val="Heading1Char"/>
<w:uiPriority w:val="1"/>
<w:qFormat/>
<w:pPr>
<w:keepNext/>
<w:keepLines/>
<w:numPr>
<w:numId w:val="1"/>
</w:numPr>
<w:spacing w:before="359" w:after="220"/>
<w:ind w:firstLineChars="0"/>
<w:jc w:val="center"/>
<w:outlineLvl w:val="0"/>
</w:pPr>
<w:rPr>
<w:rFonts w:ascii="方正小标宋简体" w:eastAsia="方正小标宋简体"/>
<w:bCs/>
<w:szCs w:val="44"/>
</w:rPr>
</w:style>`;

const ZHANG_H2 = `
<w:style w:type="paragraph" w:styleId="Heading2">
<w:name w:val="heading 2"/>
<w:aliases w:val="| 条"/>
<w:basedOn w:val="Normal"/>
<w:link w:val="Heading2Char"/>
<w:uiPriority w:val="2"/>
<w:qFormat/>
<w:pPr>
<w:numPr>
<w:ilvl w:val="1"/>
<w:numId w:val="1"/>
</w:numPr>
<w:ind w:firstLine="632"/>
<w:outlineLvl w:val="1"/>
</w:pPr>
<w:rPr>
<w:rFonts w:cstheme="majorBidi"/>
<w:bCs/>
</w:rPr>
</w:style>`;

const ZHANG_H3 = `
<w:style w:type="paragraph" w:styleId="Heading3">
<w:name w:val="heading 3"/>
<w:aliases w:val="| 项"/>
<w:basedOn w:val="Normal"/>
<w:link w:val="Heading3Char"/>
<w:uiPriority w:val="3"/>
<w:qFormat/>
<w:pPr>
<w:numPr>
<w:ilvl w:val="2"/>
<w:numId w:val="1"/>
</w:numPr>
<w:ind w:firstLine="632"/>
<w:outlineLvl w:val="2"/>
</w:pPr>
<w:rPr>
<w:bCs/>
</w:rPr>
</w:style>`;

const ZHANG_H4 = `
<w:style w:type="paragraph" w:styleId="Heading4">
<w:name w:val="heading 4"/>
<w:basedOn w:val="Normal"/>
<w:link w:val="Heading4Char"/>
<w:uiPriority w:val="4"/>
<w:qFormat/>
<w:pPr>
<w:numPr>
<w:ilvl w:val="3"/>
<w:numId w:val="1"/>
</w:numPr>
<w:ind w:firstLine="632"/>
<w:outlineLvl w:val="3"/>
</w:pPr>
<w:rPr>
<w:rFonts w:cstheme="majorBidi"/>
<w:bCs/>
<w:szCs w:val="28"/>
</w:rPr>
</w:style>`;

// ── 字符样式：Title 与四层标题各一条（与段落样式互链）─────────
const TITLE_CHAR = `
<w:style w:type="character" w:customStyle="1" w:styleId="TitleChar">
<w:name w:val="Title Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Title"/>
<w:uiPriority w:val="10"/>
<w:rPr>
<w:rFonts w:ascii="方正小标宋简体" w:eastAsia="方正小标宋简体" w:cstheme="majorBidi"/>
<w:bCs/>
<w:kern w:val="0"/>
<w:sz w:val="44"/>
</w:rPr>
</w:style>`;

const headingChar = (n, rPr) => `
<w:style w:type="character" w:customStyle="1" w:styleId="Heading${n}Char">
<w:name w:val="Heading ${n} Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading${n}"/>
<w:uiPriority w:val="9"/>
<w:semiHidden/>
<w:rPr>${rPr}</w:rPr>
</w:style>`;

const WEN_HEADING_CHARS = [
  headingChar(1, `
<w:rFonts w:ascii="SimHei" w:eastAsia="SimHei" w:hAnsi="SimHei"/>
<w:bCs/>
<w:kern w:val="0"/>
<w:szCs w:val="44"/>`),
  headingChar(2, `
<w:rFonts w:ascii="KaiTi" w:eastAsia="KaiTi" w:hAnsi="KaiTi" w:cstheme="majorBidi"/>
<w:bCs/>
<w:kern w:val="0"/>`),
  headingChar(3, `
<w:bCs/>
<w:kern w:val="0"/>`),
  headingChar(4, `
<w:rFonts w:cstheme="majorBidi"/>
<w:bCs/>
<w:kern w:val="0"/>
<w:szCs w:val="28"/>`),
];

const ZHANG_H1_CHAR = `
<w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char">
<w:name w:val="Heading 1 Char"/>
<w:aliases w:val="| 章 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading1"/>
<w:uiPriority w:val="9"/>
<w:semiHidden/>
<w:rPr>
<w:rFonts w:ascii="方正小标宋简体" w:eastAsia="方正小标宋简体"/>
<w:bCs/>
<w:kern w:val="0"/>
<w:szCs w:val="44"/>
</w:rPr>
</w:style>`;

const ZHANG_H2_CHAR = `
<w:style w:type="character" w:customStyle="1" w:styleId="Heading2Char">
<w:name w:val="Heading 2 Char"/>
<w:aliases w:val="| 条 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading2"/>
<w:uiPriority w:val="9"/>
<w:semiHidden/>
<w:rPr>
<w:rFonts w:cstheme="majorBidi"/>
<w:bCs/>
<w:kern w:val="0"/>
</w:rPr>
</w:style>`;

const ZHANG_H3_CHAR = `
<w:style w:type="character" w:customStyle="1" w:styleId="Heading3Char">
<w:name w:val="Heading 3 Char"/>
<w:aliases w:val="| 项 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading3"/>
<w:uiPriority w:val="9"/>
<w:semiHidden/>
<w:rPr>
<w:bCs/>
<w:kern w:val="0"/>
</w:rPr>
</w:style>`;

const ZHANG_H4_CHAR = `
<w:style w:type="character" w:customStyle="1" w:styleId="Heading4Char">
<w:name w:val="Heading 4 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading4"/>
<w:uiPriority w:val="9"/>
<w:semiHidden/>
<w:rPr>
<w:rFonts w:cstheme="majorBidi"/>
<w:bCs/>
<w:kern w:val="0"/>
<w:szCs w:val="28"/>
</w:rPr>
</w:style>`;

const SHARED_HEAD = [DOC_DEFAULTS, NORMAL, DEFAULT_PARAGRAPH_FONT];
const SHARED_TAIL = [FOOTNOTE_REFERENCE, ENDNOTE_REFERENCE, TABLE_NORMAL, NO_LIST];
// 标题字体占位在每个体式里应出现的次数（每条 rFonts 行 ascii＋eastAsia 两处）：
// 行文体 Title＋TitleChar 两个样式共 4 处；条文体再加章标题与其字符样式，共 8 处。
// 对不上就是模板被改坏，字体替换会不完整。
const PLACEHOLDER_COUNT = { wen: 4, zhang: 8 };

const FORMS = {
  wen: [...SHARED_HEAD, WEN_TITLE, TITLE_CHAR, WEN_ZHUSONG, WEN_ZHUSONG_CHAR,
    WEN_H1, ...WEN_HEADING_CHARS.slice(0, 1), WEN_H2, WEN_HEADING_CHARS[1],
    WEN_H3, WEN_HEADING_CHARS[2], WEN_H4, WEN_HEADING_CHARS[3], ...SHARED_TAIL],
  zhang: [...SHARED_HEAD, ZHANG_TITLE, TITLE_CHAR,
    ZHANG_H1, ZHANG_H1_CHAR, ZHANG_H2, ZHANG_H2_CHAR,
    ZHANG_H3, ZHANG_H3_CHAR, ZHANG_H4, ZHANG_H4_CHAR, ...SHARED_TAIL],
};

/** 生成某个体式的最小 styles.xml，并把标题字体占位替换成实际字体。
 *  占位数量对不上就抛错——静默的话字体替换会失灵而产出看着一切正常。 */
function stylesFor(form, titleFont) {
  const blocks = FORMS[form];
  if (!blocks) throw new Error(`体式只有 ${Object.keys(FORMS).join("／")}，给的是「${form}」`);
  const xml = XML_DECL + ROOT + blocks.join("") + "\n</w:styles>";
  const n = xml.split(PLACEHOLDER).length - 1;
  const want = PLACEHOLDER_COUNT[form];
  if (n !== want) {
    throw new Error(`样式模板里「${PLACEHOLDER}」应出现 ${want} 处，实际 ${n} 处——`
      + "模板被动过，字体替换不完整");
  }
  return xml.replaceAll(PLACEHOLDER, titleFont);
}

module.exports = { stylesFor, PLACEHOLDER };
