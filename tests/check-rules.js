#!/usr/bin/env node
/**
 * check-rules.js —— 机检规则的独立负/正对照（2026-10-23 立）。
 *
 * 为什么单独一份：README 说「每条规则都有必错标本」。原有的 negative.js
 * 覆盖了核心规则，但 check-doc 还有一批规则没有负对照；修复工具的口径
 * 变更（hint/blocking、vMerge、表格属性、书名号）也需要独立验收。
 *
 * 每个 case 只断言：报不报（blocking / hint / pass）、退出码、以及规则名是否
 * 出现在输出里。规则实现内部怎么判不在这里复制。
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const JSZip = require("jszip");
const docx = require("docx");
const G = require("../lib/gongwen-core").form("wen");

const ROOT = path.resolve(__dirname, "..");
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "gongwen-rules-"));
process.on("exit", () => fs.rmSync(TMP, { recursive: true, force: true }));
let fails = 0;
const ok = (cond, msg) => { console.log(`  ${cond ? "✓" : "✗"} ${msg}`); if (!cond) fails++; };
const P = (s, o = {}) => new docx.Paragraph({ children: [new docx.TextRun(s)], ...o });

async function build(name, children, props) {
  const d = new docx.Document({
    externalStyles: G.externalStyles(),
    numbering: G.numbering,
    sections: [{
      properties: props || {
        page: { size: { width: 11906, height: 16838 }, margin: G.MARGIN },
        grid: G.GRID,
      },
      children,
    }],
  });
  const f = path.join(TMP, name);
  fs.writeFileSync(f, await docx.Packer.toBuffer(d));
  return f;
}

const run = (checker, file, args = []) => spawnSync(process.execPath,
  [path.join(ROOT, "checks", checker), file, "--strict", ...args], { encoding: "utf8" });

function expect(label, checker, file, mode, rule) {
  const r = run(checker, file);
  const out = (r.stdout || "") + (r.stderr || "");
  if (mode === "fail") {
    ok(r.status === 1 && out.includes(rule), `${label}：${checker} 应拦稿并报出「${rule}」`);
  } else if (mode === "hint") {
    ok(r.status === 0 && out.includes("提示") && out.includes(rule), `${label}：${checker} 应只提示「${rule}」、不拦稿`);
  } else {
    ok(r.status === 0 && !out.includes("提示"), `${label}：${checker} 不应报出`);
  }
}

async function withRootKey(file) {
  const z = await JSZip.loadAsync(fs.readFileSync(file));
  const x = await z.file("word/document.xml").async("string");
  z.file("word/document.xml", x.replace("<w:body>", "<w:body><rootKey>"));
  const out = path.join(TMP, "rootkey.docx");
  fs.writeFileSync(out, await z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return out;
}

async function withChineseHeading(file) {
  const z = await JSZip.loadAsync(fs.readFileSync(file));
  let x = await z.file("word/document.xml").async("string");
  x = x.replace(/w:pStyle w:val="Heading1"/g, 'w:pStyle w:val="1"')
    .replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, "");
  z.file("word/document.xml", x);
  let st = await z.file("word/styles.xml").async("string");
  st = st.replace(/w:styleId="Heading1"/g, 'w:styleId="1"');
  z.file("word/styles.xml", st);
  const out = path.join(TMP, "cnheading.docx");
  fs.writeFileSync(out, await z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return out;
}

async function withTableAttr(file) {
  const z = await JSZip.loadAsync(fs.readFileSync(file));
  const x = await z.file("word/document.xml").async("string");
  z.file("word/document.xml", x.replace("<w:tbl>", '<w:tbl w:rsidR="00A1B2C3">'));
  const out = path.join(TMP, "tblattr.docx");
  fs.writeFileSync(out, await z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return out;
}

(async () => {
  console.log("== 体例规则 ==");
  expect("空文档", "check-doc.js", await build("empty.docx", []), "fail", "文档无正文");
  expect("rootKey", "check-doc.js", await withRootKey(await build("rk.docx", [P("正文。")])), "fail", "段落序列化坏了");
  expect("自称", "check-doc.js", await build("wo.docx", [P("我司拟于近期开展工作。")]), "hint", "自称");
  expect("称对方", "check-doc.js", await build("gui.docx", [P("请贵司予以支持。")]), "hint", "称对方");
  expect("数字与单位", "check-doc.js", await build("unit.docx", [P("容量 100kWh。")]), "fail", "数字与单位");
  expect("数值范围", "check-doc.js", await build("range.docx", [P("增长 10-20%。")]), "fail", "数值范围");
  expect("纯小数", "check-doc.js", await build("dec.docx", [P("厚度.5毫米。")]), "fail", "纯小数");
  expect("年份简写", "check-doc.js", await build("year.docx", [P("26年初完成。")]), "fail", "年份简写");
  expect("半角括号", "check-doc.js", await build("paren.docx", [P("各县(市、区)参加。")]), "fail", "半角括号");
  expect("斜体", "check-doc.js", await build("ital.docx", [new docx.Paragraph({
    children: [new docx.TextRun({ text: "斜体", italics: true })],
  })]), "hint", "斜体");
  expect("全角数字", "check-doc.js", await build("full.docx", [P("共１２项。")]), "fail", "全角数字");
  expect("数位混用", "check-doc.js", await build("mix.docx", [P("共 3 千万元。")]), "fail", "数位混用");
  expect("日期简写", "check-doc.js", await build("ymd.docx", [P("2026-10 完成。")]), "fail", "日期简写");
  expect("日期顿号", "check-doc.js", await build("dun.docx", [P("2026、10、12 完成。")]), "fail", "日期顿号");
  expect("表格同上", "check-doc.js", await build("same.docx", [...G.table({
    head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "同上"]],
  })]), "fail", "表格同上");
  expect("表格空格", "check-doc.js", await build("emptycell.docx", [...G.table({
    head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", ""]],
  })]), "fail", "表格空格");
  expect("表格句号", "check-doc.js", await build("period.docx", [...G.table({
    head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "10。"]],
  })]), "fail", "表格句号");
  expect("表头夹英文", "check-doc.js", await build("gdp.docx", [...G.table({
    head: ["指标", "GDP（亿元）"], widths: [3000, 3000], rows: [["地区生产总值", "120"]],
  })]), "hint", "表头夹英文");
  expect("数值列非数值", "check-doc.js", await build("rightprose.docx", [...G.table({
    head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "说明"]],
    align: [docx.AlignmentType.LEFT, docx.AlignmentType.RIGHT],
  })]), "hint", "数值列非数值");
  expect("标题重复编号", "check-doc.js", await build("duphead.docx", [G.h1("一、工作")]), "fail", "标题重复编号");
  expect("中文样式 ID 标题重复", "check-doc.js", await withChineseHeading(await build("cnhead-base.docx", [G.h1("一、工作")])), "fail", "标题重复编号");
  expect("公文层次序数", "check-doc.js", await build("dotted.docx", [P("1.1 这应写成 一、")]), "fail", "公文层次序数");
  expect("中文数字日期", "check-doc.js", await build("cndate.docx", [P("二〇二六年十月。")]), "fail", "中文数字日期");
  expect("月日编虚位", "check-doc.js", await build("padded.docx", [P("2026 年 05 月 08 日。")]), "fail", "月日编虚位");
  expect("成文日期缺日", "check-doc.js", await build("monthonly.docx", [P("2026年10月")]), "fail", "成文日期");
  expect("发文字号两位年份", "check-doc.js", await build("shortyear.docx", [P("×发〔26〕12号")]), "fail", "发文字号");
  expect("书名号里的年份", "check-doc.js", await build("booktitle.docx", [P("《二〇三五年远景目标纲要》提出。")]), "pass", "");
  expect("表格属性", "check-doc.js", await withTableAttr(await build("attrbase.docx", [P("表。"), ...G.table({
    head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "10"], ["乙", "20"], ["合计", "99"]],
  })])), "pass", "");

  // 纵向合并续行不是空格
  {
    const cell = (t, vm) => new docx.TableCell({ verticalMerge: vm, children: [P(t)] });
    const t = new docx.Table({ rows: [
      new docx.TableRow({ tableHeader: true, children: [new docx.TableCell({ children: [P("项目")] }), new docx.TableCell({ children: [P("金额")] })] }),
      new docx.TableRow({ children: [cell("甲", docx.VerticalMergeType.RESTART), cell("10")] }),
      new docx.TableRow({ children: [cell("", docx.VerticalMergeType.CONTINUE), cell("20")] }),
      new docx.TableRow({ children: [new docx.TableCell({ children: [P("合计")] }), new docx.TableCell({ children: [P("30")] })] }),
    ] });
    expect("纵向合并", "check-doc.js", await build("vmerge.docx", [P("表。"), t]), "pass", "");
  }

  // 表后第一段不能被当成表格内容
  expect("表后 1.1", "check-doc.js", await build("aftertable.docx", [
    P("表前。"), ...G.table({ head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "10"], ["乙", "20"], ["合计", "30"]] }),
    P("1.1 这应写成 一、"),
  ]), "fail", "公文层次序数");

  // 版心：多节只查第一节会漏；这里用单节负对照
  expect("版心", "check-doc.js", await build("margin.docx", [P("正文。")], {
    page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1984, left: 1587, right: 1474 } },
    grid: G.GRID,
  }), "fail", "公文版式");
  // 多节：第二节版心错，不能只看第一节
  {
    const good = { page: { size: { width: 11906, height: 16838 }, margin: G.MARGIN }, grid: G.GRID };
    const bad = { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1984, left: 1587, right: 1474 } }, grid: G.GRID };
    const d = new docx.Document({ externalStyles: G.externalStyles(), numbering: G.numbering, sections: [
      { properties: good, children: [P("第一节。")] },
      { properties: bad, children: [P("第二节。")] },
    ] });
    const f = path.join(TMP, "multisect.docx");
    fs.writeFileSync(f, await docx.Packer.toBuffer(d));
    expect("多节版心", "check-doc.js", f, "fail", "公文版式");
  }

  console.log("\n== 数字规则 ==");
  expect("表格属性合计", "check-num.js", await withTableAttr(await build("numattr.docx", [...G.table({
    head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "10"], ["乙", "20"], ["合计", "99"]],
  })])), "fail", "合计不符");
  expect("合计带单位", "check-num.js", await build("totalunit.docx", [...G.table({
    head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "10"], ["乙", "20"], ["合计（万元）", "99"]],
  })]), "fail", "合计不符");

  console.log("\n== 文风规则 ==");
  const writeTxt = (name, text) => { const f = path.join(TMP, name); fs.writeFileSync(f, text); return f; };
  expect("文件名不改", "check-style.js", writeTxt("file.txt", "详见公文.docx，访问中文.com。"), "pass", "");
  expect("书名号不改", "check-style.js", writeTxt("book.md", "《筹画学》是一本旧书。"), "pass", "");
  {
    const { loadTsv } = require("../checks/style-scan");
    let threw = false;
    try { loadTsv("no-such.tsv", 4, { required: true }); } catch { threw = true; }
    ok(threw, "规范层词表缺失必须抛错（不能静默放行）");
  }

  console.log(fails ? `\n规则对照失败 ${fails} 项` : "\n规则对照全部通过");
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error("规则对照跑不起来：", e); process.exit(2); });
