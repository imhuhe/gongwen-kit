#!/usr/bin/env node
/**
 * fix.js —— 修复工具（tools/fix-doc.js）的测试。
 *
 *   一、规则标本：每条修复规则一对「改前→改后」，外加该放过的（网址、引号、10:30、序号点……）
 *   二、乱稿：tests/messy-docx.js 造的手写稿
 *        · 全部拒绝修订 = 原稿，逐段逐字（修复只能多出可撤销的改动）
 *        · 全部接受修订后，体例与文风的规范层机检没有剩余
 *        · 数字问题只插批注、不改数；超链接原样保留；版心、网格、样式改到位
 *   三、合规件：本仓生成的样张修一遍，修订与批注都必须是 0——不许误改
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const JSZip = require("jszip");
const docx = require("docx");
const { editsFor } = require("../lib/fix/text-rules");
const { acceptAll, rejectAll } = require("../lib/fix/revisions");
const { textOf } = require("../lib/table-read");
const { build } = require("./messy-docx");
const G = require("../lib/gongwen-core").form("wen");
const checkDoc = require("../checks/check-doc");
const { scan } = require("../checks/style-scan");

const ROOT = path.join(__dirname, "..");
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "gongwen-fixtest-"));
let fails = 0;
const ok = (cond, msg) => { console.log(`  ${cond ? "✓" : "✗"} ${msg}`); if (!cond) fails++; };
const apply = (t, es) => { let o = "", i = 0; for (const e of es) { o += t.slice(i, e.start) + e.text; i = e.end; } return o + t.slice(i); };
// 段落文字（实体解码后比：原稿写 &quot;、改后写 " 是同一个字）
const { decode } = require("../lib/fix/runs");
const paraTexts = (xml) => (xml.match(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g) || []).map((p) => decode(textOf(p)));
// 段落签名：文字 + pPr + 每个 run 的 rPr。合规件 fix 前后必须逐段相同；
// 只数修订/批注会漏掉「表题/附件页/版记被直接改字体」这类静默误改。
const paraSig = (xml) => (xml.match(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g) || []).map((p) => {
  const ppr = ((p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [""])[0]).replace(/\s+/g, " ");
  const runs = [...p.matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)].map((r) => ((r[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0]).replace(/\s+/g, " "));
  return JSON.stringify([decode(textOf(p)), ppr, runs]);
});
const ME = "gongwen-kit";
const docXml = async (file) => (await JSZip.loadAsync(fs.readFileSync(file))).file("word/document.xml").async("string");
const fix = (file, out) => execFileSync("node", [path.join(ROOT, "tools/fix-doc.js"), file, "--out", out], { encoding: "utf8" });

(async () => {
  console.log("== 规则标本 ==");
  const CASES = [
    ["月日编虚位", "于2026年05月08日前报送", "于2026年5月8日前报送"],
    ["发文字号括号", "×发【2026】12号", "×发〔2026〕12号"],
    ["发文字号第与虚位", "×发〔2026〕第012号", "×发〔2026〕12号"],
    ["中文数字日期", "二〇二六年十月十三日", "2026年10月13日"],
    ["中文数字日期（年月）", "二〇二六年三月", "2026年3月"],
    ["全角数字", "共１２个项目", "共12个项目"],
    ["半角括号", "各县(市、区)", "各县（市、区）"],
    ["半角标点", "现通知如下:一是,二是.", "现通知如下：一是，二是。"],
    ["右括号后的句点", "（详见附件).", "（详见附件）。"],
    ["异形词", "筹画项目", "筹划项目"],
    ["数字与单位", "容量100kWh", "容量100 kWh"],
    ["纯小数", "厚度.5毫米", "厚度0.5毫米"],
    // 该放过的
    ["网址不改", "见 https://a.b/c,d 页", "见 https://a.b/c,d 页"],
    ["文件名不改", "详见公文.docx", "详见公文.docx"],
    ["域名不改", "访问中文.com", "访问中文.com"],
    ["引号里的原话不改", "「筹画」一词", "「筹画」一词"],
    ["时刻不改", "10:30开会", "10:30开会"],
    ["序号点不改", "附件：1.工作方案", "附件：1.工作方案"],
    ["小数不改", "增长2.5亿元。", "增长2.5亿元。"],
    ["英文缩写不改", "see e.g. this", "see e.g. this"],
    ["无〇的年份不改", "二零二六年", "二零二六年"],
    ["非发文字号的方括号不改", "参见[2026]年报", "参见[2026]年报"],
    ["书名号里的文件名照录", "《二〇三五年远景目标纲要》", "《二〇三五年远景目标纲要》"],
    ["留空待填的金额不补 0", "（¥      .00元整）", "（¥      .00元整）"],
    ["断开的数不补 0", "纯度≥ 99 .8%", "纯度≥ 99 .8%"],
  ];
  for (const [name, before, after] of CASES) {
    const got = apply(before, editsFor(before));
    ok(got === after, `${name}：${before} → ${got}${got === after ? "" : `（应为 ${after}）`}`);
  }
  {
    const h = "一、总体要求";
    ok(apply(h, editsFor(h, { heading: true, numbered: true })) === "总体要求", "标题重复编号：自动编号的标题去掉手写的「一、」");
    ok(apply(h, editsFor(h, { heading: true, numbered: false })) === h, "标题重复编号：没有自动编号的标题不动");
  }

  console.log("\n== 乱稿 ==");
  const messy = path.join(TMP, "messy.docx");
  const fixed = path.join(TMP, "messy-修订.docx");
  await build(messy);
  const report = fix(messy, fixed);
  const orig = await docXml(messy);
  const out = await docXml(fixed);

  const before = paraTexts(orig);
  const rejected = paraTexts(rejectAll(out, ME));
  ok(before.length === rejected.length && before.every((t, i) => t === rejected[i]),
    `全部拒绝修订 = 原稿（${before.length} 段逐字相同）`);
  {
    const diff = before.map((t, i) => [t, rejected[i]]).filter(([a, b]) => a !== b);
    if (diff.length) diff.slice(0, 3).forEach(([a, b]) => console.log(`      原：${a}\n      拒：${b}`));
  }
  ok(!/<w:i\/>/.test(acceptAll(out, ME)), "全部接受后没有斜体；拒绝后斜体回来：" + /<w:i\/>/.test(rejectAll(out, ME)));

  const acc = paraTexts(acceptAll(out, ME));
  const has = (s) => acc.some((t) => t.includes(s));
  for (const s of ["×发改〔2026〕12号", "各县（市、区）发展改革局：", "为做好项目申报，现将有关事项通知如下。",
    "筹划项目，于2026年5月8日前报送", "（详见附件）。", "100 kWh", "2026年10月13日"]) {
    ok(has(s), `接受后：${s}`);
  }
  // 结构：红头（样式红）、文本框、页眉、页面设置修订
  const paraOf = (xml, s) => (xml.match(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g) || []).find((p) => decode(textOf(p)).includes(s)) || "";
  ok(/关于做好2026年度项目申报工作的通知/.test(report) && !/人民政府办公室」改为|内部资料」改为/.test(report), "标题认对了：样式红的红头和居中文本框都不是标题");
  ok(/w:sz w:val="24"/.test(paraOf(out, "人民政府办公室")) && /w:sz w:val="24"/.test(paraOf(out, "内部资料")), "红头（样式红）与文本框里的字体字号不动");
  for (const s of ["各县", "人民政府办公室", "×发改", "内部资料", "地区", "××市发展和改革委员会", "13日"]) {
    ok(/w:firstLine="0"/.test(paraOf(out, s)), `原来不缩进的段落补了首行不缩进：${s}`);
  }
  ok(!/w:firstLine="0"/.test(paraOf(out, "为做好项目申报")), "正文段落照样跟着 Normal 缩进二字");
  ok(/<w:headerReference w:type="even" r:id="rIdHdr"\/>/.test(out), "只有默认页眉的节，双页沿用同一个页眉");
  {
    const sect = out.slice(out.lastIndexOf("<w:body>")).match(/<w:sectPr>[\s\S]*<\/w:sectPr>/)[0];
    const change = sect.match(/<w:sectPrChange[\s\S]*<\/w:sectPrChange>/)[0];
    ok(/w:top="2098"/.test(sect.replace(change, "")) && /w:top="1440"/.test(change) && sect.endsWith("</w:sectPrChange></w:sectPr>"),
      "页面设置修订里的旧 sectPr 原样保留，新版心写在外层");
  }
  ok(/<w:hyperlink r:id="rIdLink">[\s\S]*申报系统[\s\S]*?<\/w:hyperlink>/.test(out), "超链接原样保留");
  ok(has("1200") && has("800") && has("2100"), "表格里的数不改（合计错也不改）");
  ok((out.match(/w:author="审核人"/g) || []).length === 3, "原稿里别人的修订原样保留");
  ok(has("分包专业：装饰工程"), "带旧属性嵌套（pPrChange）的段落正常处理");

  // 接受全部修订后另存，跑规范层机检
  const accFile = path.join(TMP, "accepted.docx");
  {
    const z = await JSZip.loadAsync(fs.readFileSync(fixed));
    z.file("word/document.xml", acceptAll(out, ME));
    fs.writeFileSync(accFile, await z.generateAsync({ type: "nodebuffer" }));
  }
  const left = checkDoc.check(accFile, await checkDoc.readDocx(accFile)).map((h) => h.rule);
  ok(left.length === 0, `接受后体例机检无剩余${left.length ? "：" + left.join("、") : ""}`);
  const st = scan(acc.join("\n").replace(/「[^」\n]*」/g, "　"));
  ok(st.norm.length === 0, `接受后文风规范层无剩余${st.norm.length ? "：" + st.norm.join("；") : ""}`);

  const z = await JSZip.loadAsync(fs.readFileSync(fixed));
  const comments = await z.file("word/comments.xml").async("string");
  ok(/合计不符/.test(comments), "合计错插了批注");
  ok(/计数与列举/.test(comments), "「以下三点」只列两条插了批注");
  ok(!/位置未定位/.test(comments), "批注都定位到了原文段落");
  ok(/w:top="2098"/.test(out) && /w:linePitch="579"/.test(out), "版心 37mm、每页 22 行");
  const styles = await z.file("word/styles.xml").async("string");
  ok(/<w:rPrDefault><w:rPr><w:rFonts w:ascii="FangSong"/.test(styles), "默认字体仿宋");
  ok(/w:styleId="1">[\s\S]*?SimHei[\s\S]*?<\/w:style>/.test(styles), "中文样式 ID 的一级标题改成黑体");
  ok(/w:firstLineChars="200"/.test(styles), "正文首行缩进二字");
  ok(!/w:line="360"/.test(paraOf(out, "为做好项目申报")), "正文手设的 1.5 倍行距交还网格");
  ok(/w:line="360"/.test(paraOf(out, "××市发展和改革委员会")), "落款署名的手设行距不被正文规则改掉");
  ok(z.file("word/footerGwOdd.xml") && /PAGE/.test(await z.file("word/footerGwOdd.xml").async("string")), "加了公文页码");
  ok(/原文件未改动/.test(report), "报告写明原文件未改动");
  {
    let refused = false;
    try { fix(messy, messy); } catch (e) { refused = e.status === 2; }
    ok(refused && (await docXml(messy)) === orig, "--out 指向原文件时拒绝执行，原稿不被覆盖");
  }
  ok(fs.readFileSync(messy).equals(fs.readFileSync(messy)) && (await docXml(messy)) === orig, "原文件没被改");

  console.log("\n== 角色保护 ==");
  {
    const mk = async (name, children) => {
      const d = new docx.Document({ externalStyles: G.externalStyles(), numbering: G.numbering, sections: [{ ...G.sectionProps(), children }] });
      const f = path.join(TMP, name);
      fs.writeFileSync(f, await docx.Packer.toBuffer(d));
      return f;
    };
    const red = path.join(TMP, "red.docx");
    await mk("red.docx", [
      new docx.Paragraph({ children: [new docx.TextRun({ text: "××市筹画局", color: "FF0000" })] }),
      G.p("正文。"),
    ]);
    fix(red, path.join(TMP, "red-out.docx"));
    const redText = paraTexts(acceptAll(await docXml(path.join(TMP, "red-out.docx")), ME)).join("\n");
    ok(redText.includes("××市筹画局"), "红头机关名不按异形词改");

    const tblFile = path.join(TMP, "tbl.docx");
    await mk("tbl.docx", [
      G.p("表。"),
      ...G.table({ head: ["项目", "数值"], widths: [3000, 3000], rows: [["甲", "１００"], ["乙", "100kWh"]] }),
    ]);
    fix(tblFile, path.join(TMP, "tbl-out.docx"));
    const tblXml = await docXml(path.join(TMP, "tbl-out.docx"));
    ok(!/<w:(?:ins|del)\b/.test(tblXml), "表格里的数字不自动改（只插批注）");
    ok(/１００/.test(tblXml) && /100kWh/.test(tblXml), "表格原值保留");

    // 表格开标签带属性（Word 的 w:rsidR）时，整张表不能被当成正文
    const attrFile = path.join(TMP, "tbl-attr.docx");
    {
      const z = await JSZip.loadAsync(fs.readFileSync(tblFile));
      const x = await z.file("word/document.xml").async("string");
      z.file("word/document.xml", x.replace("<w:tbl>", '<w:tbl w:rsidR="00A1B2C3">'));
      fs.writeFileSync(attrFile, await z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    }
    fix(attrFile, path.join(TMP, "tbl-attr-out.docx"));
    const attrXml = await docXml(path.join(TMP, "tbl-attr-out.docx"));
    ok(!/<w:(?:ins|del)\b/.test(attrXml) && /１００/.test(attrXml) && /100kWh/.test(attrXml),
      "表格带属性时仍按表格保护");
    ok(!/项目金额/.test(attrXml), "表格头不会被当成标题");

    // 表题没有 keepNext（手写/旧模板常见），靠“紧邻表格”识别
    const capFile = path.join(TMP, "caption.docx");
    await mk("caption.docx", [
      new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [
        new docx.TextRun({ text: "表 1 说明", font: { eastAsia: "SimHei", ascii: "SimHei" }, size: 28 }),
      ] }),
      ...G.table({ head: ["项目", "金额"], widths: [3000, 3000], rows: [["甲", "10"], ["合计", "10"]] }),
    ]);
    fix(capFile, path.join(TMP, "caption-out.docx"));
    const capOut = await docXml(path.join(TMP, "caption-out.docx"));
    // 表题允许补「首行不缩进」以抵消 Normal 的两字缩进，但字体字号必须保留。
    ok(/表 1 说明/.test(capOut) && /w:rFonts[^>]*SimHei/.test(capOut) && /<w:sz w:val="28"\/>/.test(capOut),
      "无 keepNext 的表题字体字号保留");

    // 上行文多个签发人：续行只有姓名、右对齐，不能被当成正文
    const signersFile = path.join(TMP, "signers.docx");
    await mk("signers.docx", [
      ...G.letterhead({ agency: "××局", docNo: { dept: "×发", year: 2026, seq: 1 }, signers: ["张××", "李××", "王××"] }),
      ...G.title(["关于测试事项的请示"]),
      G.p("正文。"),
      ...G.luokuan({ unit: "××局", date: "2026 年 10 月 23 日" }),
    ]);
    const signersBefore = await docXml(signersFile);
    fix(signersFile, path.join(TMP, "signers-out.docx"));
    ok(JSON.stringify(paraSig(signersBefore)) === JSON.stringify(paraSig(await docXml(path.join(TMP, "signers-out.docx")))),
      "多签发人续行不被当成正文");

    // 正文引用发文字号不是版头字号，不能整段跳过修订
    const citeFile = path.join(TMP, "cite.docx");
    await mk("cite.docx", [new docx.Paragraph({ children: [
      new docx.TextRun({ text: "根据国办发〔2026〕5号要求，统一筹画规范。", font: { eastAsia: "KaiTi", ascii: "KaiTi" }, size: 24 }),
    ] })]);
    const citeReport = fix(citeFile, path.join(TMP, "cite-out.docx"));
    const citeXml = await docXml(path.join(TMP, "cite-out.docx"));
    const citeText = paraTexts(acceptAll(citeXml, ME)).join("\n");
    ok(citeText.includes("筹划") && /异形词 1 处/.test(citeReport) && !/w:sz w:val="24"/.test(citeXml),
      "正文引用发文字号后仍改异形词、手设字号交还样式");

    // 斜体计数按 run，不按段
    const italFile = path.join(TMP, "ital2.docx");
    await mk("ital2.docx", [new docx.Paragraph({ children: [
      new docx.TextRun({ text: "斜体一", italics: true }),
      new docx.TextRun({ text: "与" }),
      new docx.TextRun({ text: "斜体二", italics: true }),
    ] })]);
    const italReport = fix(italFile, path.join(TMP, "ital2-out.docx"));
    ok(/斜体改正体 2 处/.test(italReport), "斜体改正体计数按 run");

    const dry = path.join(TMP, "dry.docx");
    execFileSync("node", [path.join(ROOT, "tools/fix-doc.js"), tblFile, "--dry-run", "--out", dry], { encoding: "utf8", stdio: "pipe" });
    ok(!fs.existsSync(dry), "--dry-run 不写文件");

    let noOut = false;
    try { execFileSync("node", [path.join(ROOT, "tools/fix-doc.js"), tblFile, "--out"], { encoding: "utf8", stdio: "pipe" }); } catch (e) { noOut = e.status === 2; }
    ok(noOut, "--out 缺值返回 2");
  }

  console.log("\n== 合规件不误改 ==");
  for (const t of ["notice", "gongwen-qingshi", "complete", "xinhan", "jiyao", "qiye", "gongwen-guanli"]) {
    const label = t === "gongwen-guanli" ? "公文管理办法" : t;
    const src = path.join(TMP, `${label}.docx`);
    execFileSync("node", [path.join(ROOT, "templates", `${t}.js`), src], { encoding: "utf8", stdio: "pipe" });
    const dst = path.join(TMP, `${label}-修订.docx`);
    const before = await docXml(src);
    const r = fix(src, dst);
    const x = await docXml(dst);
    const nRev = (x.match(/<w:(?:ins|del|rPrChange)\b/g) || []).length;
    const nCom = (x.match(/<w:commentReference\b/g) || []).length;
    const sigA = paraSig(before), sigB = paraSig(x);
    const diff = sigA.findIndex((v, i) => v !== sigB[i]);
    ok(nRev === 0 && nCom === 0 && sigA.length === sigB.length && diff < 0,
      `${label}：修订 ${nRev} 处、批注 ${nCom} 条、段落格式 ${diff < 0 ? "不变" : "第 " + (diff + 1) + " 段变了"}${nRev + nCom || diff >= 0 ? "\n" + r : ""}`);
    if (t === "gongwen-guanli") {
      const z = await JSZip.loadAsync(fs.readFileSync(dst));
      const styles = await z.file("word/styles.xml").async("string");
      const h1 = (styles.match(/<w:style\b[^>]*w:styleId="Heading1"[\s\S]*?<\/w:style>/) || [""])[0];
      const titleFont = process.env.GONGWEN_FONT_TITLE || "方正小标宋简体";
      ok(/<w:jc w:val="center"\/>/.test(h1) && h1.includes(titleFont) && /章小标宋居中/.test(r),
        "公文管理办法：章标题样式居中、字体为标题字体，报告口径正确");
    }
  }

  fs.rmSync(TMP, { recursive: true, force: true });
  console.log(fails ? `\n修复测试 ${fails} 项不过` : "\n修复测试全部通过");
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
