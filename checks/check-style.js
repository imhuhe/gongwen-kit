#!/usr/bin/env node
/**
 * check-style.js —— 文风机检（出稿链第三道），由 writeDoc 自动调用。
 *
 * 用法：node check-style.js <文件.docx|.txt> [--strict] [--json]
 *       --json 时只在 stdout 输出一个 JSON：{check, file, hits:[{rule, detail, blocking}]}，
 *       规范层 blocking 为 true，风格层为 false，退出码不变。
 *       .docx 抽正文；.txt/.md 直接读（供测试）。
 *       【规范】命中（异形词、中文里的半角标点）--strict 时退出 1；
 *       【风格】命中（AI 套话、商业黑话、形状）只打印，不影响退出码。
 *       文件读不进或扫描器跑崩，退出 2（工具故障，不是稿子的问题）。
 *
 * 扫描器与词表在同目录 style-scan.js 与 ../data/——与私有环境的扫描器
 * （Python）同一套判据；词表出的坑（命中自家语料就降级）见 data/ 表头。
 * 引文豁免：「」与“”里的原话先抠掉再扫。
 */
"use strict";
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { scan } = require("./style-scan");
const { textOf } = require("../lib/table-read");

async function bodyOf(file) {
  if (/\.(txt|md)$/i.test(file)) return fs.readFileSync(file, "utf8");
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file("word/document.xml").async("string");
  return xml.split("</w:p>").map(textOf).filter(Boolean).join("\n");
}

(async () => {
  const file = process.argv[2];
  if (!file) {
    console.error("用法：node check-style.js <文件.docx|.txt> [--strict] [--json]");
    process.exit(2);
  }
  const strict = process.argv.includes("--strict");
  const name = path.basename(file);
  let raw;
  try {
    raw = await bodyOf(file);
  } catch (e) {
    console.error("文风机检读不进这份文件（工具故障，不是稿子的问题）：", e.message);
    process.exit(2);
  }
  const body = raw.replace(/「[^」\n]*」/g, "　").replace(/“[^”\n]*”/g, "　");
  const { norm, style } = scan(body);
  if (process.argv.includes("--json")) {
    const row = (blocking) => (l) => ({ rule: l.split(/[：（\s]/)[0], detail: l, blocking });
    console.log(JSON.stringify({
      check: "文风", file: name,
      hits: [...norm.map(row(true)), ...style.map(row(false))],
    }));
    if (norm.length && strict) process.exit(1);
    return;
  }
  if (!norm.length && !style.length) {
    console.log(`文风机检通过　${name}`);
    return;
  }
  console.log(`文风机检${norm.length ? "不过" : "有提示"}　${name}`);
  if (norm.length) {
    console.log("  【规范】命中即错，改掉：");
    norm.forEach((l) => console.log("    " + l));
  }
  if (style.length) {
    console.log("  【风格】看语境，该留就留：");
    style.forEach((l) => console.log("    " + l));
  }
  if (norm.length && strict) process.exit(1);
})();
