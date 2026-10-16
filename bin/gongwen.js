#!/usr/bin/env node
/**
 * gongwen —— 统一入口，从任何工作目录都能调用，不依赖当前目录和相对路径。
 *
 * 用法：
 *   gongwen new  --form wen|zhang --out 稿子.js [--title "标题"]   起稿源脚本
 *   gongwen fix  稿子.docx [--out 输出.docx] [--no-layout]         修复现成 docx
 *   gongwen check 稿子.docx [--strict]                             依次跑体例、数字、文风三道机检
 *   gongwen pdf  导出件.pdf                                        核字体与版面（需要 pymupdf）
 *   gongwen where                                                  打印本工具包所在目录
 *
 * 只做分派：每个子命令转给 tools/ 或 checks/ 里的原脚本，参数、输出、退出码原样透传，
 * 所以直接调原脚本和走本入口结果相同。
 */
"use strict";
const path = require("path");
const { spawnSync } = require("child_process");

const KIT = path.resolve(__dirname, "..");
const [cmd, ...rest] = process.argv.slice(2);

const run = (bin, script, args) =>
  spawnSync(bin, [path.join(KIT, script), ...args], { stdio: "inherit" }).status ?? 1;

const usage = () => {
  console.error("用法：gongwen new|fix|check|pdf|where …（详见本文件头部注释）");
  process.exit(2);
};

switch (cmd) {
  case "new":
    process.exit(run(process.execPath, "tools/new-doc.js", rest));
  case "fix":
    process.exit(run(process.execPath, "tools/fix-doc.js", rest));
  case "check": {
    if (!rest.length) usage();
    // 三道都跑完再退出，让用户一次看到全部问题；任何一道非零就返回非零
    let worst = 0;
    for (const c of ["check-doc", "check-num", "check-style"]) {
      worst = Math.max(worst, run(process.execPath, `checks/${c}.js`, rest));
    }
    process.exit(worst);
  }
  case "pdf":
    process.exit(run(process.platform === "win32" ? "python" : "python3", "tools/check-pdf.py", rest));
  case "where":
    console.log(KIT);
    process.exit(0);
  default:
    usage();
}
