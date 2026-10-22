#!/usr/bin/env node
/**
 * gongwen —— 统一入口，从任何工作目录都能调用，不依赖当前目录和相对路径。
 *
 * 用法：
 *   gongwen new  --form wen|zhang --out 稿子.js [--title "标题"]   起稿源脚本
 *   gongwen fix  稿子.docx [--out 输出.docx] [--no-layout]         修复现成 docx
 *   gongwen check 稿子.docx [--strict] [--json]                    依次跑体例、数字、文风三道机检
 *                                                                  退出码：0 全过，1 命中，2 工具故障
 *                                                                  --json：stdout 只出一份汇总 JSON（出错也是），退出码由 status 决定：
 *                                                                  pass 0、hit 1、fault 2，与 --strict 无关
 *   gongwen pdf  导出件.pdf                                        核字体与版面（需要 pymupdf）
 *   gongwen where                                                  打印本工具包所在目录
 *
 * 只做分派：new、fix、pdf 转给 tools/ 里的原脚本，参数、输出、退出码原样透传；
 * check 直接调 lib/run-checks.js 的 runChecks，与出稿时自动跑的机检链是同一份实现。
 */
"use strict";
const path = require("path");
const { spawnSync } = require("child_process");
const pkg = require("../package.json");

const KIT = path.resolve(__dirname, "..");
const [cmd, ...rest] = process.argv.slice(2);

// 没能启动或被信号杀死时打印原因并返回 2，与核心库的约定一致：2 是工具故障，1 是稿子命中
const run = (bin, script, args) => {
  const r = spawnSync(bin, [path.join(KIT, script), ...args], { stdio: "inherit" });
  if (r.error || r.status === null) {
    console.error(`没能跑起来 ${bin} ${script}（工具故障，不是稿子的问题）：`,
      r.error ? r.error.message : `被信号 ${r.signal} 终止`);
    return 2;
  }
  return r.status;
};

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
    // 文件名和选项分开：选项写在文件名前面也行
    const file = rest.find((a) => !a.startsWith("--"));
    const json = rest.includes("--json");
    let r;
    try {
      if (!file) throw new Error("缺少文件名。用法：gongwen check 稿子.docx [--strict] [--json]");
      r = require("../lib/run-checks").runChecks(file, rest.includes("--strict"), { json });
    } catch (e) {
      if (!json) {
        if (!file) usage();
        throw e;
      }
      // --json 时 stdout 必须有一份可解析的 JSON，出错也不例外
      r = { fault: true, report: { file: file ? path.basename(file) : null, status: "fault", blocking: 0, message: e.message, checks: [],
        toolVersion: pkg.version, generatedAt: new Date().toISOString(), sha256: null } };
    }
    if (json) {
      console.log(JSON.stringify(r.report, null, 2));
      process.exitCode = { pass: 0, hit: 1, fault: 2 }[r.report.status];
    } else {
      process.exitCode = r.fault ? 2 : r.ok ? 0 : 1;
    }
    break;
  }
  case "pdf":
    process.exit(run(process.platform === "win32" ? "python" : "python3", "tools/check-pdf.py", rest));
  case "where":
    console.log(KIT);
    process.exit(0);
  default:
    usage();
}
