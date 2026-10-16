#!/usr/bin/env node
/**
 * pack-offline.js —— 打一个内网和离线机器可用的安装包（含生产依赖，不含样张图片和测试）。
 *
 * 用法：在能上网的机器上 node tools/pack-offline.js [--out 目录]
 *       产出 gongwen-kit-v<版本>-offline.tgz，默认写到当前目录。
 *
 * 目标机器只需要 Node 18 以上：解压后直接用，不跑 npm、不联网。
 * 依赖都是纯 JavaScript（docx、jszip、sax），包在 Windows、macOS、Linux 之间通用。
 *
 * 包里的文件取自 git 已跟踪的文件，所以打包前先提交；排除 assets/（宣传图）、
 * docs/样张/（PDF 样张）、.github/ 与 tests/。
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const args = process.argv.slice(2);
const outDir = path.resolve(args.includes("--out") ? args[args.indexOf("--out") + 1] : process.cwd());
const version = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8")).version;
const name = `gongwen-kit-v${version}`;
const exclude = [/^assets\//, /^docs\/样张\//, /^\.github\//, /^tests\//];

const stage = fs.mkdtempSync(path.join(os.tmpdir(), "gongwen-pack-"));
try {
  const dest = path.join(stage, name);
  const files = execFileSync("git", ["ls-files", "-z"], { cwd: ROOT, encoding: "utf8" })
    .split("\0").filter((f) => f && !exclude.some((re) => re.test(f)));
  for (const f of files) {
    fs.mkdirSync(path.dirname(path.join(dest, f)), { recursive: true });
    fs.copyFileSync(path.join(ROOT, f), path.join(dest, f));
  }
  // package.json 里的 test 脚本指向已排除的 tests/，包里没有它
  const pkgFile = path.join(dest, "package.json");
  const pkg = JSON.parse(fs.readFileSync(pkgFile, "utf8"));
  delete pkg.scripts;
  fs.writeFileSync(pkgFile, JSON.stringify(pkg, null, 2) + "\n");

  const npm = process.platform === "win32" ? "npm.cmd" : "npm";
  execFileSync(npm, ["ci", "--omit=dev", "--ignore-scripts"], { cwd: dest, stdio: "inherit", shell: process.platform === "win32" });

  fs.mkdirSync(outDir, { recursive: true });
  const out = path.join(outDir, `${name}-offline.tgz`);
  execFileSync("tar", ["-czf", out, "-C", stage, name], { stdio: "inherit" });
  console.log(`→ ${out} (${(fs.statSync(out).size / 1024 / 1024).toFixed(1)} MB)`);
  console.log("目标机器：解压后 node <解压目录>/bin/gongwen.js where 确认能跑，再按 README 链进 skill 目录。");
} finally {
  fs.rmSync(stage, { recursive: true, force: true });
}
