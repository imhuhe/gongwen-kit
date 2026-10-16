#!/usr/bin/env node
/**
 * pack.js —— 离线打包脚本 tools/pack-offline.js 的测试。
 *
 * 不联网：造一个只有几个文件的临时 git 仓库，用 --root 指过去、--skip-install 跳过 npm ci，
 * 只验文件选择、排除项、脏工作区拒绝和 tar 产物。npm ci 那一步由人在能上网的机器上打包时走。
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");
const { spawnSync, execFileSync } = require("child_process");

const PACK = path.resolve(__dirname, "..", "tools", "pack-offline.js");
const base = fs.mkdtempSync(path.join(os.tmpdir(), "gongwen-pack-test-"));
process.on("exit", () => fs.rmSync(base, { recursive: true, force: true }));

const repo = path.join(base, "repo");
const outDir = path.join(base, "out");
const put = (rel, text = "x\n") => {
  fs.mkdirSync(path.dirname(path.join(repo, rel)), { recursive: true });
  fs.writeFileSync(path.join(repo, rel), text);
};
const git = (...a) => execFileSync("git", ["-c", "user.name=t", "-c", "user.email=t@t", ...a], { cwd: repo, stdio: "pipe" });
const pack = (...extra) =>
  spawnSync(process.execPath, [PACK, "--root", repo, "--skip-install", "--out", outDir, ...extra], { encoding: "utf8" });

fs.mkdirSync(repo);
git("init", "-q");
put("package.json", JSON.stringify({ name: "fake", version: "9.9.9", scripts: { test: "x" }, dependencies: {} }));
for (const f of ["SKILL.md", "README.md", "AGENTS.md", "bin/g.js", "docs/说明.md", "docs/样张/a.pdf", "assets/a.png", "tests/t.js"]) put(f);
git("add", "-A");
git("commit", "-q", "-m", "init");

// 1. 干净工作区：打包成功；排除项不在包里，保留项在，scripts 被去掉，附有说明文件
let r = pack();
assert.strictEqual(r.status, 0, "干净工作区应能打包：" + r.stdout + r.stderr);
const tgz = "gongwen-kit-v9.9.9-offline.tgz";
assert.ok(fs.existsSync(path.join(outDir, tgz)), "没写出 " + tgz);
// tar 只用相对路径：Windows 上 Git Bash 带的 GNU tar 会把盘符当成「主机:路径」
fs.mkdirSync(path.join(outDir, "x"));
execFileSync("tar", ["-xzf", tgz, "-C", "x"], { cwd: outDir, stdio: "pipe" });
const pkgDir = path.join(outDir, "x", "gongwen-kit-v9.9.9");
const has = (rel) => fs.existsSync(path.join(pkgDir, rel));
for (const f of ["SKILL.md", "README.md", "bin/g.js", "docs/说明.md", "package.json", "离线包说明.md"]) assert.ok(has(f), `包里应有 ${f}`);
for (const f of ["AGENTS.md", "assets", "tests", "docs/样张"]) assert.ok(!has(f), `包里不应有 ${f}`);
assert.ok(!("scripts" in JSON.parse(fs.readFileSync(path.join(pkgDir, "package.json"), "utf8"))), "package.json 的 scripts 应被去掉");

// 2. 负对照：未跟踪的新文件必须让打包拒绝，且点名这个文件，而不是静默漏掉
put("lib/new.js");
r = pack();
assert.notStrictEqual(r.status, 0, "有未跟踪的新文件时应拒绝打包");
assert.ok(r.stderr.includes("lib/new.js"), "应点名未跟踪的文件：" + r.stderr);

// 3. 已跟踪的文件有改动也拒绝；被排除目录里的改动不算
git("add", "lib/new.js");
git("commit", "-q", "-m", "add");
put("README.md", "changed\n");
r = pack();
assert.notStrictEqual(r.status, 0, "已跟踪文件有未提交改动时应拒绝");
assert.ok(r.stderr.includes("README.md"));
git("checkout", "--", "README.md");
put("tests/t.js", "changed\n");
assert.strictEqual(pack().status, 0, "只有被排除目录里的改动时不应拒绝");
git("checkout", "--", "tests/t.js");

// 4. --allow-dirty 放行
put("lib/another.js");
assert.strictEqual(pack().status === 0, false);
assert.strictEqual(pack("--allow-dirty").status, 0, "--allow-dirty 应放行");

console.log("离线打包测试通过（文件选择、排除项、脏工作区拒绝、--allow-dirty）");
