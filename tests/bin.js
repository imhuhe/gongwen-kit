#!/usr/bin/env node
/**
 * bin.js —— 断言统一入口 bin/gongwen.js 在仓库之外的目录也能用，退出码如实透传。
 *
 * agent 的工作目录是用户的项目，不是本仓库，所以入口不能依赖当前目录：
 * 这里在系统临时目录里走完 new → 出稿 → check → fix，再各造一个必错的调用。
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");
const { spawnSync } = require("child_process");

const BIN = path.resolve(__dirname, "..", "bin", "gongwen.js");
const cwd = fs.mkdtempSync(path.join(os.tmpdir(), "gongwen-bin-"));
// 断言失败也要清掉自己建的目录
process.on("exit", () => fs.rmSync(cwd, { recursive: true, force: true }));
const run = (args) => spawnSync(process.execPath, args, { cwd, encoding: "utf8" });
const gw = (...args) => run([BIN, ...args]);

let r = gw("where");
assert.strictEqual(r.status, 0);
assert.strictEqual(fs.realpathSync(r.stdout.trim()), fs.realpathSync(path.dirname(path.dirname(BIN))), "where 应打印工具包目录");

r = gw("new", "--form", "wen", "--out", "a.js");
assert.strictEqual(r.status, 0, "new 失败：" + r.stderr);
r = run(["a.js"]);
assert.strictEqual(r.status, 0, "稿源脚本出稿失败：" + r.stdout + r.stderr);
assert.ok(fs.existsSync(path.join(cwd, "a.docx")), "没写出 a.docx");

r = gw("check", "a.docx", "--strict");
assert.strictEqual(r.status, 0, "check 应通过：" + r.stdout + r.stderr);
for (const w of ["体例机检通过", "数字自洽通过", "文风机检通过"]) {
  assert.ok(r.stdout.includes(w), `check 输出里没有「${w}」`);
}

r = gw("fix", "a.docx");
assert.strictEqual(r.status, 0, "fix 失败：" + r.stderr);
assert.ok(fs.existsSync(path.join(cwd, "a-修订.docx")), "没写出 a-修订.docx");

// 负对照：读不到的文件、未知子命令、缺参数，退出码都必须非零
assert.notStrictEqual(gw("check", "不存在.docx").status, 0, "读不到文件时 check 不应返回 0");
assert.strictEqual(gw("bogus").status, 2, "未知子命令应返回 2");
assert.strictEqual(gw("check").status, 2, "check 缺参数应返回 2");

// 选项写在文件名前面也要认；读不到文件是工具故障（2），且只报一次，不再接着跑后两道
assert.strictEqual(gw("check", "--strict", "a.docx").status, 0, "选项在前时 check 应通过");
r = gw("check", "--strict", "不存在.docx");
assert.strictEqual(r.status, 2, "读不到文件应返回 2");
assert.strictEqual((r.stderr.match(/读不进这份文件/g) || []).length, 1, "工具故障后应立即中止，不再跑后两道");

// --json：通过时 status 为 pass；标题里放一个异形词（筹画），文风规范层必须命中并标 blocking
r = gw("check", "a.docx", "--json");
assert.strictEqual(r.status, 0, "--json 通过时应返回 0");
let rep = JSON.parse(r.stdout);
assert.strictEqual(rep.status, "pass");
assert.strictEqual(rep.toolVersion, require("../package.json").version, "JSON 应带工具版本");
assert.match(rep.generatedAt, /^\d{4}-\d{2}-\d{2}T/, "JSON 应带生成时间");
assert.match(rep.sha256, /^[0-9a-f]{64}$/, "JSON 应带文档 sha256");
// agent 的 shell 常把 stdout 与 stderr 合并，--json 通过时 stderr 必须为空，否则合并后解析不了
assert.strictEqual(r.stderr, "", "check --json 通过时 stderr 应为空：" + r.stderr);
assert.deepStrictEqual(rep.checks.map((c) => c.check), ["体例", "数字自洽", "文风"]);

assert.strictEqual(gw("new", "--form", "wen", "--out", "b.js", "--title", "关于筹画工作的通知").status, 0);
run(["b.js"]); // 出稿时机检会命中并以非零退出，docx 照样写出
assert.ok(fs.existsSync(path.join(cwd, "b.docx")), "没写出 b.docx");
r = gw("check", "b.docx", "--json");
assert.strictEqual(r.status, 1, "--json 时退出码由 status 决定：有拦稿命中就是 1，不看 --strict");
rep = JSON.parse(r.stdout);
assert.strictEqual(rep.status, "hit");
assert.ok(rep.blocking >= 1, "应有拦稿命中");
const style = rep.checks.find((c) => c.check === "文风");
assert.ok(style.hits.some((h) => h.rule === "异形词" && h.blocking), "文风层应报出异形词且标 blocking");
assert.ok(style.hits.some((h) => h.rule === "异形词"), "rule 应是规则名");
assert.strictEqual(gw("check", "b.docx", "--json", "--strict").status, 1, "带 --strict 结果相同");
r = gw("check", "不存在.docx", "--json");
const missing = JSON.parse(r.stdout);
assert.strictEqual(missing.status, "fault", "读不到文件时 JSON 的 status 应为 fault");
assert.strictEqual(missing.sha256, null, "读不到文件时 sha256 应为 null");
assert.strictEqual(r.status, 2);
// 组名带空格的风格层命中（「AI 套话」），rule 必须是完整组名，不能被切成「AI」
assert.strictEqual(gw("new", "--form", "wen", "--out", "c.js", "--title", "值得一提的是工作通知").status, 0);
run(["c.js"]);
rep = JSON.parse(gw("check", "c.docx", "--json").stdout);
const aiHit = rep.checks.find((c) => c.check === "文风").hits.find((h) => h.detail.includes("值得一提的是"));
assert.ok(aiHit, "文风层应命中「值得一提的是」");
assert.strictEqual(aiHit.rule, "AI 套话", "rule 应为完整组名");
assert.strictEqual(aiHit.blocking, false, "风格层不拦稿");
// 出错路径也必须有可解析的 JSON：缺文件名
r = gw("check", "--json");
assert.strictEqual(r.status, 2, "缺文件名应返回 2");
assert.strictEqual(JSON.parse(r.stdout).status, "fault", "缺文件名时 --json 也要给出 fault 的 JSON");

console.log("统一入口测试通过（在仓库之外的目录走完 new、出稿、check、fix、check --json 与各负对照）");
