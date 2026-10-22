"use strict";
/**
 * run-checks.js —— 依次跑三道机检的汇总器，出稿时自动调用，bin/gongwen.js 的 check 也调它。
 *
 * 单独成模块、不依赖 docx：只用 child_process 与 fs。入口脚本只为跑机检，不该为此加载整个
 * 排版库——加载 docx 会让 Node 往 stderr 写 ExperimentalWarning，--json 时 agent 把两路输出
 * 合并后就解析不了。lib/gongwen-core.js 仍转出 runChecks，老调用方不用改。
 */
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const pkg = require("../package.json");

const ROOT = path.dirname(__dirname);   // 仓根：本文件在 lib/

// 合规凭据要能回答「哪个版本、什么时候、检的哪一份文件」。sha256 是文件内容的
// 指纹，file 只是 basename 便于人读；路径与内容以 sha256 为准。
function meta(file) {
  let sha256 = null;
  try { sha256 = file ? crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex") : null; } catch { /* 读不到时留 null */ }
  return { toolVersion: pkg.version, generatedAt: new Date().toISOString(), sha256 };
}

/** 依次跑体例（check-doc）、数字自洽（check-num）、文风（check-style）三道机检。
 *  脚本退出码 2 或没能起来按工具故障处理：打印原因、立即中止，返回 fault: true；
 *  其余返回 ok 表示三道是否都没命中（命中只在 strict 下以非零退出码体现）。
 *  json 为真时不转出各脚本的文字输出，改为在返回值的 report 里给汇总：
 *  {file, status: "pass"|"hit"|"fault", blocking: 拦稿命中数, checks: [各脚本的 JSON]}。
 *  status 不看 strict：有拦稿命中就是 hit。返回值里的 ok 仍随 strict（命中只在 strict 下算不 ok）；
 *  bin/gongwen.js 的 --json 模式不用 ok，退出码直接由 status 决定。 */
function runChecks(file, strict, { json = false } = {}) {
  const { spawnSync } = require("child_process");
  const jobs = ["check-doc.js", "check-num.js", "check-style.js"];
  let ok = true;
  const checks = [];
  const fault = (message) => ({
    fault: true, ok: false,
    report: json ? { file: path.basename(file), status: "fault", blocking: 0, message, checks, ...meta(file) } : undefined,
  });
  for (const name of jobs) {
    const checker = path.join(ROOT, "checks", name);
    if (!fs.existsSync(checker)) throw new Error("机检脚本不在：" + checker);
    const r = spawnSync(process.execPath,
      [checker, file, ...(strict ? ["--strict"] : []), ...(json ? ["--json"] : [])], { encoding: "utf8" });
    if (!json) process.stdout.write(r.stdout || "");
    process.stderr.write(r.stderr || "");
    if (r.error || r.status === 2 || r.status === null) {
      const why = r.error ? r.error.message : `退出码 ${r.status}`;
      console.error(`机检 ${name} 没能跑起来（工具故障，不是稿子的问题）：`, why);
      return fault(`${name} 没能跑起来：${why}`);
    }
    if (r.status !== 0) ok = false;
    if (json) {
      try {
        checks.push(JSON.parse(r.stdout));
      } catch (e) {
        console.error(`机检 ${name} 的 JSON 输出读不懂（工具故障）：`, e.message);
        return fault(`${name} 的 JSON 输出读不懂：${e.message}`);
      }
    }
  }
  if (!json) return { fault: false, ok };
  const blocking = checks.reduce((n, c) => n + c.hits.filter((h) => h.blocking).length, 0);
  return { fault: false, ok, report: { file: path.basename(file), status: blocking ? "hit" : "pass", blocking, checks, ...meta(file) } };
}

module.exports = { runChecks };
