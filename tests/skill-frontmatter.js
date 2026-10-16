#!/usr/bin/env node
/**
 * skill-frontmatter.js —— 校验 SKILL.md 的 frontmatter 能被各家 agent 加载。
 *
 * 依据：
 *   DeepSeek Harness（dsh-skill-filesystem）：必填 name、description；
 *     disable-model-invocation、user-invocable 只认布尔写法，写错整个 skill 会被丢弃。
 *   Codex：name、description 必填；agents/openai.yaml 可选。
 *   Agent Skills 规范：name 小写字母数字加连字符，最长 64；description 最长 1024。
 *
 * 不引入 YAML 库：frontmatter 只用顶层标量，逐行解析够用；遇到缩进或列表就报错提醒。
 */
"use strict";
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const text = fs.readFileSync(path.join(root, "SKILL.md"), "utf8").replace(/\r\n/g, "\n");
const errors = [];
const fail = (m) => errors.push(m);

const m = text.match(/^---\n([\s\S]*?)\n---\n/);
if (!m) {
  fail("SKILL.md 没有以 --- 包起来的 frontmatter");
} else {
  const fm = {};
  for (const line of m[1].split("\n")) {
    if (!line.trim()) continue;
    if (/^\s/.test(line) || line.startsWith("- ")) {
      fail(`frontmatter 出现缩进或列表行，本校验器不解析：${line}`);
      continue;
    }
    const i = line.indexOf(":");
    if (i < 0) { fail(`frontmatter 行缺少冒号：${line}`); continue; }
    fm[line.slice(0, i).trim()] = line.slice(i + 1).trim();
  }
  if (!fm.name) fail("缺少 name");
  else {
    if (!/^[a-z0-9]+(-[a-z0-9]+)*$/.test(fm.name)) fail(`name 不合规范：${fm.name}`);
    if (fm.name.length > 64) fail("name 超过 64 字符");
  }
  if (!fm.description) fail("缺少 description");
  else if ([...fm.description].length > 1024) fail(`description 超过 1024 字符：${[...fm.description].length}`);
  const bool = /^(true|false|yes|no|on|off|1|0)$/i;
  for (const k of ["disable-model-invocation", "user-invocable"]) {
    if (k in fm && !bool.test(fm[k])) fail(`${k} 不是布尔写法：${fm[k]}`);
  }
}

const yamlFile = path.join(root, "agents", "openai.yaml");
if (fs.existsSync(yamlFile)) {
  const y = fs.readFileSync(yamlFile, "utf8");
  if (/\t/.test(y)) fail("agents/openai.yaml 含制表符，YAML 缩进必须用空格");
  const im = y.match(/allow_implicit_invocation:\s*(\S+)/);
  if (im && !/^(true|false)$/.test(im[1])) fail(`allow_implicit_invocation 必须是 true 或 false：${im[1]}`);
}

if (errors.length) {
  console.error("skill frontmatter 校验失败：");
  for (const e of errors) console.error("  - " + e);
  process.exit(1);
}
console.log("skill frontmatter 校验通过");
