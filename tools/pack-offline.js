#!/usr/bin/env node
/**
 * pack-offline.js —— 打一个内网和离线机器可用的安装包（含生产依赖，不含样张图片和测试）。
 *
 * 用法：在能上网的机器上 node tools/pack-offline.js [--out 目录]
 *       产出 gongwen-kit-v<版本>-offline.tgz，默认写到当前目录。
 *       --allow-dirty   有未提交的改动也打（默认拒绝，见下）
 *       --skip-install  不跑 npm ci，只打文件（测试用，产出的包没有依赖，不能用）
 *       --root 目录     打这个目录而不是本仓库（测试用）
 *
 * 目标机器只需要 Node 18 以上：解压后直接用，不跑 npm、不联网。
 * 依赖都是纯 JavaScript（docx、jszip、sax），包在 Windows、macOS、Linux 之间通用。
 *
 * 包里的文件取自 git 已跟踪的文件，所以打包前先提交：有未提交的改动或未跟踪的新文件时拒绝，
 * 否则新文件会被静默漏掉，到内网机器上才报找不到模块，那时已没法联网补救。
 * 排除 assets/（宣传图）、docs/样张/（PDF 样张）、.github/、tests/ 与 AGENTS.md
 * （给改本仓库的人看的），并附一份《离线包说明.md》写明这些缺什么。
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const args = process.argv.slice(2);
const die = (msg) => { console.error(msg); process.exit(2); };
const opt = (k) => {
  if (!args.includes(k)) return undefined;
  const v = args[args.indexOf(k) + 1];
  if (!v || v.startsWith("--")) die(`${k} 后面要跟一个目录`);
  return v;
};
const ROOT = path.resolve(opt("--root") || path.join(__dirname, ".."));
const outDir = path.resolve(opt("--out") || process.cwd());
const allowDirty = args.includes("--allow-dirty");
const skipInstall = args.includes("--skip-install");
const version = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8")).version;
const name = `gongwen-kit-v${version}`;
const exclude = [/^assets\//, /^docs\/样张\//, /^\.github\//, /^tests\//, /^AGENTS\.md$/];
const kept = (f) => f && !exclude.some((re) => re.test(f));
const git = (...a) => execFileSync("git", a, { cwd: ROOT, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });

try {
  git("rev-parse", "--show-toplevel");
} catch {
  die("这里不是 git 仓库：打包要在完整仓库里做（离线包本身不能再打包）。");
}

if (!allowDirty) {
  // -z：路径不加引号转义；每项是「两位状态码、空格、路径」。状态含 R 或 C（重命名、复制）时，
  // 紧跟的下一项是旧路径，只有路径、没有状态码，要整体取下来，不能再切前三个字符
  const entries = git("status", "--porcelain", "-z", "--untracked-files=all").split("\0");
  const paths = [];
  for (let i = 0; i < entries.length; i++) {
    if (!entries[i]) continue;
    paths.push(entries[i].slice(3));
    if (/[RC]/.test(entries[i].slice(0, 2))) paths.push(entries[++i]);
  }
  const dirty = paths.filter(kept);
  if (dirty.length) {
    console.error("有未提交的改动或未跟踪的文件，打出来的包会缺它们。先提交，或加 --allow-dirty：");
    for (const f of dirty) console.error("  " + f);
    process.exit(1);
  }
}

const stage = fs.mkdtempSync(path.join(os.tmpdir(), "gongwen-pack-"));
try {
  const dest = path.join(stage, name);
  for (const f of git("ls-files", "-z").split("\0").filter(kept)) {
    fs.mkdirSync(path.dirname(path.join(dest, f)), { recursive: true });
    fs.copyFileSync(path.join(ROOT, f), path.join(dest, f));
  }
  // package.json 里的 test 脚本指向已排除的 tests/，包里没有它
  const pkgFile = path.join(dest, "package.json");
  const pkg = JSON.parse(fs.readFileSync(pkgFile, "utf8"));
  delete pkg.scripts;
  fs.writeFileSync(pkgFile, JSON.stringify(pkg, null, 2) + "\n");

  fs.writeFileSync(path.join(dest, "离线包说明.md"), `# 离线包说明

这是 gongwen-kit v${version} 的离线安装包，已带生产依赖，目标机器只需要 Node 18 以上，不跑 npm、不联网。

\`\`\`bash
node bin/gongwen.js where     # 能打印目录就说明可用
\`\`\`

用法见 README.md「当作 agent skill 用」与 SKILL.md。

**这个包比完整仓库少的东西**：宣传图（assets/）、PDF 样张（docs/样张/）、测试（tests/）、
给改仓库的人看的 AGENTS.md，以及 package.json 里的 scripts。所以 README 里指向图片和样张的链接
在包内打不开，\`npm test\` 也不可用。需要这些请取完整仓库：https://github.com/imhuhe/gongwen-kit
`);

  if (!skipInstall) {
    const npm = process.platform === "win32" ? "npm.cmd" : "npm";
    execFileSync(npm, ["ci", "--omit=dev", "--ignore-scripts"], { cwd: dest, stdio: "inherit", shell: process.platform === "win32" });
  }

  // tar 只用相对路径：Windows 上 Git Bash 带的 GNU tar 会把 `C:\\…` 的盘符当成「主机:路径」
  const tgz = `${name}-offline.tgz`;
  execFileSync("tar", ["-czf", tgz, name], { cwd: stage, stdio: "inherit" });
  fs.mkdirSync(outDir, { recursive: true });
  const out = path.join(outDir, tgz);
  fs.copyFileSync(path.join(stage, tgz), out);
  console.log(`→ ${out} (${(fs.statSync(out).size / 1024 / 1024).toFixed(1)} MB)`);
  console.log("目标机器：解压后 node <解压目录>/bin/gongwen.js where 确认能跑，再按 README 链进 skill 目录。");
} finally {
  fs.rmSync(stage, { recursive: true, force: true });
}
