# gongwen-kit

中文公文（GB/T 9704）Word 生成、修复与机检工具。本仓库同时是一个 agent skill：
入口与工作流在 [SKILL.md](SKILL.md)，先读它。

- 命令入口与路径约定见 SKILL.md「先定位工具包」，不要写相对路径。
- 改版式之前先读 [docs/版式.md](docs/版式.md)。
- 改代码后跑 `npm test`：七份样张、全部负对照、规则对照（tests/check-rules.js）、修复幂等（tests/fix.js）都要过。
- 机检不过就改稿子，不要改检查。
