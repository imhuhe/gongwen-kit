# gongwen-kit

中文公文（GB/T 9704）Word 生成、修复与机检工具。本仓库同时是一个 agent skill：
入口与工作流在 [SKILL.md](SKILL.md)，先读它。

- 命令统一走 `node "$KIT/bin/gongwen.js" new|fix|check|pdf|where`，`$KIT` 是本仓库的根目录，
  从任何工作目录都能调，不要写相对路径。
- 改版式之前先读 [docs/版式.md](docs/版式.md)。
- 改代码后跑 `npm test`，七份样张、全部负对照和修复工具的测试都要过。
- 机检不过就改稿子，不要改检查。
