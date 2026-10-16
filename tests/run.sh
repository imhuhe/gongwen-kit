#!/usr/bin/env bash
# gongwen-kit 测试入口：正案例（七份样张，writeDoc 内跑完整机检链）
# + 负对照（每道机检造一份必错的稿，断言它报出且非零退出）。
set -euo pipefail
cd "$(dirname "$0")/.."

echo "== 正案例：七份样张 =="
node templates/complete.js | tail -4
node templates/gongwen-qingshi.js | tail -4
node templates/gongwen-guanli.js | tail -4
node templates/notice.js | tail -4
node templates/xinhan.js | tail -4
node templates/jiyao.js | tail -4
node templates/qiye.js | tail -4
echo
echo "== 负对照 =="
node tests/negative.js
echo
echo "== 修复工具 =="
node tests/fix.js
echo
echo "== skill 入口 =="
node tests/skill-frontmatter.js
node tests/bin.js
node tests/pack.js
echo
echo "全部通过"
