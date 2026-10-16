---
name: gongwen-kit
description: 中文公文（党政机关、央国企场景）Word 生成、修复与合规机检：按 GB/T 9704 生成标准 .docx——通知、报告、请示、函、纪要、管理办法等，含红头、发文字号、版记；出稿自动机检（体例、数字、文风）；现成 .docx 按国标修复，改动写成修订痕迹。不适用于合同、简历、普通商务报告等非公文的 Word 文档。Use when creating, fixing, formatting, or machine-checking Chinese official documents (gongwen) as Word files; not for general Word documents.
---

# gongwen-kit

## 什么时候用它

用户要起草、生成或检查中文公文（通知、请示、报告、函、纪要、管理办法、条例等）
的 Word（.docx）文件时。

**不用它的情形**：合同、简历、普通商务报告、PPT、Excel 等非公文，用通用的 Office 或 Word skill。
同时装了通用 Word skill 时，中文公文一律走本 skill：国标版式和数字机检只在这里，
不要用通用 skill 重做版式，也不要先用通用 skill 出稿再来套格式。

## 先定位工具包

本文件所在的目录就是工具包目录，下文记作 `$KIT`。你的工作目录通常是用户的项目，不是 `$KIT`，
所以命令里不要写相对路径。统一用入口脚本，它从任何目录都能调。

**`$KIT` 只是本文的记号，不是环境变量。** 执行命令前先把它换成 `SKILL.md` 所在目录的绝对路径
（取你加载这个 skill 时看到的路径），每条命令都写完整，不要依赖上一条命令里设过的变量。例如
`SKILL.md` 在 `/Users/me/.agents/skills/gongwen-kit/` 时：

```bash
node "/Users/me/.agents/skills/gongwen-kit/bin/gongwen.js" where   # 打印工具包目录，用来确认路径对不对
```

找不到这个目录就问用户，不要猜。下文 `templates/`、`docs/` 的相对路径都指 `$KIT` 下面。

## 工作流

1. **起稿源脚本**：`node "$KIT/bin/gongwen.js" new --form wen --out 稿子.js`
   （行文体 `wen`；条文体/办法条例 `zhang`）；或照 `templates/` 里最近的一份改。
2. **写正文**：只调构件——`letterhead`、`title`、`zhusong`、`p`、`h1`～`h4`、
   `table`、`luokuan`、`fujian`、`fuzhu`、`fujianPage`、`banji`。不手拼 docx 元素、不写 `new Paragraph`。
   长标题按意群分两行给 `title(["上行", "下行"])`；标题与正文都不加空格缩进（样式给）。
   所有构件的用法一页看全：`templates/complete.js`。
3. **出稿**：`node 稿子.js`。写盘后自动跑三道机检（体例、数字、文风）——
   输出全绿才算过，非零退出就去改，别改检查。数字机检分两档：「计数与列举」
   （「以下三点」与实际条数）等拦截项必须改；「提示」开头的正文取值对照不拦稿，
   核对后属实就改，不属实（不同主语、不同口径）就留着，不必为消提示改写原文。
   **同一份稿子最多改三轮。** 三轮后仍有拦稿命中，停下，把机检输出原样交给用户，
   说明哪几条没过、为什么没能改掉，不要再试。
   要逐条的结构化结果时用 `check --json`，看 `status`：`pass` 全过；`hit` 有拦稿命中，
   `blocking` 为 true 的必须改，为 false 的只是提示；`fault` 是工具故障，不是稿子的问题，
   先修环境，不要改稿子。
4. **交付**：.docx 给用户；要 PDF 就在 Word 里导出，再按需跑
   `node "$KIT/bin/gongwen.js" pdf 导出件.pdf` 核字体与版面。

只想体检、不改文件时：`node "$KIT/bin/gongwen.js" check 稿子.docx --strict`，体例、数字、文风三道依次跑完。

用户给的是**现成的 .docx**（手写稿、旧模板、别人发来的）要改格式时，不重写，走修复：
`node "$KIT/bin/gongwen.js" fix 稿子.docx`，输出「稿子-修订.docx」，原文件不动。版式直接改；
体例写成修订痕迹由用户逐条接受或拒绝；数字问题只插批注，不替用户改数。
把终端里的修复报告转述给用户，提醒在 Word「审阅」里过一遍修订和批注。细则见 `docs/修复.md`。

## 公文要素速查（GB/T 9704—2012）

- 要素顺序：版头（份号／密级和保密期限／紧急程度／发文机关标志／发文字号／签发人／
  红色分隔线）→ 标题 → 主送机关 → 正文 → 附件说明 → 发文机关署名和成文日期 →
  附注 → 版记
- 层次序数四层，h1～h4 自动编号：一、／（一）／1.／（1）
- 成文日期用阿拉伯数字标全；月、日不编虚位（10 月 5 日，不是 05）
- 发文字号：机关代字〔年份〕序号号；顺序号不加「第」字、不编虚位
- 主送机关顶格、末尾全角冒号；正文每段左空二字
- 上行文（请示、上报的报告）：给 `letterhead` 传 `signers`，发文字号自动到左侧、
  与最后一个签发人同排
- 信函格式（函，GB/T 9704 10.1）：节属性用 `xinhanSection({ agency })` 代替
  `sectionProps()`，正文开头用 `xinhanHead({ docNo, copyNo, secret, urgency })` 代替
  `letterhead`；版记用 `banji({ letter: true, copyTo })`（只留抄送）。照 `templates/xinhan.js`
- 纪要格式（GB/T 9704 10.3）：版头用 `jiyaoHead({ mark: "×××纪要", issue: "〔2026〕第 1 期" })`
  代替 `letterhead`；正文后接 `attendees([["出席", "…"], ["请假", "…"], ["列席", "…"]])`。
  照 `templates/jiyao.js`
- 不要红头的（企业内部报告、请示、工作方案，央国企日常多是这种）：不接 `letterhead`，
  从 `title` 起，标题可带单位名；版心与其余要素不变。照 `templates/qiye.js`
- 红头分「数字红头」（直接绘制）与「预印套打」（`preprinted: true` 只留空位）
- 需要抄送与印发信息时接 `banji({ copyTo, issuer, date })`；附件说明用
  `fujian([...])`（多个自动编号），成文日期之后的括注用 `fuzhu("（……）")`
- 附件另面编排：附件页接在落款、附注之后、版记之前，用
  `fujianPage({ n, title, body })`（标记顶格三号黑体、标题第三行居中、页码连续；
  单附件不编顺序号；不与正文一起装订的传 `detached: "×发〔2026〕1号"`）
- 落款默认加盖印章形态（右空四字）；无公章的传 `luokuan({ ..., sealed: false })`
- 结尾用语：请示用「妥否，请批示」，报告用「特此报告」
- 上手就照 `templates/complete.js`（全构件）与 `templates/notice.js`（真实场景）；
  **改版式之前先读 `docs/版式.md`**——校准值与坑都在那里

## 边界

不做内容与政治表述校对（判断和责任归写稿人）；不做 OFD、印章图像、流转审批。
详见 README「边界」一节。
