"use strict";
/**
 * text-rules.js —— 体例与规范层文风的自动修复规则（tools/fix-doc.js 用）。
 *
 * 只收「改法唯一」的：命中即错、怎么改没有第二种答案。改法要靠判断的
 * （年份简写补哪个世纪、日期顿号、数值范围补什么单位、1.1 式编号怎么拆层次、
 * 「我司」改「我方」还是「我单位」）一律不改，交给批注。
 *
 * 每条规则吃段落文本，吐改动 [{ start, end, text, rule }]；start === end 是插入。
 * 判据与 checks/ 下的机检同源：机检报什么，这里改什么，改不了的机检会再报出来。
 */
const { loadTsv } = require("../../checks/style-scan");
const { protectedRanges } = require("../text-guard");

const CJK = "\\u4e00-\\u9fff\\u3000-\\u303f";
const CJK_RE = new RegExp(`[${CJK}]`);
const isCjk = (ch) => !!ch && CJK_RE.test(ch);
// 词表只读一次（每段都读一遍，上万段的文件会慢下来）
let yixing;
const yixingRows = () => (yixing = yixing || loadTsv("yixingci.tsv", 4, { required: true }));
const CN_DIGIT = { 〇: 0, 零: 0, 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };

// 「十」「十二」「二十」「二十三」→ 数
function cnSmall(s) {
  if (!s.includes("十")) return CN_DIGIT[s];
  const [a, b] = s.split("十");
  return (a ? CN_DIGIT[a] : 1) * 10 + (b ? CN_DIGIT[b] : 0);
}

const RULES = [
  {
    name: "月日编虚位",
    edits(text) {
      const out = [];
      for (const m of text.matchAll(/(年\s*)0(?=\d\s*月)|(月\s*)0(?=\d\s*日)/g)) {
        const at = m.index + (m[1] || m[2]).length;
        out.push({ start: at, end: at + 1, text: "" });
      }
      return out;
    },
  },
  {
    name: "发文字号",
    edits(text) {
      const out = [];
      // 年份括号：只在发文字号的形状里改（后面跟顺序号和「号」）
      for (const m of text.matchAll(/([【[])(\d{4})([】\]])(?=\s*第?\s*\d+\s*号)/g)) {
        out.push({ start: m.index, end: m.index + 1, text: "〔" });
        out.push({ start: m.index + 5, end: m.index + 6, text: "〕" });
      }
      for (const m of text.matchAll(/([〔【[]\d{4}[〕】\]])(\s*第\s*)(?=\d+\s*号)/g)) {
        const at = m.index + m[1].length;
        out.push({ start: at, end: at + m[2].length, text: "" });
      }
      for (const m of text.matchAll(/([〔【[]\d{4}[〕】\]]\s*(?:第\s*)?)(0+)(?=[1-9]\d*\s*号)/g)) {
        const at = m.index + m[1].length;
        out.push({ start: at, end: at + m[2].length, text: "" });
      }
      return out;
    },
  },
  {
    name: "中文数字日期",
    edits(text, ctx) {
      const out = [];
      const re = /([〇零一二三四五六七八九]{4})(\s*)年(?:(\s*)([一二三四五六七八九十]{1,3})(\s*)月(?:(\s*)([一二三四五六七八九十]{1,3})(\s*)日)?)?/g;
      for (const m of text.matchAll(re)) {
        if (!m[1].includes("〇")) continue;   // 判据同机检：只认带〇的年份
        const sp = ctx.spacedDates ? " " : "";
        let s = [...m[1]].map((c) => CN_DIGIT[c]).join("") + sp + "年";
        const mo = m[4] && cnSmall(m[4]);
        const d = m[7] && cnSmall(m[7]);
        if (m[4] && !(mo >= 1 && mo <= 12)) continue;
        if (m[7] && !(d >= 1 && d <= 31)) continue;
        if (m[4]) s += sp + mo + sp + "月";
        if (m[7]) s += sp + d + sp + "日";
        out.push({ start: m.index, end: m.index + m[0].length, text: s });
      }
      return out;
    },
  },
  {
    name: "全角数字",
    edits(text) {
      return [...text.matchAll(/[０-９]+/g)].map((m) => ({
        start: m.index, end: m.index + m[0].length,
        text: [...m[0]].map((c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).join(""),
      }));
    },
  },
  {
    name: "半角括号",
    edits(text) {
      const out = [];
      const stack = [];
      const pairs = [];
      for (let i = 0; i < text.length; i++) {
        if (text[i] === "(") stack.push(i);
        else if (text[i] === ")") pairs.push([stack.length ? stack.pop() : -1, i]);
      }
      stack.forEach((i) => pairs.push([i, -1]));
      for (const [a, b] of pairs) {
        const inner = a >= 0 && b >= 0 ? text.slice(a + 1, b) : "";
        const near = (a >= 0 && (isCjk(text[a - 1]) || isCjk(text[a + 1])))
          || (b >= 0 && (isCjk(text[b - 1]) || isCjk(text[b + 1])))
          || CJK_RE.test(inner);
        if (!near) continue;
        if (a >= 0) out.push({ start: a, end: a + 1, text: "（" });
        if (b >= 0) out.push({ start: b, end: b + 1, text: "）" });
      }
      return out;
    },
  },
  {
    name: "半角标点",
    edits(text) {
      const MAP = { ",": "，", ";": "；", ":": "：", "!": "！", "?": "？", ".": "。" };
      const out = [];
      for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (!MAP[c]) continue;
        const prev = text[i - 1], next = text[i + 1];
        if (!isCjk(prev) && !isCjk(next) && !(c === "." && /[)）」”]/.test(prev || ""))) continue;
        if (c === ".") {
          // 「1.××」的序号点、省略号、句点前不是中文的（缩写）都不动
          // 后面跟数字的是小数点（纯小数归下面一条），前面是右括号、右引号的照样是句末
          const closer = /[)）」”]/.test(prev || "");
          if (/\d/.test(prev || "") || /\d/.test(next || "") || prev === "." || next === "."
              || !(isCjk(prev) || closer)) continue;
        }
        if (c === ":" && /\d/.test(prev || "") && /\d/.test(next || "")) continue;   // 10:30
        out.push({ start: i, end: i + 1, text: MAP[c] });
      }
      return out;
    },
  },
  {
    name: "异形词",
    edits(text) {
      const out = [];
      for (const row of yixingRows()) {
        for (const bad of row[1].split("｜").map((w) => w.trim()).filter(Boolean)) {
          let i = text.indexOf(bad);
          while (i >= 0) {
            out.push({ start: i, end: i + bad.length, text: row[0] });
            i = text.indexOf(bad, i + bad.length);
          }
        }
      }
      return out;
    },
  },
  {
    name: "数字与单位",
    edits(text) {
      const re = /\d(?=(?:mAh\/g|GWh|kWh|MWh|kg|kW|MW|mm|cm|km|ppm|t\/h|g\/cm|Wh)(?![A-Za-z]))/g;
      return [...text.matchAll(re)].map((m) => ({ start: m.index + 1, end: m.index + 1, text: " " }));
    },
  },
  {
    name: "纯小数",
    edits(text) {
      // 点号前紧挨汉字或左括号才改。前面是空格的不改：「¥    .00 元」是留空待填，
      // 「99 .8%」是排版断开的数，补 0 都是错的（真实稿回放见过这两种）
      return [...text.matchAll(/(^|[^\d\w.])\.(?=\d)/g)]
        .filter((m) => !m[1] || isCjk(m[1]) || /[（(：:，]/.test(m[1]))
        .map((m) => ({ start: m.index + m[1].length, end: m.index + m[1].length, text: "0" }));
    },
  },
  {
    name: "标题重复编号",
    edits(text, ctx) {
      if (!ctx.heading || !ctx.numbered) return [];
      const m = text.match(/^(?:[一二三四五六七八九十]+[、．.]|（[一二三四五六七八九十]+）|\d{1,2}[.．、](?!\d))\s*/);
      return m ? [{ start: 0, end: m[0].length, text: "" }] : [];
    },
  },
];

/** 一段文字的全部改动：按位置排序、去掉落在引号与网址里的、重叠的保留先出现的规则。 */
function editsFor(text, ctx = {}) {
  const guard = protectedRanges(text);
  const inGuard = (e) => guard.some(([a, b]) => e.start < b && Math.max(e.end, e.start + 1) > a);
  const all = [];
  RULES.forEach((r, ri) => {
    for (const e of r.edits(text, ctx)) if (!inGuard(e)) all.push({ ...e, rule: r.name, ri });
  });
  all.sort((a, b) => a.start - b.start || a.ri - b.ri);
  const out = [];
  for (const e of all) {
    const last = out[out.length - 1];
    if (last && (e.start < last.end || (e.start === last.start && last.start === last.end))) continue;
    out.push(e);
  }
  return out;
}

module.exports = { editsFor, RULES };
