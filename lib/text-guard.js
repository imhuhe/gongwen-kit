"use strict";
/**
 * text-guard.js —— 文本规则的保护范围（2026-10-23 立）。
 *
 * 规范层规则（异形词、中文里的半角标点）和修复工具必须对同一批「照录」
 * 文本让路，否则会出现「机检报错、修复却说这是原话不能改」的死循环：
 *   · 引号/书名号里的原话与文件名
 *   · 网址、邮箱
 *   · 文件名与扩展名：公文.docx 的「.」不是句末标点
 *   · 裸域名：中文.com 的「.」不是半角标点
 *
 * 修复、check-style、style-scan 共用这一个来源；新增一条保护规则只改这里。
 */

const FILE_EXT = "docx?|xlsx?|pptx?|pdf|wps|ofd|et|dps|zip|rar|7z|tar|gz|txt|md|csv|json|xml|ya?ml|html?|png|jpe?g|gif|svg|mp3|mp4";
const TLD = "com|cn|net|org|io|gov|edu|info|biz|me|co|top|xyz|site|online|tech|cloud|ai|tv|cc";

const PATTERNS = [
  /https?:\/\/[^\s，。；）]+/g,
  /[\w.+-]+@[\w-]+\.[\w.]+/g,
  /「[^」]*」/g,
  /“[^”]*”/g,
  /《[^》]*》/g,
  // 文件名：中文/英文/数字/下划线组成的名字 + 已知扩展名
  new RegExp(`[\\u4e00-\\u9fffA-Za-z0-9_-]+\\.(?:${FILE_EXT})\\b`, "gi"),
  // 裸域名：至少一个点，末段是常见 TLD
  new RegExp(`(?:[\\u4e00-\\u9fffA-Za-z0-9-]+\\.)+(?:${TLD})\\b`, "gi"),
];

/** 返回 [start, end) 区间数组，已按 start 排序。 */
function protectedRanges(text) {
  const out = [];
  for (const re of PATTERNS) {
    re.lastIndex = 0;
    for (const m of text.matchAll(re)) out.push([m.index, m.index + m[0].length]);
  }
  return out.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}

function inRanges(ranges, pos) {
  return ranges.some(([a, b]) => pos >= a && pos < b);
}

module.exports = { protectedRanges, inRanges, FILE_EXT, TLD };
