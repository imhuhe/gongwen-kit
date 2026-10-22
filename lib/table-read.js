"use strict";
/**
 * table-read.js —— docx 数据表单元格读取的公共部分（2026-09-14 立，2026-10-23 加固）。
 *
 * 为什么抽出来：`check-num.js` 与 `check-doc.js` 各写了一遍 `<w:tc>` 解析，
 * 同一个合并格（w:gridSpan）盲区因此要修两遍——「合计行少一格」的假阳性
 * 就是这么来的（check-num 先踩，check-doc 同款未爆）。读取只有一份源，判据各自保留。
 *
 * **合并格按真实列位展开**：合并格占其起始列，其后补 null（不是空串——空串
 * 在 check-doc 里是「该写一字线的空格」，null 是「这里没有格」）。
 * **纵向合并（w:vMerge）的续行不是空格**：续行补 null，不参与「表格空格」与
 * 数字求和；否则一张正确的纵合并表会被判成空格，合计也会算错。
 *
 * 表格/单元格开标签一律允许属性（Word 会写 w:rsidR 等）：`<w:tbl\b>`、
 * `<w:tc\b>`。2026-10-23 审计里 `<w:tbl w:rsidR>` 曾让整张表对机检静默消失。
 */

// 一个单元格碎片里的可见文本。`<w:tab/>` 还原成空格。
function textOf(frag) {
  return (frag.match(/<w:t[^>]*>[^<]*<\/w:t>|<w:tab\/>/g) || [])
    .map((m) => (m === "<w:tab\/>" ? " " : m.replace(/<[^>]+>/g, ""))).join("").trim();
}

// 全文表格区间。检查与修复都用它，避免各自写 `<w:tbl>` 精确匹配。
function tableRanges(xml) {
  const out = [];
  const re = /<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g;
  let m;
  while ((m = re.exec(xml))) out.push([m.index, m.index + m[0].length]);
  return out;
}

// 把 <w:tr> 拆成 [{ tc, span, vmerge }]。vmerge: "restart" / "continue" / ""。
function splitCells(tr) {
  return (tr.match(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g) || []).map((tc) => {
    const vm = tc.match(/<w:vMerge(?:\s[^>]*)?\/>/);
    let vmerge = "";
    if (vm) {
      const val = (vm[0].match(/w:val="([^"]+)"/) || [])[1] || "continue";
      vmerge = val === "restart" ? "restart" : "continue";
    }
    return {
      tc,
      span: Number((tc.match(/<w:gridSpan[^>]*w:val="(\d+)"/) || [])[1] || 1),
      vmerge,
    };
  });
}

// 一行展开成按真实列位排布的数组：mapFn(tcFragment) 取值，合并位补 null。
function rowCells(tr, mapFn) {
  const cells = [];
  for (const { tc, span, vmerge } of splitCells(tr)) {
    if (vmerge === "continue") {
      for (let s = 0; s < span; s++) cells.push(null);
      continue;
    }
    cells.push(mapFn(tc));
    for (let s = 1; s < span; s++) cells.push(null);
  }
  return cells;
}

module.exports = { textOf, tableRanges, splitCells, rowCells };
