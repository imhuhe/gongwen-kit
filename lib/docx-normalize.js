"use strict";
/**
 * docx-normalize.js —— 出稿后的最后一道包级整理（2026-10-23）。
 *
 * docx 库会先写一份默认 styles.xml，再拼 externalStyles。Title、Heading1～4、
 * FootnoteReference、EndnoteReference 等因此出现两个同 styleId 的定义；Word
 * 取后一份，WPS/LibreOffice 取哪一份不确定。这里按 styleId 去重，保留最后一份
 * （工具包样式），并在 docProps/core.xml 的 keywords 里打「gongwen-kit:版本」
 * 标记，供 fix-doc 识别自家出稿。
 *
 * 只改包里两个 XML 部件，不碰 document.xml；改动后用 jszip 重打包。
 */
const fs = require("fs");

function dedupeStyles(xml) {
  const blocks = [...xml.matchAll(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g)];
  if (!blocks.length) return xml;
  const last = new Map();
  for (const b of blocks) {
    const id = (b[0].match(/w:styleId="([^"]+)"/) || [])[1];
    if (id) last.set(id, b.index);
  }
  let out = "";
  let cursor = 0;
  for (const b of blocks) {
    const id = (b[0].match(/w:styleId="([^"]+)"/) || [])[1];
    if (id && last.get(id) !== b.index) {
      out += xml.slice(cursor, b.index);
      cursor = b.index + b[0].length;
    }
  }
  if (cursor === 0) return xml;
  out += xml.slice(cursor);
  return out;
}

function markCore(xml, version) {
  const marker = `gongwen-kit:${version}`;
  const existing = xml.match(/<cp:keywords>([\s\S]*?)<\/cp:keywords>/);
  if (existing) {
    const value = existing[1].trim();
    if (value.includes(marker)) return xml;
    const next = value ? `${value};${marker}` : marker;
    return xml.replace(existing[0], `<cp:keywords>${next}</cp:keywords>`);
  }
  if (/<cp:keywords\/>/.test(xml)) return xml.replace(/<cp:keywords\/>/, `<cp:keywords>${marker}</cp:keywords>`);
  return xml.replace(/<\/cp:coreProperties>/, `<cp:keywords>${marker}</cp:keywords></cp:coreProperties>`);
}

async function normalizeDocx(file, version, { mark = true } = {}) {
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  let changed = false;
  const stylesPart = zip.file("word/styles.xml");
  if (stylesPart) {
    const xml = await stylesPart.async("string");
    const next = dedupeStyles(xml);
    if (next !== xml) { zip.file("word/styles.xml", next); changed = true; }
  }
  const corePart = mark ? zip.file("docProps/core.xml") : null;
  if (corePart) {
    const xml = await corePart.async("string");
    const next = markCore(xml, version);
    if (next !== xml) { zip.file("docProps/core.xml", next); changed = true; }
  }
  if (changed) fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return changed;
}

module.exports = { dedupeStyles, markCore, normalizeDocx };
