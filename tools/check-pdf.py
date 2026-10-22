#!/usr/bin/env python3
"""check-pdf.py —— 导出的 PDF 的字体与版面机检（公文体）。

一、字体。docx 里写的字体是 Word 渲染时解析的：机器上没装、或字体名对不上，
Word 会静默回退，而 docx 机检看不出来。只有 PDF 的内嵌字体能证明实际用了什么。
默认白名单就是本工具会声明的字体（GB/T 9704 各要素用字）：

  仿宋 FangSong（正文、发文字号、版记）／黑体 SimHei（一级层次、密级）／
  楷体 KaiTi（二级层次、签发人姓名）／宋体 SimSun（页码）／
  小标宋（标题）：

小标宋的实际字体是「方正小标宋简体」（目标机器装它）；macOS 之类的预览机
没有方正字体，本工具默认用 Songti SC Black 替代——替代名在下面单独放行。
22pt 以上含汉字的文字（标题、发文机关标志）另查一遍：只许小标宋或其替代，
回退成宋体、仿宋都报出来。--allow 加的字体同时放行正文与标题，
单位用的是别的小标宋版本时用它补上字体名。

二、版面。分页是 Word 排出来的，docx 里看不到：
  · 标题孤悬页底：章节标题是本页最后一行
  · 大片留白：本页上半截就结束，下一页也不是新章
  · 空白页：正文区一个字也没有

依赖 PyMuPDF。用法：
  ~/.venvs/pdf/bin/python3 tools/check-pdf.py <pdf> [--allow "字体1,字体2"]
退出码 0 通过，1 不通过，2 用法错。
"""
import re
import sys

try:
    import pymupdf
except ImportError:
    print("需要 PyMuPDF：uv run --with pymupdf tools/check-pdf.py <pdf>，或使用装了它的 venv", file=sys.stderr)
    sys.exit(2)

# 白名单：国标五类字体 + 方正小标宋的常见文件名 + macOS 预览替代。
ALLOW = [
    "FangSong", "SimHei", "KaiTi", "SimSun",
    "方正小标宋简体", "FZXiaoBiaoSong-B05S", "FZXiaoBiaoSong-B13S",
    "FZXBS",   # PostScript 名：简体 FZXBSJW--GB1-0、GBK 版 FZXBSK--GBK1-0
    "STSongti-SC-Black",
]
# 标题形状：小标宋／黑体／楷体，三号（16pt）以上
HEAD = ("方正小标宋简体", "FZXiaoBiaoSong", "FZXBS", "SimHei", "KaiTi", "STSongti-SC-Black")
HEAD_MIN = 15

# 版心上界与页面内正文底界（pt）。A4 高 842；页脚页码另算。
TOP, BOTTOM = 70, 800
# 正文底边高于这条线，就算下半截空着（版心约 710pt，留白超过约四成）。
BLANK_LINE = 480


def lines_of(page):
    """版心内有文字的行：(y0, y1, 首个 span, 整行文字)。"""
    out = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            spans = [s for s in line["spans"] if s["text"].strip()]
            if not spans:
                continue
            y0, y1 = line["bbox"][1], line["bbox"][3]
            if y0 < TOP or y1 > BOTTOM:
                continue
            out.append((y0, y1, spans, "".join(s["text"] for s in spans).strip()))
    out.sort(key=lambda r: r[0])
    return out


def is_heading(spans, min_size=HEAD_MIN):
    s = spans[0]
    return any(h in s["font"] for h in HEAD) and s["size"] >= min_size


# 标题与发文机关标志（二号 22pt 以上，降号下限也是 22pt）只许小标宋或其预览替代。
# SimSun 在白名单里是给页码用的：没装方正小标宋时 Word 会把标题静默回退成 SimSun，
# 只查白名单会放过它（实际发生过：标题与红头全是宋体，字体机检照样通过）。
# --allow 加进来的字体同时算合格的标题字体（单位用的其他小标宋版本）。
TITLE_FONTS = ["方正小标宋简体", "FZXiaoBiaoSong", "FZXBS", "STSongti-SC-Black"]
TITLE_MIN = 21.5
CJK = re.compile(r"[\u4e00-\u9fff]")


def check_fonts(doc):
    bad, title_bad, body_bad = {}, {}, {}
    for pno, page in enumerate(doc, 1):
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    if not span["text"].strip():
                        continue
                    font, hit = span["font"], (pno, span["text"][:16])
                    if not any(a in font for a in ALLOW):
                        bad.setdefault(font, []).append(hit)
                    # 只看含汉字的片段：×、数字这类西文字符走样式里的西文字体，不是回退
                    elif CJK.search(span["text"]) and span["size"] >= TITLE_MIN \
                            and not any(t in font for t in TITLE_FONTS):
                        title_bad.setdefault(font, []).append(hit)
                    # 三号（16pt）含汉字却用宋体：正文应为仿宋/黑体/楷体，宋体是标题
                    # 或页码的字体；正文回退成宋体会在这里被抓住。
                    elif CJK.search(span["text"]) and 15.5 <= span["size"] <= 16.6 and "SimSun" in font:
                        body_bad.setdefault(font, []).append(hit)
    msgs = [f"字体不在白名单　{f} ×{len(h)}，例：{h[0][1]!r}（p{h[0][0]}）" for f, h in bad.items()]
    if msgs:
        msgs.append("　白名单：" + "、".join(ALLOW))
    msgs += [f"标题字体回退　{f} ×{len(h)}，例：{h[0][1]!r}（p{h[0][0]}）；标题与发文机关标志应为方正小标宋，"
             f"本机可能没装（见 docs/字体.md）" for f, h in title_bad.items()]
    msgs += [f"正文字体回退　{f} ×{len(h)}，例：{h[0][1]!r}（p{h[0][0]}）；三号汉字应为仿宋/黑体/楷体，"
             f"宋体是页码或表格的字体" for f, h in body_bad.items()]
    return msgs


def opens_chapter(lines):
    """下一页是不是新章开头：在正文字号的行出现之前先看到标题行。"""
    for _, _, spans, _ in lines:
        if is_heading(spans):
            return True
        if max((s["size"] for s in spans if s["text"].strip()), default=0) >= 15:
            return False
    return False


def check_layout(doc):
    pages = [lines_of(p) for p in doc]
    msgs = []
    for i in range(len(pages)):
        ls, pno = pages[i], i + 1
        if not ls:
            # 末页无正文不算空白页（公文末页可能有版记落在浮动框架里）
            if i < len(pages) - 1:
                msgs.append(f"空白页　p{pno} 正文区没有内容")
            continue
        # 标题孤悬页底：本页最后一行是标题（末页不查——版记是浮动框，正文自然结束）
        if i < len(pages) - 1 and is_heading(ls[-1][2]):
            msgs.append(f"标题孤悬页底　p{pno}「{ls[-1][3][:20]}」是本页最后一行，正文在下一页")
        # 大片留白：下一页不是新章开头，本页却在上半截就结束了
        if i + 1 < len(pages) - 1:
            if ls[-1][1] < BLANK_LINE and not opens_chapter(pages[i + 1]):
                msgs.append(f"大片留白　p{pno} 正文止于 y={ls[-1][1]:.0f}pt，下半页空着，下一页也不是新章")
    return msgs


def main(path):
    doc = pymupdf.open(path)
    msgs = check_fonts(doc) + check_layout(doc)
    if msgs:
        print(f"PDF 机检 {len(msgs)} 处　{path}", file=sys.stderr)
        for m in msgs:
            print("  · " + m, file=sys.stderr)
        return 1
    print(f"PDF 机检通过（字体、版面）　{path}")
    return 0


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--allow" in args:
        i = args.index("--allow")
        if i + 1 >= len(args):
            sys.exit("--allow 后面要给逗号分隔的字体名")
        extra = [x.strip() for x in args[i + 1].split(",") if x.strip()]
        ALLOW.extend(extra)
        TITLE_FONTS.extend(extra)
        del args[i:i + 2]
    if len(args) != 1:
        print("用法：check-pdf.py <pdf> [--allow \"字体1,字体2\"]", file=sys.stderr)
        sys.exit(2)
    sys.exit(main(args[0]))
