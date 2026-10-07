# -*- coding: utf-8 -*-
"""教員用 当日資料（docs/teacher_guide_3koma.md）をPDFにする。

    python docs/build_teacher_guide_pdf.py

出力：docs/teacher_guide_3koma.pdf（A4）
mdの全文を機械的に変換する（md と PDF がずれないようにするため）。
対応する記法：# 見出し（表題）、## / ### / #### 見出し、表、- 箇条書き（「- □」はチェック欄）、1. 番号つき、
> 引用、**強調**、`コード`、セル内の <br>。
docs/build_lesson_plan_pdf.py（指導案PDF）の変換の考え方を流用した、独立のスクリプト。
フォント：notebooks/assets/NotoSansJP-Regular.ttf（絵文字・「≒」「✔」は描画できないので、md側でも使わない）
"""
import os
import re
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import CondPageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_PATH = os.path.join(ROOT, "docs", "teacher_guide_3koma.md")
FONT_PATH = os.path.join(ROOT, "notebooks", "assets", "NotoSansJP-Regular.ttf")
OUT_PATH = os.path.join(ROOT, "docs", "teacher_guide_3koma.pdf")

pdfmetrics.registerFont(TTFont("NotoJP", FONT_PATH))
pdfmetrics.registerFontFamily("NotoJP", normal="NotoJP", bold="NotoJP", italic="NotoJP", boldItalic="NotoJP")

INK = colors.HexColor("#1f2933")
SUB = colors.HexColor("#52606d")
ACCENT = colors.HexColor("#0f4c81")
WARN = colors.HexColor("#b42318")
LINE = colors.HexColor("#9aa5b1")
HEAD_BG = colors.HexColor("#e4ecf7")
ZEBRA = colors.HexColor("#f6f8fb")

BODY = 8.2
CELL = 7.2

S = {
    "title": ParagraphStyle("title", fontName="NotoJP", wordWrap="CJK", fontSize=15, leading=20, textColor=INK, spaceAfter=3),
    "h1": ParagraphStyle("h1", fontName="NotoJP", wordWrap="CJK", fontSize=11, leading=15, textColor=colors.white, backColor=ACCENT,
                         leftIndent=0, spaceBefore=9, spaceAfter=5, borderPadding=(2.5, 4, 2.5, 5)),
    "h2": ParagraphStyle("h2", fontName="NotoJP", wordWrap="CJK", fontSize=9.6, leading=13, textColor=ACCENT, spaceBefore=6,
                         spaceAfter=2.5),
    "h3": ParagraphStyle("h3", fontName="NotoJP", wordWrap="CJK", fontSize=8.8, leading=12, textColor=INK, spaceBefore=5,
                         spaceAfter=2, leftIndent=0, borderPadding=(0, 0, 0, 0)),
    "body": ParagraphStyle("body", fontName="NotoJP", wordWrap="CJK", fontSize=BODY, leading=BODY * 1.5, textColor=INK, spaceAfter=3),
    "quote": ParagraphStyle("quote", fontName="NotoJP", wordWrap="CJK", fontSize=BODY - 0.4, leading=12.5, textColor=SUB,
                            leftIndent=8, spaceAfter=3),
    "cell": ParagraphStyle("cell", fontName="NotoJP", wordWrap="CJK", fontSize=CELL, leading=CELL * 1.5, textColor=INK),
    "cellhead": ParagraphStyle("cellhead", fontName="NotoJP", wordWrap="CJK", fontSize=CELL + 0.2, leading=CELL * 1.5, textColor=ACCENT),
    "banner": ParagraphStyle("banner", fontName="NotoJP", wordWrap="CJK", fontSize=10.5, leading=15, textColor=WARN, spaceAfter=4,
                             borderColor=WARN, borderWidth=1, borderPadding=(3, 5, 3, 5)),
}


def inline(text: str) -> str:
    t = escape(text)
    t = t.replace("&lt;br&gt;", "<br/>")
    t = re.sub(r"\*\*(.+?)\*\*", r'<font color="#0f4c81"><b>\1</b></font>', t)
    t = re.sub(r"`([^`]+)`", r'<font backColor="#eef1f4">\1</font>', t)
    return t


def split_row(line: str):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def plain_len(s: str) -> int:
    """幅の見積もり用。<br> で分けた最長の行の長さ。"""
    s = re.sub(r"\*\*|`", "", s)
    return max((len(x) for x in s.split("<br>")), default=0)


def make_table(rows, avail):
    n = len(rows[0])
    lens = []
    for j in range(n):
        m = max(plain_len(r[j]) if j < len(r) else 0 for r in rows)
        lens.append(min(max(m, 9), 70) ** 0.75)
    tot = sum(lens)
    widths = [avail * x / tot for x in lens]
    data = []
    for i, r in enumerate(rows):
        r = r + [""] * (n - len(r))
        st = S["cellhead"] if i == 0 else S["cell"]
        data.append([Paragraph(inline(c), st) for c in r])
    t = Table(data, colWidths=widths, repeatRows=1)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.45, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (-1, 0), HEAD_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.5), ("RIGHTPADDING", (0, 0), (-1, -1), 3.5),
        ("TOPPADDING", (0, 0), (-1, -1), 2.4), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.6),
    ]
    for i in range(2, len(rows), 2):
        style.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    t.setStyle(TableStyle(style))
    return t


def to_story(lines, avail):
    story, i = [], 0
    first_h1 = True
    while i < len(lines):
        ln = lines[i]
        s = ln.strip()
        if not s or s == "---":
            i += 1
            continue
        if s.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if not re.match(r"^\|[\s:|-]+\|$", lines[i].strip()):
                    rows.append(split_row(lines[i]))
                i += 1
            story.append(make_table(rows, avail))
            story.append(Spacer(1, 2.5 * mm))
            continue
        if s.startswith("#### "):
            story.append(CondPageBreak(30 * mm))
            story.append(Paragraph("<b>" + inline(s[5:]) + "</b>", S["h3"]))
        elif s.startswith("### "):
            story.append(CondPageBreak(38 * mm))   # 見出しだけがページ末に残らないように
            story.append(Paragraph(inline(s[4:]), S["h2"]))
        elif s.startswith("## "):
            story.append(CondPageBreak(48 * mm))
            story.append(Paragraph(inline(s[3:]), S["h1"]))
        elif s.startswith("# "):
            story.append(Paragraph(inline(s[2:]), S["title"]))
            if first_h1:
                story.append(Paragraph("教員用（答えを含む。生徒に配らない）", S["banner"]))
                first_h1 = False
        elif s.startswith(">"):
            story.append(Paragraph(inline(s.lstrip("> ")), S["quote"]))
        else:
            m = re.match(r"^(\s*)([-*]|\d+\.)\s+(.*)$", ln)
            if m:
                depth = len(m.group(1)) // 2
                mark = "・" if m.group(2) in "-*" else m.group(2)
                text = m.group(3)
                if mark == "・" and text.startswith("□ "):
                    mark, text = "□", text[2:]
                while i + 1 < len(lines) and lines[i + 1].startswith("   ") \
                        and not re.match(r"^\s*([-*]|\d+\.)\s", lines[i + 1]) \
                        and not lines[i + 1].strip().startswith("|"):
                    i += 1
                    text += " " + lines[i].strip()
                st = ParagraphStyle("li%d" % depth, parent=S["body"], leftIndent=12 + depth * 12,
                                    firstLineIndent=-10, spaceAfter=1.6)
                story.append(Paragraph(mark + " " + inline(text), st))
            else:
                text = s
                while i + 1 < len(lines) and lines[i + 1].strip() \
                        and not re.match(r"^\s*([-*]|\d+\.|#|\||>)", lines[i + 1]) \
                        and lines[i + 1].strip() != "---":
                    i += 1
                    text += lines[i].strip()
                if text.startswith("**教員用（答えを含む。生徒に配らない）**"):
                    i += 1
                    continue   # 表題の下の赤枠に出している（二重にしない）
                story.append(Paragraph(inline(text), S["body"]))
        i += 1
    return story


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("NotoJP", 7.2)
    canvas.setFillColor(SUB)
    canvas.drawCentredString(A4[0] / 2, 8 * mm,
                             "教員用 当日資料（答えを含む。生徒に配らない）　%d" % doc.page)
    canvas.restoreState()


def build():
    with open(SRC_PATH, encoding="utf-8") as f:
        lines = f.read().split("\n")
    avail = A4[0] - 24 * mm
    doc = SimpleDocTemplate(
        OUT_PATH, pagesize=A4,
        topMargin=10 * mm, bottomMargin=13 * mm, leftMargin=12 * mm, rightMargin=12 * mm,
        title="教員用 当日資料：月データでムーンベースの場所を決めよう（全3コマ）",
    )
    story = to_story(lines, avail)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    print("wrote", OUT_PATH)


if __name__ == "__main__":
    build()
