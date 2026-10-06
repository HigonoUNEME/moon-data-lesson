# -*- coding: utf-8 -*-
"""指導案（docs/lesson_plan_3koma_draft.md）のうち、「現場教員版」の節だけを抜き出してPDFにする。

    python docs/build_lesson_plan_pdf.py

出力：docs/lesson_plan_3koma.pdf
内容はmdから機械的に取り出す（PDFとmdがずれないようにするため）。どの節を取り出すかは
md冒頭の §0 の表と同じ：§1、§3.1、§3.3、§4、§5、§6.1・6.2・6.4・6.5、§7〜§9、§10.2。
内部向けの節（§2、§3.2、§6.3、§10.1・10.3・10.4、付録）は出さない。
フォント：notebooks/assets/NotoSansJP-Regular.ttf（絵文字は描画できないので、md側でも使わない）
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
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_PATH = os.path.join(ROOT, "docs", "lesson_plan_3koma_draft.md")
FONT_PATH = os.path.join(ROOT, "notebooks", "assets", "NotoSansJP-Regular.ttf")
OUT_PATH = os.path.join(ROOT, "docs", "lesson_plan_3koma.pdf")

pdfmetrics.registerFont(TTFont("NotoJP", FONT_PATH))
pdfmetrics.registerFontFamily("NotoJP", normal="NotoJP", bold="NotoJP",
                              italic="NotoJP", boldItalic="NotoJP")

INK = colors.HexColor("#1f2933")
SUB = colors.HexColor("#52606d")
ACCENT = colors.HexColor("#0f4c81")
LINE = colors.HexColor("#9aa5b1")
HEAD_BG = colors.HexColor("#e4ecf7")
CODE_BG = colors.HexColor("#eef1f4")

S = {
    "title": ParagraphStyle("title", fontName="NotoJP", fontSize=17, leading=22,
                             textColor=INK, alignment=1, spaceAfter=4),
    "subtitle": ParagraphStyle("subtitle", fontName="NotoJP", fontSize=9, leading=13.5,
                                textColor=SUB, alignment=1, spaceAfter=8),
    "h1": ParagraphStyle("h1", fontName="NotoJP", fontSize=12.5, leading=17,
                          textColor=colors.white, backColor=ACCENT,
                          leftIndent=6, spaceBefore=12, spaceAfter=7,
                          borderPadding=(3, 4, 3, 6), keepWithNext=1),
    "h2": ParagraphStyle("h2", fontName="NotoJP", fontSize=10.5, leading=15,
                          textColor=ACCENT, spaceBefore=8, spaceAfter=3, keepWithNext=1),
    "h3": ParagraphStyle("h3", fontName="NotoJP", fontSize=9.5, leading=14,
                          textColor=INK, spaceBefore=5, spaceAfter=2, keepWithNext=1),
    "body": ParagraphStyle("body", fontName="NotoJP", fontSize=9, leading=14,
                            textColor=INK, spaceAfter=4),
    "quote": ParagraphStyle("quote", fontName="NotoJP", fontSize=8.5, leading=13,
                             textColor=SUB, leftIndent=10, spaceAfter=2),
    "cell": ParagraphStyle("cell", fontName="NotoJP", fontSize=8, leading=11.5,
                            textColor=INK),
    "cellhead": ParagraphStyle("cellhead", fontName="NotoJP", fontSize=8.2, leading=11.5,
                                textColor=ACCENT),
}

# 現場教員版として出す行範囲は、見出しから決める（行番号に依存しない）
INCLUDE = ("1.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.")  # ## 見出しの番号
SKIP_SUB = ("3.2", "6.3", "10.1", "10.3", "10.4")  # ### 見出しのうち内部向け


def inline(text: str) -> str:
    """md の **強調** と `コード` を reportlab のタグに直す。"""
    t = escape(text)
    t = re.sub(r"\*\*(.+?)\*\*", r'<font color="#0f4c81"><b>\1</b></font>', t)
    t = re.sub(r"`([^`]+)`", r'<font backColor="#eef1f4">\1</font>', t)
    return t


def clean(text: str) -> str:
    """内部の節への参照など、現場教員版のPDFでは意味をなさない記述を除く。"""
    text = text.replace("（現場教員版）", "")
    text = re.sub(r"（内部：§[0-9.]+）", "", text)
    text = re.sub(r"（内部資料：付録[A-Z]）", "", text)
    text = re.sub(r"（§10\.1の\(a\)）", "", text)
    text = re.sub(r"（§10\.1[^）]*）", "", text)
    text = re.sub(r"（§10\.[134]）", "", text)
    text = re.sub(r"。?§10\.[134]", "", text)  # 内部の節への参照
    text = text.replace("（　）は根拠コード。", "")
    text = re.sub(r"（理[(アイ配][^（）]*）", "", text)  # 要領項目の根拠コード（対応表は内部の§2.4）
    return text


def split_row(line: str):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def make_table(rows, avail):
    n = len(rows[0])
    lens = []
    for j in range(n):
        m = max(len(re.sub(r"\*\*|`", "", r[j])) if j < len(r) else 0 for r in rows)
        lens.append(min(max(m, 4), 60) ** 0.8)
    tot = sum(lens)
    widths = [avail * x / tot for x in lens]
    data = []
    for i, r in enumerate(rows):
        r = r + [""] * (n - len(r))
        st = S["cellhead"] if i == 0 else S["cell"]
        data.append([Paragraph(inline(clean(c)), st) for c in r])
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (-1, 0), HEAD_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def selected_lines(lines):
    """§の見出しで、現場教員版の節だけを残す。§0・§2・付録は落とす。"""
    out, keep, h2 = [], False, ""
    for ln in lines:
        m2 = re.match(r"## (\S+)", ln)
        m3 = re.match(r"### (\S+)", ln)
        if ln.startswith("# "):
            keep = False
            continue
        if m2:
            h2 = m2.group(1)
            keep = h2.startswith(INCLUDE) and not h2.startswith("2.") and "付録" not in ln
        elif m3:
            keep = (m3.group(1) not in SKIP_SUB) and h2.startswith(INCLUDE)
        if keep:
            out.append(ln)
    return out


def to_story(lines, avail):
    story, i = [], 0
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
            story.append(Spacer(1, 3 * mm))
            continue
        if s.startswith("#### "):
            story.append(Paragraph(inline(clean(s[5:])), S["h3"]))
        elif s.startswith("### "):
            story.append(Paragraph(inline(clean(s[4:])), S["h2"]))
        elif s.startswith("## "):
            story.append(Paragraph(inline(clean(s[3:])), S["h1"]))
        elif s.startswith(">"):
            story.append(Paragraph(inline(s.lstrip("> ")), S["quote"]))
        else:
            m = re.match(r"^(\s*)([-*]|\d+\.)\s+(.*)$", ln)
            if m:
                depth = len(m.group(1)) // 2
                mark = "・" if m.group(2) in "-*" else m.group(2)
                text = m.group(3)
                # 続く字下げ行（同じ項目の続き）をまとめる
                while i + 1 < len(lines) and lines[i + 1].startswith("   ") \
                        and not re.match(r"^\s*([-*]|\d+\.)\s", lines[i + 1]) \
                        and not lines[i + 1].strip().startswith("|"):
                    i += 1
                    text += " " + lines[i].strip()
                st = ParagraphStyle("li%d" % depth, parent=S["body"], leftIndent=12 + depth * 12,
                                    firstLineIndent=-10, spaceAfter=2)
                story.append(Paragraph(mark + " " + inline(clean(text)), st))
            else:
                text = s
                while i + 1 < len(lines) and lines[i + 1].strip() \
                        and not re.match(r"^\s*([-*]|\d+\.|#|\||>)", lines[i + 1]) \
                        and lines[i + 1].strip() != "---":
                    i += 1
                    text += lines[i].strip()
                if clean(text).strip():
                    story.append(Paragraph(inline(clean(text)), S["body"]))
        i += 1
    return story


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("NotoJP", 7.5)
    canvas.setFillColor(SUB)
    canvas.drawCentredString(A4[0] / 2, 9 * mm,
                             "学習指導案（現場教員版・案）　%d" % doc.page)
    canvas.restoreState()


def build():
    with open(SRC_PATH, encoding="utf-8") as f:
        lines = f.read().split("\n")
    avail = A4[0] - 36 * mm
    doc = SimpleDocTemplate(
        OUT_PATH, pagesize=A4,
        topMargin=15 * mm, bottomMargin=16 * mm, leftMargin=18 * mm, rightMargin=18 * mm,
        title="学習指導案（現場教員版・案）：月データでムーンベースの場所を決めよう",
    )
    story = [
        Paragraph("学習 指 導 案（現場教員版・案）", S["title"]),
        Paragraph(
            "探究基礎「月データでムーンベースの場所を決めよう」（全3時間）<br/>"
            "未実施。時間配分・操作時間はすべて机上の見積もりで、リハーサルによる検証は未実施。"
            "準拠する学習指導要領の科目は理数探究基礎。評価規準は公式文言ではなく案。<br/>"
            "節番号は元の指導案のままで、欠番の節（根拠の詳細・所見・検証経緯など）は内部資料のため省いている。",
            S["subtitle"]),
    ]
    story += to_story(selected_lines(lines), avail)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    print("wrote", OUT_PATH)


if __name__ == "__main__":
    build()
