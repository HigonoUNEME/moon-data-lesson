# -*- coding: utf-8 -*-
"""教員用 当日資料（docs/teacher_guide_3koma.md）をPDFにする。

    python docs/build_teacher_guide_pdf.py

出力：docs/teacher_guide_3koma.pdf（A4）、docs/teacher_quickref_3koma.pdf（当日早見。A4・1枚）
mdの全文を機械的に変換する（md と PDF がずれないようにするため）。
当日早見は、md の最後の節「付録　当日早見」だけを A4・1枚に収めて別のPDFにする
（本文・表の文字は MIN_QUICK_PT＝9pt 以上。1枚に収まる最大の文字サイズを 12pt から探し、9pt でも収まらなければ、
また出来上がったPDFに9pt未満の文字があれば、エラーにする。収まらないときは md の節の内容を絞る。
教室で拾い読みできる大きさを下回らないため）。本文のPDFにも、最後の1ページとして（同じ早見のページをそのまま）つなぐ。
対応する記法：# 見出し（表題）、## / ### / #### 見出し、表、- 箇条書き（「- □」はチェック欄）、1. 番号つき、
> 引用、**強調**、`コード`、セル内の <br>。
docs/build_lesson_plan_pdf.py（指導案PDF）の変換の考え方を流用した、独立のスクリプト。
フォント：notebooks/assets/NotoSansJP-Regular.ttf（絵文字・「≒」「✔」は描画できないので、md側でも使わない）
"""
import io
import os
import re
import sys
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import CondPageBreak, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_PATH = os.path.join(ROOT, "docs", "teacher_guide_3koma.md")
FONT_PATH = os.path.join(ROOT, "notebooks", "assets", "NotoSansJP-Regular.ttf")
OUT_PATH = os.path.join(ROOT, "docs", "teacher_guide_3koma.pdf")
QUICK_PATH = os.path.join(ROOT, "docs", "teacher_quickref_3koma.pdf")
QUICK_HEAD = "## 付録　当日早見"
QUICK_TITLE = "当日早見（A4・1枚）：月データでムーンベースの場所を決めよう（全3コマ）"
MIN_QUICK_PT = 9.0     # 当日早見の最小の文字サイズ（pt）。これを下回るならエラー

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
S = {}
PAD = {"lr": 3.5, "top": 2.4, "bot": 2.6}


def set_scale(body, cell, scale=1.0, pad=None, floor=0.0):
    """本文の文字サイズ body・表の文字サイズ cell で、スタイルを作り直す（当日早見で文字を縮めるため）。"""
    global BODY, CELL
    BODY, CELL = body, cell
    # quote の文字は body - 0.4。当日早見では、これも floor 以上にする
    S.clear()
    S.update({
        "title": ParagraphStyle("title", fontName="NotoJP", wordWrap="CJK", fontSize=15 * scale, leading=20 * scale, textColor=INK,
                                spaceAfter=3 * scale),
        "h1": ParagraphStyle("h1", fontName="NotoJP", wordWrap="CJK", fontSize=11 * scale, leading=15 * scale, textColor=colors.white,
                             backColor=ACCENT, leftIndent=0, spaceBefore=9 * scale, spaceAfter=5 * scale,
                             borderPadding=(2.5, 4, 2.5, 5)),
        "h2": ParagraphStyle("h2", fontName="NotoJP", wordWrap="CJK", fontSize=max(9.6 * scale, floor), leading=13 * scale, textColor=ACCENT,
                             spaceBefore=6 * scale, spaceAfter=2.5 * scale),
        "h3": ParagraphStyle("h3", fontName="NotoJP", wordWrap="CJK", fontSize=max(8.8 * scale, floor), leading=12 * scale, textColor=INK,
                             spaceBefore=5 * scale, spaceAfter=2 * scale, leftIndent=0, borderPadding=(0, 0, 0, 0)),
        "body": ParagraphStyle("body", fontName="NotoJP", wordWrap="CJK", fontSize=body, leading=body * 1.5 if scale >= 1 else body * 1.38,
                               textColor=INK, spaceAfter=3 * scale),
        "quote": ParagraphStyle("quote", fontName="NotoJP", wordWrap="CJK", fontSize=max(body - 0.4, floor), leading=12.5 * scale, textColor=SUB,
                                leftIndent=8, spaceAfter=3 * scale),
        "cell": ParagraphStyle("cell", fontName="NotoJP", wordWrap="CJK", fontSize=cell, leading=cell * 1.5 if scale >= 1 else cell * 1.36,
                               textColor=INK),
        "cellhead": ParagraphStyle("cellhead", fontName="NotoJP", wordWrap="CJK", fontSize=cell + 0.2,
                                   leading=cell * 1.5 if scale >= 1 else cell * 1.36, textColor=ACCENT),
        "banner": ParagraphStyle("banner", fontName="NotoJP", wordWrap="CJK", fontSize=max(10.5 * scale, floor), leading=15 * scale, textColor=WARN,
                                 spaceAfter=4 * scale, borderColor=WARN, borderWidth=1, borderPadding=(3, 5, 3, 5)),
    })
    PAD.update(pad or {"lr": 3.5, "top": 2.4, "bot": 2.6})


set_scale(BODY, CELL)


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
        ("LEFTPADDING", (0, 0), (-1, -1), PAD["lr"]), ("RIGHTPADDING", (0, 0), (-1, -1), PAD["lr"]),
        ("TOPPADDING", (0, 0), (-1, -1), PAD["top"]), ("BOTTOMPADDING", (0, 0), (-1, -1), PAD["bot"]),
    ]
    for i in range(2, len(rows), 2):
        style.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    t.setStyle(TableStyle(style))
    return t


def to_story(lines, avail, quick=False):
    story, i = [], 0
    first_h1 = True
    gap = 1.2 * mm if quick else 2.5 * mm
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
            story.append(Spacer(1, gap))
            continue
        if s.startswith("#### "):
            story.append(CondPageBreak(30 * mm))
            story.append(Paragraph("<b>" + inline(s[5:]) + "</b>", S["h3"]))
        elif s.startswith("### "):
            story.append(CondPageBreak(38 * mm))   # 見出しだけがページ末に残らないように
            story.append(Paragraph(inline(s[4:]), S["h2"]))
        elif s.startswith("## "):
            if s[3:].startswith("付録") and not quick:
                story.append(PageBreak())          # 当日早見は、本文の最後に独立した1ページとして入れる
            else:
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


def footer_quick(canvas, doc):
    canvas.saveState()
    canvas.setFont("NotoJP", MIN_QUICK_PT)
    canvas.setFillColor(SUB)
    canvas.drawCentredString(A4[0] / 2, 4.2 * mm, "教員用 当日早見（生徒に配らない・映さない）")
    canvas.restoreState()


def read_md():
    with open(SRC_PATH, encoding="utf-8") as f:
        return f.read().split("\n")


def build(quick_pdf=None):
    """本文のPDF。最後の節（当日早見）は本文に入れず、build_quick() で作った1ページをつなぐ。"""
    set_scale(8.2, 7.2)
    lines = read_md()
    k = [i for i, s in enumerate(lines) if s.startswith(QUICK_HEAD)]
    if len(k) != 1:
        raise SystemExit("md に「%s」の節が1つだけない（%d）" % (QUICK_HEAD, len(k)))
    avail = A4[0] - 24 * mm
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=10 * mm, bottomMargin=13 * mm, leftMargin=12 * mm, rightMargin=12 * mm,
        title="教員用 当日資料：月データでムーンベースの場所を決めよう（全3コマ）",
    )
    story = to_story(lines[:k[0]], avail)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    from pypdf import PdfReader, PdfWriter
    w = PdfWriter()
    for src in (io.BytesIO(buf.getvalue()), io.BytesIO(quick_pdf) if quick_pdf else None):
        if src is None:
            continue
        for pg in PdfReader(src).pages:
            w.add_page(pg)
    w.add_metadata({"/Title": "教員用 当日資料：月データでムーンベースの場所を決めよう（全3コマ）"})
    with open(OUT_PATH, "wb") as f:
        w.write(f)
    print("wrote", OUT_PATH, "（本文%dページ＋当日早見1ページ）" % (len(PdfReader(io.BytesIO(buf.getvalue())).pages)))


def quick_lines():
    lines = read_md()
    k = [i for i, s in enumerate(lines) if s.startswith(QUICK_HEAD)]
    if len(k) != 1:
        raise SystemExit("md に「%s」の節が1つだけない（%d）" % (QUICK_HEAD, len(k)))
    rest = lines[k[0] + 1:]
    for j, s in enumerate(rest):          # 次の「## 」の節までを早見とする
        if s.startswith("## "):
            rest = rest[:j]
            break
    return rest


def min_font_size(pdf_bytes):
    """PDF中の文字（空白以外）の、実効の最小の文字サイズ（pt）を返す。"""
    from pypdf import PdfReader
    sizes = []

    def visitor(text, cm, tm, font_dict, font_size):
        if text.strip():
            sc = abs(cm[0] * tm[0] + cm[2] * tm[1]) if cm and tm else 1.0
            sc = sc if sc > 0 else 1.0
            sizes.append(round(float(font_size) * sc, 2))

    for pg in PdfReader(io.BytesIO(pdf_bytes)).pages:
        pg.extract_text(visitor_text=visitor)
    return min(sizes) if sizes else 0.0


def build_quick():
    """当日早見：MIN_QUICK_PT（9pt）以上で1枚に収まる最大の文字サイズを探して、A4・1枚のPDFにする。
    9ptでも1枚に収まらない／出来上がりに9pt未満の文字がある、ときはエラー（md の節を絞る）。"""
    rest = quick_lines()
    mar = 7 * mm
    avail = A4[0] - 2 * mar
    body_sizes = [round(MIN_QUICK_PT + x / 10, 1) for x in range(30, -1, -2)]   # 12.0 → 9.0 pt
    for body in body_sizes:
        cell = body
        set_scale(body, cell, scale=0.78, pad={"lr": 2.2, "top": 1.2, "bot": 1.3}, floor=MIN_QUICK_PT)
        story = [Paragraph(inline(QUICK_TITLE), S["title"]),
                 Paragraph("教員用（答えを含む。生徒に配らない）", S["banner"])] + to_story(rest, avail, quick=True)
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=7 * mm, bottomMargin=8 * mm, leftMargin=mar, rightMargin=mar,
                                title="教員用 当日早見（A4・1枚）")
        doc.build(story, onFirstPage=footer_quick, onLaterPages=footer_quick)
        if doc.page == 1:
            mn = min_font_size(buf.getvalue())
            if mn < MIN_QUICK_PT - 0.01:
                raise SystemExit("当日早見に %.1fpt の文字がある（最小 %.1fpt）。スタイルを直す" % (mn, MIN_QUICK_PT))
            with open(QUICK_PATH, "wb") as f:
                f.write(buf.getvalue())
            print("wrote", QUICK_PATH, "（本文%.1fpt・表%.1fpt・PDF中の最小%.1fpt。A4・1枚）" % (body, cell, mn))
            return buf.getvalue()
    raise SystemExit("当日早見が、%.0fpt でもA4・1枚に収まらない（%.0fpt未満にはしない）。md の「%s」の節の内容を絞る" % (MIN_QUICK_PT, MIN_QUICK_PT, QUICK_HEAD))


if __name__ == "__main__":
    qpdf = build_quick()
    build(qpdf)
