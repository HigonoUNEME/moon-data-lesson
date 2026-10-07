# -*- coding: utf-8 -*-
"""3コマ版の紙教材（B1〜B5・B4c・B8・B9）をPDFで出力し、検査する。

    python course/build_materials.py            # 出力して検査する
    python course/build_materials.py build      # 出力だけ
    python course/build_materials.py check      # 既存PDFと本文データを検査だけ

出力先：course/materials/*.pdf
  B1  個人用ワークシート        B5・表裏（2ページ）
  B2  班用ワークシート          A3・表裏（2ページ・横）
  B3  ミッションカード4種       A5・横（4ページ）
  B4a 地域カード8件             A4・縦（2ページ、1ページ4枚）
  B4b 地域カードの全球図（番号）A4・横（1ページ）
  B4c クレーターカード（第2時。4枚）A4・縦（1ページ、A6×4。数値は書かない）
  B5a 黒板用の記入表            A3・横
  B5b 付箋用の月全球図          A3・横
  B8  班編成・ペア確定シート    A4・縦（教員用。2ページ：1＝班編成・ペア、2＝第2時の担当クレーターの割当）
  B9  Web操作カード             A5・縦

方針：紙は空欄を主とし、答え（半球・地域名とミッションの対応、標準の重み、想定値）は書かない。
地域カードの名前・緯度経度の範囲は course/course_moonbase.xlsx の「データ_地域」から読む
（data/candidate_regions.csv とも照合する）。中立な1行は、data/ の公開データで確認した事実だけを書く
（REGION_LINES の根拠は _FACT_BASIS に残す）。
フォント：notebooks/assets/NotoSansJP-Regular.ttf（絵文字は描画できない）。
"""
import csv
import os
import re
import subprocess
import sys
import tempfile
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_PATH = os.path.join(ROOT, "notebooks", "assets", "NotoSansJP-Regular.ttf")
MOON_IMG = os.path.join(ROOT, "notebooks", "assets", "lroc_color_2k.jpg")
XLSX_PATH = os.path.join(ROOT, "course", "student", "course_moonbase_student.xlsx")   # データ_地域（表示名の基準）
CSV_PATH = os.path.join(ROOT, "course", "data", "candidate_regions.csv")              # 表示名（学習者向け）
CSV_ORIG_PATH = os.path.join(ROOT, "data", "candidate_regions.csv")                   # 元データ（名前だけ旧名の地域がある）
# 元データ（data/）と表示名（course/data・データ_地域）の名前の対応。範囲の数値は同じ。
NAME_ALIAS = {"裏側・赤道（月の裏側の赤道帯）": "裏側・赤道（電波天文の候補域）"}
OUT_DIR = os.path.join(ROOT, "course", "materials")

FONT = "NotoJP"
pdfmetrics.registerFont(TTFont(FONT, FONT_PATH))
pdfmetrics.registerFontFamily(FONT, normal=FONT, bold=FONT, italic=FONT, boldItalic=FONT)

INK = colors.HexColor("#1f2933")
SUB = colors.HexColor("#52606d")
ACCENT = colors.HexColor("#0f4c81")
LINE = colors.HexColor("#7b8794")
RULE = colors.HexColor("#b8c2cc")
HEAD_BG = colors.HexColor("#e4ecf7")
SOFT_BG = colors.HexColor("#f4f6f8")
MARK = colors.HexColor("#d64545")
WHITE = colors.white

# 用紙（mm）。(幅, 高さ)
B5 = (176.0, 250.0)
A3L = (420.0, 297.0)
A4P = (210.0, 297.0)
A4L = (297.0, 210.0)
A5P = (148.0, 210.0)
A5L = (210.0, 148.0)

# =====================================================================
# 文言（変わりうるものはここに集める）
# =====================================================================

# --- Web操作カード（B9）と B2 の「探し方」の文言。Webの改修（W1〜W4）の実装どおり（2026-10-06 担当Wの完了報告）。
#     Webが変わったら、ここの文言だけを直す。
WEB_TEXT = {
    "start": "起動直後は月の実写だけ。データ層はオフ。",
    "btn_play": "自転の再生／停止（地球も一緒に空を巡る。はじめは×1.0で、1周約60秒）",
    "btn_image": "月の実写画像のオン／オフ（データの色を読むときはオフ）",
    "btn_layer": "データ層のオン／オフ（温度計のボタン）",
    "btn_wire": "ワイヤーフレーム",
    "btn_zoom": "拡大・縮小",
    "btn_reset": "視点のリセット",
    "buttons_note": "ボタンは絵だけ。カーソルを置くと名前が出る（タッチ操作では出ない）。",
    "layer_panel": "データ層をオンにすると、左上に凡例パネルが出る。パネルのプルダウンで層を選ぶ。一度に重ねられる層は1つだけ。",
    "layers": "1日の温度差／夜の最低温度／正午の太陽高度／地球の仰角／全球の傾斜／地質年代／標高／1日の温度（時刻で動く）",
    "diurnal": "「1日の温度」層は、左下のスライダーか右下の再生ボタンで時刻が進む。スライダーは太陽の真下の経度で、ピンを立てた地点の現地時刻ではない。",
    "hover": "月面にカーソルを合わせると、緯度経度と、いま選んでいる層の値が出る。",
    "coord_input": "画面上部中央の「緯度・経度」欄に値を入れて「ピンを立てる」（Enterでも可）。入力値は3°格子の中心（…−1.5、1.5、4.5…）にそろい、ピン・展開図・カードが同じ点を指し、月がその地点へ向く。緯度は−90〜90、経度は−180〜180（東経＋・西経−）。範囲外・空欄・数字以外はメッセージが出て、ピンは動かない。",
    "pin_click": "月面をクリックしても、ピンを立てられる（ピンは同時に1つ）。クリックすると、座標欄にそのピンの緯度・経度が入る。「ピンを立てる」を押すと、3°格子の中心にそろえ直せる。",
    "pin_card": "右上のピンカードに、全指標の表（1日の温度差・夜の最低温度・正午の太陽高度・地球の仰角・傾斜・地質年代・標高）が出る。層を切り替えても表は残り、選んでいる層の行だけ色が付く。表の下に、値を読んだ3°格子セルの中心が出る。",
    "pin_card_diurnal": "「1日の温度」層をオンにすると、ピンカードに24時間の最高・最低・差が出る。",
    "follow": "カードの「追従」をオンにすると、自転してもカメラが同じ経度を追う。×でピンを消す。",
    "minimap": "左下の展開図：水色は今見えている範囲、黄色の点がピン。",
    "place_hint": "画面上部中央の「地点名ヒント」スイッチは、はじめはオフ（既知の地点の名前もカテゴリも出ず、地点の近くをクリックしてもカメラは寄らない）。指示があったときだけオンにする。",
    "data_list": "右上の「データ一覧」は別のタブで開く。月の向き・ピン・層は、元のタブに残る。",
    "temp_unit": "温度は℃で表示される。表計算はK。温度の差は同じ大きさ。絶対温度は K − 273.15 ＝ ℃（例：約95 K は約−178℃）。",
    "cell": "値は3°×3°のセルの値。表計算とは位置が少しずれる。傾斜は、データのない地点（南極に近い点）で「―（欠測）」と出る。",
    "narrow": "画面の幅が約1100px未満のときは、ピンカードが座標入力の帯の下に下がる。",
    "jigsaw": "ジグソー（第1時）：担当地点（A〜E）の緯度・経度を座標欄に入れて、ピンカードの表の「1日の温度差」「夜の最低温度」「正午の太陽高度」「地球の仰角」の4行を読む（層は切り替えなくてよい）。「1日の温度」層をオンにしたときに出る24時間の差とは別の行。",
    # B2（班用ワークシート）の「探し方」注
    "find_site_b2": "探し方：Webの画面上部中央の「緯度・経度」欄に、担当地点の値を入れて「ピンを立てる」。ピンカードの表から、4つの値を読む（層を切り替えなくてよい）。",
}

# --- 個人用ワークシート（B1）の記録欄の名前。検査（check）がPDF本文と突き合わせる。
L_PREDICT_TEMP = "月の昼と夜の温度差の予想"
L_PREDICT = "予想（地域タイプ・場所・理由）"
L_NICHI = "日較差（赤道帯・極付近）"
L_DENSITY = "海・陸の密度と倍率"
L_SEARIKU = "海と陸の違い（1文）"
L_MEMO = "第2時の判断メモ（画面の数字と理由を書く）"
L_DECIDE = "最終決定（地域タイプ・1地点）"
L_REASON = "選んだ理由（2文＋弱点1つ）"
L_WEIGHTCHG = "重みを変えたときの変化（1行）"
L_QA = "質疑メモ"
L_REFLECT = "振り返り（予想と最終決定のずれ・理由）"
L_DOUBT = "このデータで信じきれないこと"
# ⑥ の理由欄（判断の理由を一言。評価しない。地域タイプと線の内外の対応は紙に書かない）
L_LINE_REASON = "線をそこに引いた理由（数字1つ）"
L_IO_REASON = "内か外か、そう決めた理由（数字1つ）"
B1_FIELDS = [L_PREDICT_TEMP, L_PREDICT, L_NICHI, L_DENSITY, L_SEARIKU, L_MEMO,
             L_DECIDE, L_REASON, L_WEIGHTCHG, L_QA, L_REFLECT, L_DOUBT]

QUESTION_TYPES = ["なぜその指標？", "別の地域は考えた？", "そのデータは信じてよい？"]
REASON_TYPE = "〜のミッションなので〜を重視した。〜の地域タイプで〜に決めた。弱点は〜。"
K_CELSIUS = "温度の差は、Kでも℃でも同じ大きさ。絶対温度は K − 273.15 ＝ ℃（Webは℃、表計算はK）。"

# --- ジグソーの5地点（指導案 §7.1。Webの3°グリッド中心）。値は書かない。
JIGSAW_SITES = [
    ("A", "赤道の海", 7.5, 31.5),
    ("B", "裏側・赤道", 1.5, 169.5),
    ("C", "南極付近", -88.5, 58.5),
    ("D", "中緯度の火砕丘", 22.5, -49.5),
    ("E", "裏側・南極エイトケン", -43.5, 175.5),
]

# --- ミッションカード（B3）。本文は答えを含まない目的の一文＋問い。ヒントは氷採掘だけ。
ICE_HINT = "氷は、ずっと冷たい場所にしか残らない"
MISSION_Q = "大事にしたい観点を3つ挙げよう"
MISSIONS = [
    ("氷採掘", "月で水（氷）を掘り出して、基地で使えるようにする。", ICE_HINT),
    ("電波天文台", "月に電波望遠鏡を建てて、宇宙から届く弱い電波を観測する。", None),
    ("太陽光発電", "月に太陽電池を並べて、基地で使う電気をつくる。", None),
    ("有人総合", "人が長く暮らせる基地をひとつ建てる。いろいろな条件を合わせて考える。", None),
]

# --- 地域カード（B4）の中立な1行。data/ の公開データで事実を確認した（_FACT_BASIS）。
#     ミッションとの相性・「第一候補」「理想」・半球や極の呼び名（南極・裏側・表側）は書かない。
REGION_LINES = {
    "赤道の海（静かの海）":
        "地質図で約86%が『海』（黒っぽい平らな地形）。傾斜は平均約1°。地球は空の高いところに見える（仰角 約48〜70°）。",
    "危機の海（東の孤立盆地）":
        "地質図で約73%が『海』。標高は約−3900〜+200 m。地球は空の低めから中ほどに見える（仰角 約20〜36°）。",
    "中緯度の火砕丘（Aristarchus 高原）":
        "『海』（約65%）と『陸』（約35%）が混じる。地球は空の中ほどに見える（仰角 約30〜44°）。",
    "溶岩チューブ天窓（Marius Hills）":
        "直径約58 m・深さ約40 mの縦孔が1つ記録されている。『陸』と『海』が混じり、地球は空の中ほどに見える（仰角 約30〜37°）。",
    "裏側・南極エイトケン（フォン・カルマン）":
        "約82%が『陸』で、標高は約−6900〜−1300 m。地球は地平線の下にあり、空に見えない。",
    "裏側・赤道（月の裏側の赤道帯）":
        "すべて『陸』で、標高の差が大きい（約−3400〜+5700 m）。地球は地平線の下にあり、空に見えない。",
    "南極（Shackleton-de Gerlache）":
        "太陽は正午でも地平線すれすれ（高さ 約0.5〜2.5°）。地球も地平線すれすれに見える（仰角 ±2.5°以内）。",
    "嵐の大洋（KREEP の広い海）":
        "月で最も大きい『海』（海の境界データで半径約1300 km）の一部。約78%が『海』。地球は空の中ほどから高めに見える（仰角 約18〜52°）。",
}
_FACT_BASIS = """
値はすべて data/site_environment.csv を、各地域の範囲（lat_min〜lat_max、lon_min〜lon_max）で切り出して集計：
  区分（USGS地質図の海/陸）の割合、earth_elev_deg の最小〜最大、noon_sun_elev_deg、elev_m（LOLA標高）、slope_deg の平均。
縦孔：data/lunar_pits.csv の Marius Hills Pit（lat 14.09, lon -56.77、opening_m 58、depth_m 40）が範囲内。
最大の海：data/maria_boundaries.csv の Oceanus Procellarum（radius_km 1296.12 が最大）。
火砕物・イルメナイト・実績地・資源の記述は、データで確認できないので書かない。
"""
FOOT_NOTE_REGION = "数値は同梱の公開データ（USGS地質図・LOLA標高・地球の仰角の計算）を、範囲内の1°セルで集めたもの。"

# 検査用の禁止語
FORBIDDEN_CARD = ["南極", "北極", "裏側", "表側", "半球", "第一候補", "理想", "実績", "永久影", "雑音", "静穏", "失格"]
FORBIDDEN_MISSION_WORDS = ["氷", "電波", "発電", "採掘", "天文", "通信", "太陽光", "有人", "向く", "最適"]
FORBIDDEN_PATTERN = re.compile(r"(氷採掘|電波天文台|太陽光発電|有人総合)(は|なら|には)")
ANSWER_PHRASES = ["第一候補", "電波天文に", "氷採掘は", "氷＝南極", "電波＝裏側", "見えたら失格", "標準の重み", "標準の解"]


# =====================================================================
# 地域データ（データ_地域 と candidate_regions.csv）
# =====================================================================
def load_regions_xlsx():
    import openpyxl
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=False, read_only=True)
    ws = wb["データ_地域"]
    rows = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or r[0] is None:
            continue
        rows.append((str(r[0]), float(r[1]), float(r[2]), float(r[3]), float(r[4])))
    wb.close()
    return rows


def load_regions_csv(path=None):
    with open(path or CSV_PATH, encoding="utf-8-sig") as f:
        return [(r["name"], float(r["lat_min"]), float(r["lat_max"]),
                 float(r["lon_min"]), float(r["lon_max"])) for r in csv.DictReader(f)]


def num(v):
    v = int(v) if float(v) == int(v) else v
    return str(v).replace("-", "−")


def rng(a, b):
    return f"{num(a)}°〜{num(b)}°"


def lon_str(lo0, lo1):
    """経度の表記。180°をこえる範囲は、日付変更線をまたぐ（181→西経179）として書く。"""
    if lo1 > 180:
        return f"東経{num(lo0)}°〜西経{num(360 - lo1)}°"
    return rng(lo0, lo1) + ("（全経度）" if (lo0 == -180 and lo1 == 180) else "")


# =====================================================================
# 描画の道具（単位mm、yは上から下へ）
# =====================================================================
def sw(s, size):
    return pdfmetrics.stringWidth(s, FONT, size) / mm


class Sheet:
    def __init__(self, c, W, H):
        self.c, self.W, self.H = c, W, H

    def Y(self, y):
        return (self.H - y) * mm

    def text(self, x, y, s, size=9, color=INK, bold=False, align="l"):
        c = self.c
        if align == "c":
            x -= sw(s, size) / 2
        elif align == "r":
            x -= sw(s, size)
        t = c.beginText(x * mm, self.Y(y))
        t.setFont(FONT, size)
        t.setFillColor(color)
        if bold:
            t.setTextRenderMode(2)
            c.setStrokeColor(color)
            c.setLineWidth(size * 0.018)
        else:
            t.setTextRenderMode(0)  # Tr は描画状態として残るので、毎回戻す
        t.textOut(s)
        t.setTextRenderMode(0)
        c.drawText(t)

    def rotated_text(self, x, y, s, size=9, color=INK):
        """縦書き風（反時計まわりに90°）。(x, y)は文字列の中央。"""
        c = self.c
        c.saveState()
        c.translate(x * mm, self.Y(y))
        c.rotate(90)
        t = c.beginText(-sw(s, size) / 2 * mm, 0)
        t.setFont(FONT, size)
        t.setFillColor(color)
        t.setTextRenderMode(0)
        t.textOut(s)
        c.drawText(t)
        c.restoreState()

    def para(self, x, y, w, s, size=9, leading=None, color=INK, align=0, raw=False):
        """折り返す文章。上端y。使った高さ(mm)を返す。"""
        style = ParagraphStyle("p", fontName=FONT, fontSize=size, leading=leading or size * 1.45,
                               textColor=color, wordWrap="CJK", alignment=align)
        p = Paragraph(s if raw else escape(s), style)
        _, h = p.wrap(w * mm, 10000 * mm)
        p.drawOn(self.c, x * mm, self.Y(y) - h)
        return h / mm

    def para_h(self, w, s, size=9, leading=None):
        style = ParagraphStyle("p", fontName=FONT, fontSize=size, leading=leading or size * 1.45, wordWrap="CJK")
        _, h = Paragraph(escape(s), style).wrap(w * mm, 10000 * mm)
        return h / mm

    def line(self, x1, y1, x2, y2, lw=0.3, color=LINE, dash=None):
        c = self.c
        c.setStrokeColor(color)
        c.setLineWidth(lw)
        c.setDash(*dash) if dash else c.setDash()
        c.line(x1 * mm, self.Y(y1), x2 * mm, self.Y(y2))
        c.setDash()

    def rect(self, x, y, w, h, lw=0.4, stroke=LINE, fill=None, r=0, dash=None):
        c = self.c
        c.setLineWidth(lw)
        if stroke is not None:
            c.setStrokeColor(stroke)
        if fill is not None:
            c.setFillColor(fill)
        c.setDash(*dash) if dash else c.setDash()
        if r:
            c.roundRect(x * mm, self.Y(y + h), w * mm, h * mm, r * mm,
                        stroke=1 if stroke is not None else 0, fill=1 if fill is not None else 0)
        else:
            c.rect(x * mm, self.Y(y + h), w * mm, h * mm,
                   stroke=1 if stroke is not None else 0, fill=1 if fill is not None else 0)
        c.setDash()

    def circle(self, x, y, r, stroke=LINE, fill=None, lw=0.4):
        c = self.c
        c.setLineWidth(lw)
        if stroke is not None:
            c.setStrokeColor(stroke)
        if fill is not None:
            c.setFillColor(fill)
        c.circle(x * mm, self.Y(y), r * mm, stroke=1 if stroke is not None else 0, fill=1 if fill is not None else 0)

    def bar(self, x, y, w, h, s, size=10, fill=ACCENT, color=WHITE):
        self.rect(x, y, w, h, stroke=None, fill=fill)
        self.text(x + 2.5, y + h / 2 + size * 0.35 * 0.3528, s, size, color, bold=True)

    def ruled(self, x, y, w, n, pitch=6.8, color=RULE):
        """手書きの罫線n本（上端yから、pitchごと）。下端のyを返す。"""
        for i in range(n):
            self.line(x, y + pitch * (i + 1), x + w, y + pitch * (i + 1), 0.3, color)
        return y + pitch * n

    def blank(self, x, y, w, label, size=8.5, unit="", lw=0.35):
        """「ラベル ________ 単位」。yは文字のベースライン。"""
        self.text(x, y, label, size)
        lx = x + sw(label, size) + 1.2
        ux = x + w - (sw(unit, size) + 0.8 if unit else 0)
        self.line(lx, y + 0.8, ux, y + 0.8, lw, LINE)
        if unit:
            self.text(ux + 0.6, y, unit, size)

    def judge(self, x, y, kind, w=34):
        """教員の判定用チェック欄（知・思・態）。左上x,y。"""
        self.rect(x, y, w, 5.2, lw=0.35, stroke=LINE, fill=WHITE, r=0.8, dash=(1.2, 0.8))
        self.text(x + 1.6, y + 3.7, "教員判定", 6.2, SUB)
        self.rect(x + 12.3, y + 0.9, 3.4, 3.4, lw=0.35, stroke=ACCENT)
        self.text(x + 14.0, y + 3.5, kind, 5.6, ACCENT, align="c")
        self.text(x + 19.5, y + 3.8, "○", 7.5, INK)
        self.text(x + 25.5, y + 3.8, "△", 7.5, INK)
        self.text(x + 30.5, y + 3.7, "", 6)

    def grid(self, x, y, colw, rowh, cells, size=8.5, head_rows=1, head_fill=HEAD_BG, pad=1.6,
             head_size=None, valign="m", fills=None, color=INK, heads_left=0):
        """表。cells[行][列]（文字列またはNone）。下端のyを返す。"""
        if not isinstance(rowh, (list, tuple)):
            rowh = [rowh] * len(cells)
        yy = y
        for ri, row in enumerate(cells):
            xx = x
            for ci, cell in enumerate(row):
                w, h = colw[ci], rowh[ri]
                is_head = ri < head_rows or ci < heads_left
                fill = None
                if is_head:
                    fill = head_fill
                if fills and fills.get((ri, ci)) is not None:
                    fill = fills[(ri, ci)]
                self.rect(xx, yy, w, h, lw=0.4, stroke=LINE, fill=fill)
                if cell:
                    sz = (head_size or size) if is_head else size
                    ph = self.para_h(w - 2 * pad, cell, sz)
                    ty = yy + (h - ph) / 2 if valign == "m" else yy + pad
                    self.para(xx + pad, ty, w - 2 * pad, cell, sz, color=ACCENT if is_head else color)
                xx += w
            yy += rowh[ri]
        return yy


def new_canvas(name, size):
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, name)
    c = canvas.Canvas(path, pagesize=(size[0] * mm, size[1] * mm))
    c.setTitle(os.path.splitext(name)[0])
    c.setAuthor("")
    c.setCreator("course/build_materials.py")
    return c, Sheet(c, size[0], size[1]), path


# =====================================================================
# 月の全球図（正距円筒図）
# =====================================================================
_FADE_CACHE = {}


def faded_moon(alpha):
    """月画像を白と混ぜて薄くし、JPEGで一時保存して返す（図の上に書き込めるように）。"""
    if alpha in _FADE_CACHE:
        return _FADE_CACHE[alpha]
    from PIL import Image
    im = Image.open(MOON_IMG).convert("RGB")
    white = Image.new("RGB", im.size, (255, 255, 255))
    im = Image.blend(im, white, alpha)
    d = tempfile.mkdtemp(prefix="moonmap_")
    p = os.path.join(d, f"moon_{int(alpha * 100)}.jpg")
    im.save(p, quality=88)
    _FADE_CACHE[alpha] = p
    return p


class MoonMap:
    """経度-180〜180・緯度-90〜90の正距円筒図。左上(x,y)、幅w、高さw/2。"""

    def __init__(self, s, x, y, w, alpha=0.5):
        self.s, self.x, self.y, self.w, self.h = s, x, y, w, w / 2
        s.c.drawImage(faded_moon(alpha), x * mm, s.Y(y + self.h), w * mm, self.h * mm)
        s.rect(x, y, w, self.h, lw=0.6, stroke=INK)

    def px(self, lon):
        return self.x + (lon + 180) / 360 * self.w

    def py(self, lat):
        return self.y + (90 - lat) / 180 * self.h

    def grid(self, step=15, color=colors.HexColor("#6b7785"), lw=0.2):
        s = self.s
        for lon in range(-180, 181, step):
            if lon not in (-180, 180):
                s.line(self.px(lon), self.y, self.px(lon), self.y + self.h, 0.45 if lon == 0 else lw, color)
        for lat in range(-90, 91, step):
            if lat not in (-90, 90):
                s.line(self.x, self.py(lat), self.x + self.w, self.py(lat), 0.45 if lat == 0 else lw, color)

    def ticks(self, lab_size=6.5):
        """10°ごとの目盛と数字（四辺）。"""
        s = self.s
        for lon in range(-180, 181, 10):
            xx = self.px(lon)
            L = 2.0 if lon % 30 == 0 else 1.2
            s.line(xx, self.y - L, xx, self.y, 0.4, INK)
            s.line(xx, self.y + self.h, xx, self.y + self.h + L, 0.4, INK)
            lab = num(lon)
            s.text(xx, self.y + self.h + L + 3.1, lab, lab_size, INK, align="c")
            s.text(xx, self.y - L - 0.9, lab, lab_size, INK, align="c")
        for lat in range(-90, 91, 10):
            yy = self.py(lat)
            L = 2.0 if lat % 30 == 0 else 1.2
            s.line(self.x - L, yy, self.x, yy, 0.4, INK)
            s.line(self.x + self.w, yy, self.x + self.w + L, yy, 0.4, INK)
            lab = num(lat)
            s.text(self.x - L - 0.8, yy + 1.0, lab, lab_size, INK, align="r")
            s.text(self.x + self.w + L + 0.8, yy + 1.0, lab, lab_size, INK, align="l")

    def box(self, lat_min, lat_max, lon_min, lon_max, color=MARK, lw=0.8, fill=None):
        """範囲の枠。経度が180をこえる分は反対側へ折り返す。"""
        parts = [(lon_min, lon_max)]
        if lon_max > 180:
            parts = [(lon_min, 180), (-180, lon_max - 360)]
        for a, b in parts:
            self.s.rect(self.px(a), self.py(lat_max), self.px(b) - self.px(a), self.py(lat_min) - self.py(lat_max),
                        lw=lw, stroke=color, fill=fill)

    def label(self, lon, lat, n, dlon, dlat, r=2.4, size=7.5, anchor=None):
        """番号の丸。枠の中心から(dlon, dlat)ずらした位置に置き、細い線でつなぐ。"""
        s = self.s
        cx, cy = self.px(lon), self.py(lat)
        lx, ly = self.px(lon + dlon), self.py(lat + dlat)
        if anchor:
            cx, cy = self.px(anchor[0]), self.py(anchor[1])
        if dlon or dlat:
            s.line(cx, cy, lx, ly, 0.35, MARK)
        s.circle(lx, ly, r, stroke=MARK, fill=WHITE, lw=0.6)
        s.text(lx, ly + size * 0.35 * 0.3528 * 1.0, str(n), size, MARK, bold=True, align="c")


# =====================================================================
# B1 個人用ワークシート（B5・表裏）
# =====================================================================
def page_header(s, tag, with_mission=True):
    W = s.W
    s.text(10, 12.5, "月データでムーンベースの場所を決めよう", 12.5, ACCENT, bold=True)
    s.text(W - 10, 12.5, "個人用ワークシート　" + tag, 8.5, SUB, align="r")
    s.line(10, 15, W - 10, 15, 0.8, ACCENT)
    y = 17
    s.rect(10, y, W - 20, 15, lw=0.5, stroke=LINE)
    s.line(10, y + 7.5, W - 10, y + 7.5, 0.3, LINE)
    s.blank(12, y + 5.4, 100, "氏名", 8.5)
    s.blank(118, y + 5.4, 44, "班", 8.5)
    s.blank(12, y + 12.9, 74, "ミッション", 8.5)
    s.text(114, y + 12.9, "出席", 7.5, SUB)
    for i, lab in enumerate(["第1時", "第2時", "第3時"]):
        bx = 128 + i * 12.5
        s.rect(bx, y + 9.3, 3.4, 3.4, lw=0.35, stroke=LINE)
        s.text(bx + 4.4, y + 12.4, lab, 7, INK)
    return y + 15 + 2.5


def cbox(s, x, y, label, size=8.0):
    """□ラベル。x,yは文字のベースライン。右端のxを返す。"""
    s.rect(x, y - 2.7, 3.2, 3.2, lw=0.35, stroke=LINE)
    s.text(x + 4.3, y, label, size)
    return x + 4.3 + sw(label, size)


def b1_memo(s, x0, w, y):
    """⑥ 第2時の判断メモ（評価に入れない。画面の数字を写し、理由を一言書く。教員の確認印欄）。上端yから。下端のyを返す。"""
    title = "⑥ " + L_MEMO
    s.text(x0, y + 1.6, title, 8.5, ACCENT, bold=True)
    s.text(x0 + sw(title, 8.5) + 2.0, y + 1.6, "評価には入れない", 6.8, SUB)
    # 教員の確認印（小さな枠）
    s.rect(x0 + w - 24, y - 2.4, 24, 7.6, lw=0.35, stroke=LINE, fill=WHITE, r=0.8, dash=(1.2, 0.8))
    s.text(x0 + w - 22.6, y + 0.4, "教員確認印", 5.8, SUB)
    y += 5.0
    # (a) 判断①：線（ステップ1b）。数字は画面の値を写す
    s.blank(x0 + 1, y + 3.6, 33, "(a) 線", 8.5, "°")
    s.blank(x0 + 38, y + 3.6, 47, "内側の最大", 8.5, "％")
    s.blank(x0 + 89, y + 3.6, 66, "使えなくなる月の面積", 8.5, "％")
    y += 6.2
    s.text(x0 + 3, y + 2.4, "ステップ1b　内側の最大＝線の内側で、最高が21〜3時に来る割合の最大（B9）／面積（B10）", 6.5, SUB)
    y += 3.8
    # (a) の理由（一言）
    s.blank(x0 + 6, y + 3.6, 149, L_LINE_REASON, 8.0)
    y += 6.7
    # (b) 判断②：担当クレーター（ステップ1c）
    s.blank(x0 + 1, y + 3.6, 84, "(b) 担当クレーター", 8.5)
    s.blank(x0 + 89, y + 3.6, 66, "内部の夜の最低温度", 8.5, "K")
    y += 6.9
    lab = "帯の9割が入る範囲"
    s.text(x0 + 6, y + 3.6, lab, 8.5)
    lx = x0 + 6 + sw(lab, 8.5) + 1.2
    s.line(lx, y + 4.4, lx + 14, y + 4.4, 0.35, LINE)
    s.text(lx + 15.2, y + 3.6, "〜", 8.5)
    s.line(lx + 20.5, y + 4.4, lx + 34.5, y + 4.4, 0.35, LINE)
    s.text(lx + 36, y + 3.6, "K", 8.5)
    s.text(x0 + 76, y + 3.6, "内部は範囲の", 8.0)
    k = cbox(s, x0 + 76 + sw("内部は範囲の", 8.0) + 1.5, y + 3.6, "外")
    k = cbox(s, k + 2.0, y + 3.6, "内")
    s.text(k + 3.0, y + 3.6, "→　帯と", 8.0)
    k = cbox(s, k + 3.0 + sw("→　帯と", 8.0) + 1.5, y + 3.6, "違う")
    cbox(s, k + 2.0, y + 3.6, "違わない")
    y += 6.7
    # (b) の理由（一言）
    s.blank(x0 + 6, y + 3.6, 149, L_IO_REASON, 8.0)
    return y + 7.0


def b1_page1(s):
    W = s.W
    x0, w = 10, s.W - 20
    y = page_header(s, "（表：第1時・第2時）")

    # ---------------- 第1時
    s.bar(x0, y, w, 6, "第1時　月を見て、自分のミッションに向く地域タイプを予想する", 9.5)
    y += 8.5
    # 温度差の予想（1行）
    s.text(x0, y + 1.6, "① " + L_PREDICT_TEMP + "（1行）", 8.5, ACCENT, bold=True)
    y += 3.2
    s.text(x0 + 1, y + 6.3, "月の昼と夜の温度差は、およそ", 8.5)
    s.line(x0 + 44, y + 7.1, x0 + 70, y + 7.1, 0.35, LINE)
    s.text(x0 + 71, y + 6.3, "℃ くらい。なぜなら", 8.5)
    s.line(x0 + 103, y + 7.1, x0 + w, y + 7.1, 0.35, LINE)
    y = s.ruled(x0, y + 6.8, w, 2, 6.6) + 0.2
    y += 1.2
    # 予想
    s.text(x0, y + 1.6, "② " + L_PREDICT, 8.5, ACCENT, bold=True)
    s.text(x0 + sw("② " + L_PREDICT, 8.5) + 2.5, y + 1.6, "個人で先に書く（班で見せ合う前に）", 7, SUB)
    y += 4.5
    s.blank(x0 + 1, y + 3.8, 76, "地域タイプ（地域カードの番号・名前）", 8)
    s.blank(x0 + 88, y + 3.8, 30, "緯度", 8, "°")
    s.blank(x0 + 122, y + 3.8, 32, "経度", 8, "°")
    y += 5.4
    s.text(x0 + 1, y + 3.4, "理由（どの指標を重く見たか）", 8, SUB)
    y = s.ruled(x0, y + 3.2, w, 5, 6.5) + 0.6
    y += 1.6

    # ---------------- 第2時
    s.bar(x0, y, w, 6, "第2時　温度と海陸のデータで、平均と群を比べる", 9.5)
    y += 8.0
    s.text(x0, y + 1.6, "③ " + L_NICHI, 8.5, ACCENT, bold=True)
    s.judge(x0 + w - 34, y - 2.2, "知")
    y += 4.2
    s.text(x0 + 1, y + 3.8, "表計算ステップ1のAの表（緯度の絶対値の4つの帯）から、日較差の平均を読む", 7.2, SUB)
    y += 5.0
    s.blank(x0 + 1, y + 3.8, 36, "赤道帯", 8.5, "K")
    s.blank(x0 + 40, y + 3.8, 36, "中緯度帯", 8.5, "K")
    s.blank(x0 + 79, y + 3.8, 36, "高緯度帯", 8.5, "K")
    s.blank(x0 + 118, y + 3.8, 36, "極付近", 8.5, "K")
    y += 6.6
    s.blank(x0 + 1, y + 3.8, 100, "日較差が大きいのは（赤道帯・極付近）", 8.5)
    y += 7.4
    s.rect(x0, y, w, 6.0, lw=0.3, stroke=RULE, fill=SOFT_BG, r=0.8)
    s.text(x0 + 2, y + 4.1, K_CELSIUS, 7.3, SUB)
    y += 9.6
    s.text(x0, y + 1.6, "④ " + L_DENSITY, 8.5, ACCENT, bold=True)
    s.judge(x0 + w - 34, y - 2.2, "知")
    y += 4.2
    s.blank(x0 + 1, y + 3.8, 34, "海の密度", 8.5)
    s.blank(x0 + 40, y + 3.8, 34, "陸の密度", 8.5)
    s.blank(x0 + 80, y + 3.8, 44, "陸は海の", 8.5, "倍")
    y += 6.6
    s.blank(x0 + 1, y + 3.8, 104, "USGS地質図の年代（数が大きいほど新しい）　海", 8)
    s.blank(x0 + 110, y + 3.8, 40, "陸", 8.5)
    y += 9.0
    s.text(x0, y + 1.6, "⑤ " + L_SEARIKU, 8.5, ACCENT, bold=True)
    s.text(x0 + sw("⑤ " + L_SEARIKU, 8.5) + 2.5, y + 1.6, "数値（密度・年代）を入れて書く", 7, SUB)
    s.judge(x0 + w - 34, y - 2.2, "思")
    y += 3.2
    y = s.ruled(x0, y + 0.5, w, 3, 6.8) + 0.5
    y += 1.6
    y = b1_memo(s, x0, w, y)
    s.text(W / 2, 247.6, "B1 個人用ワークシート（表）　裏面は第3時", 6.5, SUB, align="c")
    return y


def b1_page2(s):
    W = s.W
    x0, w = 10, s.W - 20
    y = page_header(s, "（裏：第3時）")

    s.bar(x0, y, w, 6, "第3時　重みを決めて1地点を選び、質問を受けながら共有する", 9.5)
    y += 8.5
    s.text(x0, y + 1.6, "⑦ " + L_DECIDE, 8.5, ACCENT, bold=True)
    s.text(x0 + sw("⑦ " + L_DECIDE, 8.5) + 2.5, y + 1.6, "第1時の予想を見て、変えてもよい", 7, SUB)
    y += 4.2
    s.blank(x0 + 1, y + 3.8, 98, "地域タイプ（地域名）", 8.5)
    s.text(x0 + 104, y + 3.8, "予想から", 8, SUB)
    s.rect(x0 + 118, y + 0.9, 3.2, 3.2, lw=0.35, stroke=LINE)
    s.text(x0 + 122.6, y + 3.8, "変えた", 8)
    s.rect(x0 + 136, y + 0.9, 3.2, 3.2, lw=0.35, stroke=LINE)
    s.text(x0 + 140.6, y + 3.8, "変えない", 8)
    y += 6.2
    s.blank(x0 + 1, y + 3.8, 50, "決めた1地点　緯度", 8.5, "°")
    s.blank(x0 + 56, y + 3.8, 36, "経度", 8.5, "°")
    s.text(x0 + 96, y + 3.8, "（付箋にも書く）", 7.5, SUB)
    y += 8.2

    s.text(x0, y + 1.6, "⑧ " + L_REASON, 8.5, ACCENT, bold=True)
    s.judge(x0 + w - 34, y - 2.2, "思")
    y += 3.6
    s.rect(x0, y, w, 6.0, lw=0.3, stroke=RULE, fill=SOFT_BG, r=0.8)
    s.text(x0 + 2, y + 4.1, "型：" + REASON_TYPE, 7.0, SUB)
    y += 6.0
    y = s.ruled(x0, y, w, 4, 6.5) + 0.5
    y += 2.0
    s.text(x0, y + 1.6, "⑨ " + L_WEIGHTCHG, 8.5, ACCENT, bold=True)
    s.judge(x0 + w - 34, y - 2.2, "思")
    y += 3.0
    s.text(x0, y + 3.0, "重みを変えたとき、上位10の経度帯（南極付近の班は、日照率・永久影までの距離）はどう変わった？",
           7.0, SUB)
    y = s.ruled(x0, y + 3.2, w, 2, 6.5) + 0.5
    y += 2.2

    s.text(x0, y + 1.6, "⑩ " + L_QA, 8.5, ACCENT, bold=True)
    s.text(x0 + sw("⑩ " + L_QA, 8.5) + 2.5, y + 1.6, "質問の型：" + "　".join("「" + q + "」" for q in QUESTION_TYPES),
           7, SUB)
    s.judge(x0 + w - 34, y - 2.2, "知")
    y += 3.6
    s.text(x0 + 1, y + 3.4, "相手班", 7.5, SUB)
    s.line(x0 + 11, y + 3.9, x0 + 24, y + 3.9, 0.35, LINE)
    s.text(x0 + 27, y + 3.4, "わたしが出した質問", 7.5, SUB)
    y = s.ruled(x0, y + 3.2, w, 2, 6.5)
    s.text(x0 + 1, y + 4.0, "相手の答え（数値か指標名を入れて）", 7.5, SUB)
    y = s.ruled(x0, y + 3.8, w, 2, 6.5) + 0.5
    y += 2.0

    s.text(x0, y + 1.6, "⑪ " + L_REFLECT, 8.5, ACCENT, bold=True)
    s.text(x0 + sw("⑪ " + L_REFLECT, 8.5) + 2.5, y + 1.6, "何が変わった？（変わらなかった？）なぜ？", 7, SUB)
    s.judge(x0 + w - 34, y - 2.2, "態")
    y += 3.2
    y = s.ruled(x0, y + 0.5, w, 3, 6.6) + 0.5
    y += 2.2

    s.text(x0, y + 1.6, "⑫ " + L_DOUBT, 8.5, ACCENT, bold=True)
    s.text(x0 + sw("⑫ " + L_DOUBT, 8.5) + 2.5, y + 1.6, "⑪を書き終えてから。黒板に出た例から3つ選び、各1文", 7, SUB)
    y += 3.5
    for i in range(3):
        s.text(x0 + 1, y + 5.4, f"{i + 1}.", 8)
        s.line(x0 + 7, y + 6.2, x0 + w, y + 6.2, 0.3, RULE)
        y += 7.4
    # 教員の総括欄（紙の端）
    s.rect(x0, 237.2, w, 7.2, lw=0.4, stroke=LINE, fill=SOFT_BG, r=0.8)
    s.text(x0 + 2, 241.7, "教員（総括）", 7, SUB)
    for i, (lab, kx) in enumerate([("知識・技能", 26), ("思考・判断・表現", 62), ("主体的に学習に取り組む態度", 108)]):
        s.text(x0 + kx, 241.7, lab, 7, ACCENT)
        s.text(x0 + kx + sw(lab, 7) + 2, 241.9, "○　△", 8, INK)
    s.text(W / 2, 247.6, "B1 個人用ワークシート（裏）　回収は第3時の終わり", 6.5, SUB, align="c")
    return y


def build_b1():
    c, s, path = new_canvas("B1_個人用ワークシート.pdf", B5)
    b1_page1(s)
    c.showPage()
    b1_page2(s)
    c.showPage()
    c.save()
    return path


# =====================================================================
# B2 班用ワークシート（A3・表裏・横）
# =====================================================================
def b2_header(s, tag, with_region=False):
    W = s.W
    s.text(10, 14, "班用ワークシート", 16, ACCENT, bold=True)
    s.text(10 + sw("班用ワークシート", 16) + 4, 14, "作業用（記録はしない。評価は個人用ワークシートに書く）　" + tag, 9, SUB)
    s.line(10, 17, W - 10, 17, 1.0, ACCENT)
    s.blank(10, 25.5, 30, "班", 10)
    s.blank(46, 25.5, 66, "ミッション", 10)
    if with_region:
        s.blank(118, 25.5, 120, "地域タイプ（地域カードの番号・地域名）", 10)
    return 30


def cw(total, ratios):
    t = float(sum(ratios))
    return [total * r / t for r in ratios]


def b2_front(s):
    W = s.W
    y0 = b2_header(s, "（表：第1時・第2時）")
    L, mid, R = 10, 214, W - 10
    lw_ = mid - L - 8
    rw_ = R - mid

    # ===== 左：第1時
    y = y0
    s.bar(L, y, lw_, 7.5, "第1時　Webで5地点を読み、地域タイプを選ぶ", 12)
    y += 10.5
    s.text(L, y + 2.5, "役割（コマごとに交代。名前または○）", 9, ACCENT, bold=True)
    y += 4.2
    cells = [["", "第1時", "第2時", "第3時"]]
    for r in ["操作", "記録", "発表", "検算"]:
        cells.append([r, "", "", ""])
    y = s.grid(L, y, cw(lw_, [20, 30, 30, 30]), [6.2] + [6.4] * 4, cells, size=9, heads_left=1) + 3.5

    s.text(L, y + 2.5, "ジグソーの5地点（A〜E。3°グリッドの中心）　教員が、班番号を書き込む", 9, ACCENT, bold=True)
    y += 4.2
    cells = [["", "地点", "緯度", "経度", "読む班（教員記入）"]]
    for n, nm, la, lo in JIGSAW_SITES:
        cells.append([n, nm, num(la) + "°", num(lo) + "°", "班　　　・班"])
    y = s.grid(L, y, cw(lw_, [8, 40, 16, 16, 34]), [6.2] + [6.8] * 5, cells, size=9) + 1.2
    h = s.para(L, y + 0.5, lw_, WEB_TEXT["find_site_b2"], 7.6, color=SUB)
    y += h + 3.0

    s.text(L, y + 2.5, "記録表（黒板の表にも書く）　担当地点の記号", 9, ACCENT, bold=True)
    s.rect(L + 76, y - 1.0, 13, 5.4, lw=0.4, stroke=LINE)
    s.text(L + 93, y + 2.5, "地点名", 9, SUB)
    s.line(L + 106, y + 3.2, L + 175, y + 3.2, 0.35, LINE)
    y += 4.5
    cells = [["指標（Webのピンカードの表の順）", "値", "単位", "出典（ピンカードの表の行の名前）", "相手班の値"],
             ["1日の温度差（日較差）", "", "℃", "", ""],
             ["夜の最低温度", "", "℃", "", ""],
             ["正午の太陽高度", "", "°", "", ""],
             ["地球の仰角", "", "°", "", ""]]
    y = s.grid(L, y, cw(lw_, [40, 18, 12, 36, 22]), [8.0] + [9.4] * 4, cells, size=9, heads_left=1) + 1.0
    s.text(L, y + 3.2, "地球の仰角：＋は地球が空に見える、−は地平線の下。ジグソーでは、「1日の温度」層の24時間の差ではなく、表の「1日の温度差」を読む。", 7.4, SUB)
    s.text(L, y + 7.6, "相手班の値がちがったら、入力した緯度・経度と、表の下の「セルの中心」を見比べる。", 7.4, SUB)
    y += 11.5

    s.text(L, y + 2.5, "黒板の表を見て考える：自分のミッションにとって、値は高いほどよい？　低いほどよい？", 9, ACCENT, bold=True)
    y += 4.2
    cells = [["指標", "高いほどよい", "低いほどよい", "どちらでもない", "理由・気づき"]]
    for nm in ["1日の温度差（日較差）", "夜の最低温度", "正午の太陽高度", "地球の仰角"]:
        cells.append([nm, "", "", "", ""])
    y = s.grid(L, y, cw(lw_, [28, 16, 16, 17, 42]), [6.2] + [8.0] * 4, cells, size=8.8, heads_left=1) + 3.5

    s.text(L, y + 2.5, "班の地域タイプと予想（個人の予想を見比べて決める）", 9, ACCENT, bold=True)
    y += 4.2
    s.blank(L + 1, y + 5.5, 80, "地域カードの番号・名前", 9)
    s.blank(L + 86, y + 5.5, 50, "予想　緯度", 9, "°")
    s.blank(L + 142, y + 5.5, 50, "経度", 9, "°")
    s.text(L + 1, y + 13.0, "理由", 9, SUB)
    s.ruled(L, y + 8.0, lw_, 3, 7.4)

    # ===== 右：第2時
    y = y0
    s.bar(mid, y, rw_, 7.5, "第2時　温度の線・クレーターの確かめ・海と陸・重みの案", 12)
    y += 10.5
    s.text(mid, y + 2.5, "ステップ1　緯度帯ごとの日較差の平均", 9, ACCENT, bold=True)
    s.text(mid + 64, y + 2.5, "（数値は表計算を見て、個人用ワークシートに書く）", 7.8, SUB)
    y += 5.0
    s.text(mid + 1, y + 2.8, "予想：赤道と極付近で、日較差が大きいのは", 9)
    s.line(mid + 68, y + 3.6, mid + 100, y + 3.6, 0.35, LINE)
    y += 8.0

    # --- 判断①：帯を刻んで線を引く（ステップ1b）
    s.text(mid, y + 2.5, "ステップ1b　帯を刻んで、使う緯度の上限（線）を決める", 9, ACCENT, bold=True)
    s.text(mid + 94, y + 2.5, "（個人の線と数字は、個人用ワークシート⑥(a)に書く）", 7.8, SUB)
    y += 5.6
    s.blank(mid + 1, y + 3.4, 112, "班の線（黄色いセル B7 に入れた値）　緯度の絶対値", 9, "°")
    y += 7.6
    s.text(mid + 1, y + 2.5, "刻むと見えたこと（帯の幅を30→5→1°と変えて、グラフはどう変わった？）", 8.4, SUB)
    y = s.ruled(mid, y + 3.0, rw_, 2, 8.0) + 5.0

    # --- 判断②：クレーターと帯（ステップ1c）
    s.text(mid, y + 2.5, "ステップ1c　担当クレーターの確かめ", 9, ACCENT, bold=True)
    s.text(mid + 70, y + 2.5, "（数字は、個人用ワークシート⑥(b)に書く）", 7.8, SUB)
    y += 5.6
    s.blank(mid + 1, y + 3.4, 120, "担当クレーター（クレーターカードの名前）", 9)
    y += 7.6
    s.text(mid + 1, y + 2.5, "班で話す：24時間カーブの2本は、朝・昼・夕・夜の、どの時刻で近く、どの時刻で離れている？", 8.4, SUB)
    y = s.ruled(mid, y + 3.0, rw_, 2, 8.0) + 5.0

    s.text(mid, y + 2.5, "ステップ2　海と陸のクレーター", 9, ACCENT, bold=True)
    s.text(mid + 52, y + 2.5, "（数値は表計算を見て、個人用ワークシートに書く）", 7.8, SUB)
    y += 5.0
    s.text(mid + 1, y + 2.8, "予想：クレーターが少ないのは（海・陸）", 9)
    y += 6.5
    s.text(mid, y + 2.5, "班で話す：なぜ少ない？（落ちなかった？　落ちた後に消えた？）　地質図の年代と合っていた？", 8.4, SUB)
    y = s.ruled(mid, y + 3.0, rw_, 2, 8.0) + 5.0

    s.text(mid, y + 2.5, "重みの案（まず極端に：3と1など）。「大きいほど良い／小さいほど良い」は表計算に表示される。重みは0〜5の整数", 9, ACCENT, bold=True)
    y += 4.5
    s.text(mid, y + 2.5, "「根拠にした今日の数字」：今日、表計算で見た数字（個人用の③④⑥など）から、行ごとに1つ書く。", 8.2, SUB)
    y += 4.0
    wcols = cw(rw_, [34, 9, 9, 48, 54])
    cells = [["指標（ステップ4）", "○×", "重み", "理由", "根拠にした今日の数字"],
             ["太陽高度", "", "", "", ""],
             ["1日の温度差（日較差）", "", "", "", ""],
             ["地球の仰角（表側）", "", "", "", ""],
             ["地球の仰角（裏側）", "", "", "", ""],
             ["夜の最低温度", "", "", "", ""]]
    y = s.grid(mid, y, wcols, [6.5] + [7.4] * 5, cells, size=9, heads_left=1) + 2.4
    s.text(mid, y + 3.0, "南極付近の班は、ステップ4bの4行で考える", 8.2, ACCENT, bold=True)
    y += 4.4
    cells = [["指標（ステップ4b）", "○×", "重み", "理由", "根拠にした今日の数字"],
             ["日照率", "", "", "", ""],
             ["永久影までの距離", "", "", "", ""],
             ["永久影率", "", "", "", ""],
             ["傾斜", "", "", "", ""]]
    y = s.grid(mid, y, wcols, [6.5] + [7.4] * 4, cells, size=9, heads_left=1)
    s.text(W / 2, 291, "B2 班用ワークシート（表）　裏面は第3時", 7, SUB, align="c")


def b2_back(s):
    W = s.W
    y0 = b2_header(s, "（裏：第3時）", with_region=True)
    L, R = 10, W - 10
    gap = 9
    totw = R - L - 2 * gap
    w1, w2, w3 = totw * 0.31, totw * 0.31, totw * 0.38
    c1 = L
    c2 = c1 + w1 + gap
    c3 = c2 + w2 + gap

    s.bar(L, y0, R - L, 7.5, "第3時　重みを試して1地点を決め、Webで検算し、質問を受ける", 12)
    y0 += 11

    # ---- 列1：重み（試行1・試行2）
    y = y0
    s.text(c1, y + 2.5, "① 重み（0〜5の整数）　試行1と試行2", 9, ACCENT, bold=True)
    s.text(c1, y + 7.0, "重みを変えると前の結果は消える。変える前に、②に書き写す。", 7.8, SUB)
    y += 9.0
    cells = [["指標（ステップ4）", "試行1", "試行2"],
             ["太陽高度", "", ""],
             ["1日の温度差（日較差）", "", ""],
             ["地球の仰角（表側）", "", ""],
             ["地球の仰角（裏側）", "", ""],
             ["夜の最低温度", "", ""]]
    y = s.grid(c1, y, cw(w1, [50, 25, 25]), [7] + [12.0] * 5, cells, size=9.5, heads_left=1) + 6.0
    s.text(c1, y + 2.5, "南極付近の班は、次の4行（ステップ4b）", 9, ACCENT, bold=True)
    y += 4.5
    cells = [["指標（ステップ4b）", "試行1", "試行2"],
             ["日照率", "", ""],
             ["永久影までの距離", "", ""],
             ["永久影率", "", ""],
             ["傾斜", "", ""]]
    y = s.grid(c1, y, cw(w1, [50, 25, 25]), [7] + [12.0] * 4, cells, size=9.5, heads_left=1) + 7.0
    s.rect(c1, y, w1, 17, lw=0.4, stroke=RULE, fill=SOFT_BG, r=1)
    s.para(c1 + 3, y + 3, w1 - 6,
           "このシートに書かないこと：試行1と2のちがい・決めた1地点・地域名は、個人用ワークシート（⑦⑨）に書く。1地点は付箋にも書く。",
           8.2, color=SUB)

    # ---- 列2：上位10の範囲・境界の理由
    y = y0
    s.text(c2, y + 2.5, "② 上位10の範囲（表計算に表示される）", 9, ACCENT, bold=True)
    y += 4.5
    cells = [["低〜中緯度の班（ステップ4）", "試行1", "試行2"],
             ["緯度の範囲", "　　　°〜　　　°", "　　　°〜　　　°"],
             ["経度の範囲", "　　　°〜　　　°", "　　　°〜　　　°"]]
    y = s.grid(c2, y, cw(w2, [44, 30, 30]), [7, 12, 12], cells, size=9, heads_left=1) + 5.0
    cells = [["南極付近の班（ステップ4b）", "試行1", "試行2"],
             ["日照率の範囲 [%]", "　　　〜　　　", "　　　〜　　　"],
             ["傾斜の範囲 [°]", "　　　〜　　　", "　　　〜　　　"],
             ["永久影までの距離 [km]", "　　　〜　　　", "　　　〜　　　"]]
    y = s.grid(c2, y, cw(w2, [44, 30, 30]), [7, 12, 12, 12], cells, size=9, heads_left=1) + 8.0
    s.text(c2, y + 2.5, "③ 上位10が、領域の境界の近くに集まる理由は？", 9, ACCENT, bold=True)
    s.text(c2, y + 7.5, "（上位10の位置を見て、班で話す）", 7.8, SUB)
    y = s.ruled(c2, y + 8.0, w2, 5, 8.0)

    # ---- 列3：検算・質疑
    y = y0
    s.text(c3, y + 2.5, "④ 検算欄（Webアプリで。地域タイプで欄が違う）", 9, ACCENT, bold=True)
    y += 4.5
    hb = 76
    s.rect(c3, y, w3, hb, lw=0.5, stroke=LINE)
    s.text(c3 + 2, y + 5.2, "低〜中緯度：傾斜・地球の仰角・日較差を、ピンカードの表から読む", 9, ACCENT)
    yy = y + 12
    s.blank(c3 + 2, yy, 40, "傾斜", 9, "°")
    s.blank(c3 + 46, yy, 46, "地球の仰角", 9, "°")
    s.blank(c3 + 96, yy, w3 - 98, "日較差", 9, "")
    s.text(c3 + 2, yy + 8, "スコアに入っていない値で、気づいたこと（落とし穴はないか）", 8.2, SUB)
    s.ruled(c3 + 2, yy + 9, w3 - 4, 5, 8.0)
    y += hb + 4.0
    hb2 = 84
    s.rect(c3, y, w3, hb2, lw=0.5, stroke=LINE)
    s.text(c3 + 2, y + 5.2, "南極付近：傾斜・永久影までの距離は表計算、仰角はWebで読む", 9, ACCENT)
    yy = y + 12
    s.blank(c3 + 2, yy, 46, "傾斜（表計算）", 9, "°")
    s.blank(c3 + 52, yy, w3 - 54, "永久影までの距離（表計算）", 9, "km")
    yy += 8
    s.blank(c3 + 2, yy, 56, "地球の仰角（Web）", 9, "°")
    yy += 8
    s.text(c3 + 2, yy, "Webと表計算でずれの大きかった値", 9)
    s.line(c3 + 62, yy + 0.8, c3 + w3 - 2, yy + 0.8, 0.35, LINE)
    s.text(c3 + 2, yy + 8, "気づいたこと", 8.2, SUB)
    s.ruled(c3 + 2, yy + 9, w3 - 4, 4, 8.0)
    y += hb2 + 5.0
    s.text(c3, y + 2.5, "⑤ 質疑メモ（受けた質問と、わたしたちの答え）", 9, ACCENT, bold=True)
    s.text(c3, y + 8.0, "質問の型：" + "　".join("「" + q + "」" for q in QUESTION_TYPES), 7.6, SUB)
    s.ruled(c3, y + 8.0, w3, 4, 8.0)
    s.text(W / 2, 291, "B2 班用ワークシート（裏）", 7, SUB, align="c")


def build_b2():
    c, s, path = new_canvas("B2_班用ワークシート.pdf", A3L)
    b2_front(s)
    c.showPage()
    b2_back(s)
    c.showPage()
    c.save()
    return path


# =====================================================================
# B3 ミッションカード（A5・横・4枚）
# =====================================================================
def build_b3():
    c, s, path = new_canvas("B3_ミッションカード.pdf", A5L)
    W, H = A5L
    for name, purpose, hint in MISSIONS:
        s.rect(5, 5, W - 10, H - 10, lw=1.2, stroke=ACCENT, r=3)
        s.rect(5, 5, W - 10, 11, lw=0, stroke=None, fill=ACCENT, r=3)
        s.rect(5, 12, W - 10, 4, lw=0, stroke=None, fill=ACCENT)
        s.text(12, 13.2, "ミッションカード", 10, WHITE, bold=True)
        s.text(W - 70, 13.2, "班", 10, WHITE)
        s.line(W - 62, 14.0, W - 12, 14.0, 0.5, WHITE)
        s.text(14, 38, name, 30, ACCENT, bold=True)
        s.line(14, 42, W - 14, 42, 0.6, ACCENT)
        s.para(14, 47, W - 28, purpose, 14, leading=21)
        yq = 76
        s.text(14, yq, "問い　" + MISSION_Q, 13, INK, bold=True)
        for i in range(3):
            yy = yq + 11 + i * 11
            s.text(16, yy, f"{i + 1}.", 12, ACCENT)
            s.line(24, yy + 1.2, W - 14, yy + 1.2, 0.4, RULE)
        if hint:
            s.rect(14, 122, W - 28, 13, lw=0.6, stroke=ACCENT, fill=HEAD_BG, r=1.5)
            s.text(18, 130.4, "ヒント", 9, ACCENT, bold=True)
            s.text(36, 130.4, hint, 12, INK)
        c.showPage()
    c.save()
    return path


# =====================================================================
# B4 地域カード・番号つき全球図
# =====================================================================
# 全球図の番号の置き方：番号の丸を置く位置(経度, 緯度)。枠の中心から細い線でつなぐ。位置による色分けや強調はしない。
LABEL_POS = {1: None, 2: None, 3: (-36, 37), 4: (-82, 6), 5: None, 6: None, 7: (0, -76), 8: (-80, -18)}
LABEL_ANCHOR = {3: (-42, 30), 4: (-59, 14), 8: (-70, -8), 7: (0, -87)}


def region_center(rg):
    name, la0, la1, lo0, lo1 = rg
    lat = (la0 + la1) / 2
    lon = (lo0 + lo1) / 2
    if lo1 > 180:
        lon = ((lo0 + lo1) / 2 + 180) % 360 - 180
    return lat, lon


def draw_all_boxes(mm_, regions, r=2.6, size=8):
    for i, rg in enumerate(regions, 1):
        name, la0, la1, lo0, lo1 = rg
        mm_.box(la0, la1, lo0, lo1, lw=0.9)
    for i, rg in enumerate(regions, 1):
        lat, lon = region_center(rg)
        pos = LABEL_POS.get(i)
        if pos is None:
            mm_.label(lon, lat, i, 0, 0, r=r, size=size)
        else:
            mm_.label(lon, lat, i, pos[0] - lon, pos[1] - lat, r=r, size=size, anchor=LABEL_ANCHOR.get(i))


def build_b4a(regions):
    c, s, path = new_canvas("B4a_地域カード.pdf", A4P)
    cw_, ch = 105.0, 148.5
    for page in range(2):
        for k in range(4):
            i = page * 4 + k
            rg = regions[i]
            name, la0, la1, lo0, lo1 = rg
            cx = (k % 2) * cw_
            cy = (k // 2) * ch
            s.rect(cx + 3, cy + 3, cw_ - 6, ch - 9, lw=0.8, stroke=ACCENT, r=3, dash=(3, 2))
            s.circle(cx + 14, cy + 17, 7, stroke=ACCENT, fill=ACCENT)
            s.text(cx + 14, cy + 20.8, str(i + 1), 18, WHITE, bold=True, align="c")
            s.text(cx + 25, cy + 11.5, "地域カード", 7.5, SUB)
            main, _, sub = name.partition("（")
            h = s.para(cx + 25, cy + 13.0, cw_ - 33, main, 13, leading=16, color=INK)
            if sub:
                h += s.para(cx + 25, cy + 13.0 + h, cw_ - 33, "（" + sub, 9.5, leading=12, color=SUB)
            yy = cy + max(13.0 + h, 25) + 3
            s.line(cx + 8, yy, cx + cw_ - 8, yy, 0.6, ACCENT)
            yy += 3.5
            lat_s = "緯度　" + rng(la0, la1)
            lon_s = "経度　" + lon_str(lo0, lo1)
            s.text(cx + 9, yy + 4.5, lat_s, 11.5, INK)
            s.text(cx + 9, yy + 11.5, lon_s, 11.5, INK)
            yy += 15
            # 縮小全球図（この地域の枠だけ）
            mw = cw_ - 20
            mp = MoonMap(s, cx + 10, yy + 1, mw, alpha=0.35)
            mp.grid(30, lw=0.15)
            mp.box(la0, la1, lo0, lo1, lw=1.0)
            lat, lon = region_center(rg)
            s.circle(mp.px(lon), mp.py(lat), 3.2, stroke=MARK, fill=None, lw=0.5)
            s.text(cx + 10, yy + 1 + mw / 2 + 3.6, "月の全球図（赤い枠・丸が、この地域の範囲）", 6.5, SUB)
            yy += mw / 2 + 5.5
            s.line(cx + 8, yy, cx + cw_ - 8, yy, 0.3, RULE)
            line = REGION_LINES[name]
            hh = s.para(cx + 9, yy + 2.0, cw_ - 18, line, 9.6, leading=14)
            yy += 2.0 + hh + 2.5
            s.text(cx + 9, yy + 3.2, "気づいたこと（メモ）", 7.5, SUB)
            s.ruled(cx + 8, yy + 3.0, cw_ - 16, 2, 6.8)
            s.text(cx + cw_ / 2, cy + ch - 8.0, "経度は東が＋、緯度は北が＋", 6.2, SUB, align="c")
        s.text(A4P[0] / 2, 293.6, FOOT_NOTE_REGION, 6.0, SUB, align="c")
        s.text(A4P[0] / 2, 296.2, "経度が180°をこえる地域は、日付変更線をまたぐので「東経〜西経」で書く。", 6.0, SUB, align="c")
        c.showPage()
    c.save()
    return path


def build_b4b(regions):
    c, s, path = new_canvas("B4b_地域カード_全球図.pdf", A4L)
    W, H = A4L
    s.text(10, 13, "地域カードの場所（番号は地域カードの番号）", 14, ACCENT, bold=True)
    s.line(10, 16, W - 10, 16, 0.8, ACCENT)
    mw = 250
    mx, my = 21, 29
    mp = MoonMap(s, mx, my, mw, alpha=0.3)
    mp.grid(30, lw=0.15)
    mp.ticks(6)
    draw_all_boxes(mp, regions)
    s.text(mx + mw / 2, my + mw / 2 + 11.5, "経度（°）　東が＋", 7.5, SUB, align="c")
    s.rotated_text(11, my + mw / 4, "緯度（°）　北が＋", 7.5, SUB)
    # 凡例
    yy = my + mw / 2 + 18
    for i, rg in enumerate(regions):
        col, row = i // 4, i % 4
        x = 10 + col * 105
        y = yy + row * 8.0
        s.circle(x + 3.2, y + 2.2, 2.9, stroke=MARK, fill=WHITE, lw=0.6)
        s.text(x + 3.2, y + 3.4, str(i + 1), 8, MARK, bold=True, align="c")
        s.text(x + 9, y + 3.3, rg[0], 8.2, INK)
    # 重なる3つの拡大
    ix, iy, iw = 244, 163, 42
    lon0, lon1, lat0, lat1 = -78, -32, -12, 34
    s.rect(ix, iy, iw, iw * (lat1 - lat0) / (lon1 - lon0), lw=0.5, stroke=INK)
    from PIL import Image
    im = Image.open(faded_moon(0.3))
    W_, H_ = im.size
    box = (int((lon0 + 180) / 360 * W_), int((90 - lat1) / 180 * H_), int((lon1 + 180) / 360 * W_), int((90 - lat0) / 180 * H_))
    crop = im.crop(box)
    cp = os.path.join(os.path.dirname(faded_moon(0.3)), "crop.jpg")
    crop.save(cp, quality=88)
    ih = iw * (lat1 - lat0) / (lon1 - lon0)
    c.drawImage(cp, ix * mm, s.Y(iy + ih), iw * mm, ih * mm)
    s.rect(ix, iy, iw, ih, lw=0.5, stroke=INK)

    def ixy(lon, lat):
        return ix + (lon - lon0) / (lon1 - lon0) * iw, iy + (lat1 - lat) / (lat1 - lat0) * ih
    for i, rg in enumerate(regions, 1):
        name, la0, la1, lo0, lo1 = rg
        if lo0 >= lon0 and lo1 <= lon1 and la0 >= lat0 and la1 <= lat1:
            x0, y0 = ixy(lo0, la1)
            x1, y1 = ixy(lo1, la0)
            s.rect(x0, y0, x1 - x0, y1 - y0, lw=0.8, stroke=MARK)
    pos = {3: (-42, 30), 4: (-56, 10), 8: (-70, -4)}
    for i, (lo, la) in pos.items():
        x, y = ixy(lo, la)
        s.circle(x, y, 2.3, stroke=MARK, fill=WHITE, lw=0.6)
        s.text(x, y + 1.15, str(i), 7.5, MARK, bold=True, align="c")
    s.text(ix - 2, iy + 4, "拡大（3・4・8）", 7, SUB, align="r")
    c.showPage()
    c.save()
    return path


# =====================================================================
# B4c クレーターカード（第2時の判断②。A4に4枚＝A6×4、切って班に1枚）
# =====================================================================
# 数値（内部・帯の温度）は書かない。生徒は表計算ステップ1cで読む。判定・年代・原因も書かない。
CRATER_TERMS = [
    ("内部", "中心を囲む四角（半径の半分ほど）の中の0.5度のセル。"),
    ("同じ緯度の帯", "クレーターと同じ緯度（±1.5度）を、月を1周したもの（クレーターから経度が直径1つ分以上はなれたセルだけ）。"),
    ("帯の9割が入る範囲", "帯の中のセルの9割が、この下端と上端の間に入る。"),
]
CRATER_STEPS = [
    "表計算「ステップ1c」の黄色いセルB5で、このカードの名前を選ぶ。",
    "24時間の温度カーブ（内部と帯の2本）を見る。",
    "「夜の最低温度」の行の4つの数字（内部／帯の平均／帯の9割が入る範囲の下端・上端）を読む。",
    "内部の値は、帯の9割が入る範囲の外か内か。個人用ワークシート⑥(b)に、数字と、そう決めた理由を書く。",
]


def build_b4c(craters):
    c, s, path = new_canvas("B4c_クレーターカード.pdf", A4P)
    cw_, ch = 105.0, 148.5
    for k, name in enumerate(CRATER_NAMES):
        lat, lon, dia = craters[name]
        cx = (k % 2) * cw_
        cy = (k // 2) * ch
        s.rect(cx + 3, cy + 3, cw_ - 6, ch - 9, lw=0.8, stroke=ACCENT, r=3, dash=(3, 2))
        s.text(cx + 9, cy + 11.5, "クレーターカード（第2時）", 7.5, SUB)
        s.text(cx + cw_ - 9, cy + 11.5, "班", 7.5, SUB, align="r")
        s.line(cx + cw_ - 40, cy + 12.2, cx + cw_ - 13, cy + 12.2, 0.35, LINE)
        s.text(cx + 9, cy + 25.5, name, 22, ACCENT, bold=True)
        yy = cy + 29.5
        s.line(cx + 8, yy, cx + cw_ - 8, yy, 0.6, ACCENT)
        yy += 4.0
        s.text(cx + 9, yy + 4.0, "緯度　" + num(round(lat, 2)) + "°　経度　" + num(round(lon, 2)) + "°", 10.5, INK)
        s.text(cx + 9, yy + 10.6, "直径　約" + num(round(dia)) + " km", 10.5, INK)
        yy += 14.5
        s.text(cx + 9, yy + 3.2, "やること（表計算）", 9.6, ACCENT, bold=True)
        yy += 5.0
        for i, st in enumerate(CRATER_STEPS, 1):
            s.circle(cx + 12, yy + 2.5, 2.2, stroke=ACCENT, fill=ACCENT)
            s.text(cx + 12, yy + 3.55, str(i), 7, WHITE, bold=True, align="c")
            h = s.para(cx + 16.5, yy, cw_ - 28, st, 9.4, leading=12.8)
            yy += max(6.5, h + 2.6)
        yy += 1.5
        s.line(cx + 8, yy, cx + cw_ - 8, yy, 0.3, RULE)
        yy += 1.2
        s.text(cx + 9, yy + 3.2, "ことば", 9.6, ACCENT, bold=True)
        yy += 5.0
        for term, desc in CRATER_TERMS:
            h = s.para(cx + 9, yy, cw_ - 18, "<b>" + escape(term) + "</b>：" + escape(desc), 8.6, leading=11.8, raw=True)
            yy += h + 2.0
        s.text(cx + cw_ / 2, cy + ch - 8.0, "このカードに数字は書いていない。数字は表計算で読む。", 6.2, SUB, align="c")
    c.showPage()
    c.save()
    return path


# =====================================================================
# B5 黒板用の記入表／付箋用の月全球図（A3・横）
# =====================================================================
def build_b5a():
    c, s, path = new_canvas("B5a_黒板用記入表.pdf", A3L)
    W, H = A3L
    s.text(12, 22, "5地点の記入表", 28, ACCENT, bold=True)
    s.text(12 + sw("5地点の記入表", 28) + 8, 22, "ピンカードの表から読んだ値と、出典（表の行の名前）を書く", 15, SUB)
    s.line(10, 27, W - 10, 27, 1.5, ACCENT)
    # 列の並びと名前は、Webのピンカードの表（1日の温度差・夜の最低温度・正午の太陽高度・地球の仰角）と同じ
    colw = cw(404, [82, 50, 46, 50, 44, 60, 72])
    head = ["地点", "1日の温度差", "夜の最低温度", "正午の太陽高度", "地球の仰角", "出典（表の行の名前）", "読んだ班"]
    units = ["（緯度、経度）", "（日較差）　℃", "℃", "°", "°", "", ""]
    x = 8
    y = 33
    hh = 38
    rh = 39
    xx = x
    for i, wd in enumerate(colw):
        s.rect(xx, y, wd, hh, lw=1.4, stroke=INK, fill=HEAD_BG)
        txt = head[i]
        fs = next((f for f in (24, 22, 20, 19, 18, 17, 16) if sw(txt, f) < wd - 4), 15)
        if units[i]:
            s.text(xx + wd / 2, y + 17, txt, fs, INK, bold=True, align="c")
            s.text(xx + wd / 2, y + 30, units[i], 15, SUB, align="c")
        else:
            s.text(xx + wd / 2, y + 23, txt, fs, INK, bold=True, align="c")
        xx += wd
    y += hh
    for n, nm, la, lo in JIGSAW_SITES:
        xx = x
        for i, wd in enumerate(colw):
            s.rect(xx, y, wd, rh, lw=1.4, stroke=INK, fill=None)
            xx += wd
        # 1地点は2班が読む：値の欄を上下2段に分ける
        s.line(x + colw[0], y + rh / 2, x + sum(colw), y + rh / 2, 0.6, LINE, dash=(3, 2))
        for k in range(2):
            s.text(x + sum(colw[:6]) + 3, y + 13.5 + k * rh / 2, "班", 14, SUB)
        s.text(x + 4, y + 16, n + "　" + nm, 16.5, INK, bold=True)
        s.text(x + 6, y + 30, f"（{num(la)}°、{num(lo)}°）", 17, SUB)
        y += rh
    s.text(12, y + 10, "座標欄に緯度・経度を入れ、ピンカードの表から4つを読む。地球の仰角：＋は地球が空に見える、−は地平線の下。温度は℃。", 13, SUB)
    s.text(12, y + 18, "1地点は2班が読む（上段と下段に、1班ずつ書く）。ちがったら、入力した緯度・経度と、表の下の「セルの中心」を確かめる。", 13, SUB)
    c.showPage()
    c.save()
    return path


def build_b5b():
    c, s, path = new_canvas("B5b_付箋用月全球図.pdf", A3L)
    W, H = A3L
    s.text(10, 15, "月の全球図（付箋用）", 20, ACCENT, bold=True)
    s.text(10 + sw("月の全球図（付箋用）", 20) + 6, 15, "決めた1地点に、付箋を貼る（付箋に班・緯度経度を書く）", 12, SUB)
    s.line(10, 19, W - 10, 19, 1.2, ACCENT)
    mw = 372
    mx, my = 24, 54
    mp = MoonMap(s, mx, my, mw, alpha=0.5)
    mh = mw / 2
    band_y = my - 15
    bands = [(-180, -90, "裏側（地球に向かない面）", HEAD_BG), (-90, 90, "表側（地球に向いた面）", colors.HexColor("#fff2cc")),
             (90, 180, "裏側（地球に向かない面）", HEAD_BG)]
    for a, b, lab, fillc in bands:
        s.rect(mp.px(a), band_y - 8.5, mp.px(b) - mp.px(a), 8.5, lw=0.6, stroke=INK, fill=fillc)
        s.text((mp.px(a) + mp.px(b)) / 2, band_y - 2.2, lab, 11.5, INK, bold=True, align="c")
    mp.grid(15, lw=0.25)
    mp.ticks(7.5)
    s.text(mx + mw / 2, my + mh + 13, "経度（°）　東が＋・西が−　目盛は10°ごと、薄い線は15°ごと", 10, SUB, align="c")
    s.rotated_text(10, my + mh / 2, "緯度（°）　北が＋・南が−", 10, SUB)
    s.text(10, 275, "経度0°の線が、表側の中心。月の画像は、書き込みが見えるよう薄くしてある。", 10, SUB)
    mmpd = mw / 360.0   # 地図の1°あたりのmm（A3を原寸で印刷したとき）
    s.text(10, 282.5, f"付箋の大きさ：この図の1°は約{mmpd:.2f} mm（A3を拡大・縮小せず、原寸で印刷する）。付箋は小さめ（約25 mm角＝約{25 / mmpd:.0f}°分）、"
           f"または点シール（直径約10 mm＝約{10 / mmpd:.0f}°分）を、1つの班が1枚使う。", 10, SUB)
    s.text(10, 289, "付箋の中央（点シールは中心）を、決めた地点に合わせて貼り、班と緯度・経度を書く。", 10, SUB)
    c.showPage()
    c.save()
    return path


# =====================================================================
# B8 班編成・ペア確定シート（A4・教員用）
# =====================================================================
RULES_B8 = [
    "同じ地域タイプ同士は組まない。",
    "同じタイプが班の過半数で避けられないときは、ミッションが違う班どうしを優先する（それも難しければ、教員が決める）。",
    "班が奇数なら、3班の輪にする（X→Y→Z→X）。",
    "変更があれば、第3時の冒頭（0:03〜0:07）で確定する。",
]
B8_EXAMPLE = [  # 班, ミッション, ジグソー地点（A〜E）, 地域カード番号（①〜⑧）, ペア, 第2時の担当クレーター
    (1, "氷採掘", "C", "⑦", "3", "プトレマイオス"), (2, "氷採掘", "D", "⑦", "4", "アルフォンスス"),
    (3, "電波天文台", "B", "⑥", "1", "ティコ"), (4, "電波天文台", "E", "⑥", "2", "プトレマイオス"),
    (5, "電波天文台", "D", "⑤", "6", "ラングレヌス"),
    (6, "太陽光発電", "A", "①", "5", "アルフォンスス"), (7, "太陽光発電", "C", "⑧", "8", "プトレマイオス"),
    (8, "太陽光発電", "B", "①", "7", "ラングレヌス"),
    (9, "有人総合", "A", "①", "10", "アルフォンスス"), (10, "有人総合", "E", "④", "9", "プトレマイオス"),
]

# --- 第2時：担当クレーター（教員の1例はコペルニクス。担当は4つ。設計 §3.3）
CRATER_NAMES = ["ティコ", "ラングレヌス", "プトレマイオス", "アルフォンスス"]
CRATER_EXAMPLE_NAME = "コペルニクス"
CRATER_SIDE = ["ティコ", "ラングレヌス"]     # B8 の運用で、班数を全班の30％以下に抑える2つ（教員用。生徒には見せない）
CRATER_SHARE_MAX = 0.30
# 班数別の割当の目安（設計 §3.3）：班数 → {クレーター: 班数}
B8_CRATER_PLAN = {
    8: {"ティコ": 1, "ラングレヌス": 1, "プトレマイオス": 3, "アルフォンスス": 3},
    10: {"ティコ": 1, "ラングレヌス": 2, "プトレマイオス": 4, "アルフォンスス": 3},
    12: {"ティコ": 1, "ラングレヌス": 2, "プトレマイオス": 5, "アルフォンスス": 4},
}
RULES_B8_CRATER = [
    "担当は4つ（" + "・".join(CRATER_NAMES) + "）。教員の1例の「" + CRATER_EXAMPLE_NAME + "」は、班の担当にしない。",
    "ティコ・ラングレヌスを担当する班は、全班の30％以下にする。",
    "班が6以下で、30％以下にできないときは、ラングレヌスを割り当てない。",
    "割当は、班番号で事前に決めて、クレーターカード（B4c）を班に1枚ずつ配る。",
]
SITE_SHEET = "データ_地点比較"


def load_craters_xlsx():
    """学習者版の『データ_地点比較』から、クレーターの名前・緯度・経度・直径を読む。"""
    import openpyxl
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True, read_only=True)
    ws = wb[SITE_SHEET]
    rows = {}
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or r[0] is None:
            continue
        rows[str(r[0])] = (float(r[1]), float(r[2]), float(r[3]))
    wb.close()
    return rows


B8_STAMP_NOTE = ("全員分の確認印を、まとめの3分で押しきらない。印は班ごとに1枚か、抜き取りで押す。"
                 "判断メモ（B1）の回収は第2時の終わりにまとめて行い、第3時の冒頭に返す。"
                 "見るのは「数字が画面と合っているか」「理由の欄が書いてあるか」の2点だけ（判断の中身には正解を付けない。評価はしない）。"
                 "この運用で時間が足りるかは未確認。")


def b8_page2(s):
    """B8 p2：第2時の担当クレーターの割当（教員用）。"""
    W, H = A4P
    s.text(10, 15, "第2時　担当クレーターの割当", 17, ACCENT, bold=True)
    s.text(10 + sw("第2時　担当クレーターの割当", 17) + 4, 15, "教員用（生徒には配らない・映さない）", 10, SUB)
    s.line(10, 18.5, W - 10, 18.5, 1.0, ACCENT)
    y = 22
    s.text(10, y + 3.5, "決め方の規則", 10.5, ACCENT, bold=True)
    y += 6
    for i, r in enumerate(RULES_B8_CRATER, 1):
        s.circle(14.5, y + 3.2, 2.8, stroke=ACCENT, fill=ACCENT)
        s.text(14.5, y + 4.4, str(i), 8.5, WHITE, bold=True, align="c")
        hh_ = s.para(20, y + 0.6, W - 32, r, 10.5, leading=14)
        y += max(7.2, hh_ + 2.4)
    y += 2.0

    s.text(10, y + 3.5, "班の数ごとの割当の目安（班数）", 10, ACCENT, bold=True)
    y += 5.5
    colw = [24] + [31] * 4 + [19, 23]
    cells = [["班の数"] + CRATER_NAMES + ["ティコ＋\nラングレヌス", "割合"]]
    for n, plan in B8_CRATER_PLAN.items():
        side = sum(plan[k] for k in CRATER_SIDE)
        cells.append([f"{n}班"] + [str(plan[k]) for k in CRATER_NAMES] + [str(side), f"{side / n * 100:.0f}％"])
    y = s.grid(10, y, colw, [11] + [7.2] * len(B8_CRATER_PLAN), cells, size=9.5, head_size=8.5) + 2.0
    s.text(10, y + 3.2, "どの班にどのクレーターを渡すかは、教員が班番号で事前に決める。上の数は目安。班の数がこれ以外のときは、規則2・3に合わせて決める。", 7.8, SUB)
    y += 10.0

    s.text(10, y + 3.5, "今回の割当（班の番号を書く）", 10, ACCENT, bold=True)
    y += 5.5
    colw = [40, 100, 24, 26]
    cells = [["担当クレーター", "担当する班の番号", "班の数", "規則2の対象"]]
    for k in CRATER_NAMES:
        cells.append([k, "", "", "対象" if k in CRATER_SIDE else ""])
    y = s.grid(10, y, colw, [7.5] + [11.0] * 4, cells, size=10) + 3.0
    s.blank(10, y + 5.5, 190, "確認：ティコ＋ラングレヌスの班数（　　）÷ 全班の数（　　）＝", 10, "％（30％以下か）")
    y += 14.0
    s.text(10, y + 3.5, "班ごとの記録は、1ページ目の表の「第2時の担当クレーター」の列に書く。", 8.6, SUB)
    y += 9.0
    s.text(10, y + 3.5, "B1⑥（判断メモ）の確認印の運用（案）", 10, ACCENT, bold=True)
    y += 5.5
    h = s.para(10, y + 0.6, W - 20, B8_STAMP_NOTE, 9.5, leading=13.5)
    y += h + 2.0
    s.text(W / 2, 292, "B8 班編成・ペア確定シート（教員用）　2/2　第2時の担当クレーター", 6.5, SUB, align="c")


def build_b8():
    c, s, path = new_canvas("B8_班編成ペア確定シート_教員用.pdf", A4P)
    W, H = A4P
    s.text(10, 15, "班編成・ペア確定シート", 17, ACCENT, bold=True)
    s.text(10 + sw("班編成・ペア確定シート", 17) + 4, 15, "教員用（生徒には配らない）", 10, SUB)
    s.line(10, 18.5, W - 10, 18.5, 1.0, ACCENT)
    y = 22
    s.text(10, y + 3.5, "決め方の規則", 10.5, ACCENT, bold=True)
    y += 6
    for i, r in enumerate(RULES_B8, 1):
        s.circle(14.5, y + 3.2, 2.8, stroke=ACCENT, fill=ACCENT)
        s.text(14.5, y + 4.4, str(i), 8.5, WHITE, bold=True, align="c")
        hh_ = s.para(20, y + 0.6, W - 32, r, 10.5, leading=14)
        y += max(7.2, hh_ + 2.4)
    h = s.para(10, y + 1.0, W - 20,
               "ペア（異なる地域タイプの班どうし）は、第1時の終了後、地域タイプが決まってから確定する。"
               "ジグソーの地点（A〜E）は、同じ地点を読む2班で組む（ペアとは別）。地域タイプは地域カードの番号（①〜⑧）。ミッションの割当は、班番号で事前に決める。",
               8.3, color=SUB)
    y += h + 4.5

    s.text(10, y + 3.5, "記入例（10班。ミッションは氷採掘2・電波天文台3・太陽光発電3・有人総合2。数字は例）", 10, ACCENT, bold=True)
    y += 5.5
    colw = [11, 27, 20, 33, 17, 33, 49]
    HEAD = ["班", "ミッション", "ジグソー地点\n（A〜E）", "地域タイプ（第1時終了時）\n地域カード①〜⑧", "ペア\n（班）", "第2時の\n担当クレーター", "発表者\n（当日くじ）"]
    cells = [HEAD]
    for b, m, j, t, p, cr in B8_EXAMPLE:
        cells.append([str(b), m, j, t, p, cr, ""])
    fl = {}
    rh = [11] + [6.8] * 10
    y = s.grid(10, y, colw, rh, cells, size=9, head_size=7.8, heads_left=0) + 1.5
    s.text(10, y + 3.2, "地域タイプは地域カードの番号。ペアの相手は、地域タイプが違う班（上の例は5組すべて違う）。", 7.8, SUB)
    s.text(10, y + 8.0, "奇数のとき：例）班9・10・11 → 9→10→11→9 の輪にする。", 7.8, SUB)
    y += 13.0

    s.text(10, y + 3.5, "記入欄", 10, ACCENT, bold=True)
    y += 5.5
    cells = [HEAD]
    for _ in range(12):
        cells.append(["", "", "", "", "", "", ""])
    rh = [11] + [7.6] * 12
    y = s.grid(10, y, colw, rh, cells, size=9, head_size=7.8)
    s.text(W / 2, 292, "B8 班編成・ペア確定シート（教員用）　1/2　裏面は第2時の担当クレーター", 6.5, SUB, align="c")
    c.showPage()
    b8_page2(s)
    c.showPage()
    c.save()
    return path


# =====================================================================
# B9 Web操作カード（A5・縦）
# =====================================================================
def icon(s, kind, cx, cy, r=5.2):
    s.circle(cx, cy, r, stroke=INK, fill=WHITE, lw=0.7)
    c = s.c
    if kind == "play":
        p = c.beginPath()
        p.moveTo((cx - 1.6) * mm, s.Y(cy - 2.4))
        p.lineTo((cx + 2.6) * mm, s.Y(cy))
        p.lineTo((cx - 1.6) * mm, s.Y(cy + 2.4))
        p.close()
        c.setFillColor(INK)
        c.drawPath(p, stroke=0, fill=1)
    elif kind == "image":
        s.rect(cx - 3.2, cy - 2.6, 6.4, 5.2, lw=0.6, stroke=INK)
        p = c.beginPath()
        p.moveTo((cx - 3.0) * mm, s.Y(cy + 2.4))
        p.lineTo((cx - 0.6) * mm, s.Y(cy - 0.6))
        p.lineTo((cx + 0.8) * mm, s.Y(cy + 1.0))
        p.lineTo((cx + 1.8) * mm, s.Y(cy - 0.4))
        p.lineTo((cx + 3.0) * mm, s.Y(cy + 2.4))
        p.close()
        c.setFillColor(INK)
        c.drawPath(p, stroke=0, fill=1)
        s.circle(cx + 1.6, cy - 1.4, 0.7, stroke=None, fill=INK)
    elif kind == "thermo":
        s.rect(cx - 0.8, cy - 3.6, 1.6, 5.4, lw=0.6, stroke=INK)
        s.circle(cx, cy + 2.3, 1.7, stroke=INK, fill=INK, lw=0.4)
    elif kind == "wire":
        s.circle(cx, cy, 3.4, stroke=INK, lw=0.5)
        s.line(cx - 3.4, cy, cx + 3.4, cy, 0.4, INK)
        s.line(cx, cy - 3.4, cx, cy + 3.4, 0.4, INK)
        c.setStrokeColor(INK)
        c.setLineWidth(0.4)
        c.ellipse((cx - 1.5) * mm, s.Y(cy + 3.4), (cx + 1.5) * mm, s.Y(cy - 3.4), stroke=1, fill=0)
    elif kind == "zoom":
        s.text(cx, cy + 1.6, "＋", 9, INK, bold=True, align="c")
        s.text(cx + 0.0, cy + 6.0, "", 6)
    elif kind == "reset":
        c.setStrokeColor(INK)
        c.setLineWidth(0.7)
        c.arc((cx - 3) * mm, s.Y(cy + 3), (cx + 3) * mm, s.Y(cy - 3), 40, 280)
        p = c.beginPath()
        p.moveTo((cx + 3.6) * mm, s.Y(cy - 2.6))
        p.lineTo((cx + 1.0) * mm, s.Y(cy - 3.0))
        p.lineTo((cx + 3.2) * mm, s.Y(cy - 0.2))
        p.close()
        c.setFillColor(INK)
        c.drawPath(p, stroke=0, fill=1)


def build_b9():
    c, s, path = new_canvas("B9_Web操作カード.pdf", A5P)
    W, H = A5P
    x0, w = 8, W - 16

    def title(tag):
        s.text(x0, 13, "Webアプリ操作カード", 14, ACCENT, bold=True)
        s.text(W - 8, 13, "班に1枚　" + tag, 8.5, SUB, align="r")
        s.line(x0, 15.5, W - x0, 15.5, 0.9, ACCENT)
        return 18.0

    def section(y, head, keys, fill=ACCENT, size=9.8, bullets=True):
        s.bar(x0, y, w, 5.8, head, 9, fill=fill)
        y += 8.0
        for key in keys:
            txt = WEB_TEXT[key] if key in WEB_TEXT else key
            if bullets:
                s.text(x0 + 0.8, y + 2.7, "・", size)
                h = s.para(x0 + 4, y, w - 4, txt, size, leading=size * 1.5)
            else:
                h = s.para(x0, y, w, txt, size, leading=size * 1.5)
            y += h + 1.4
        return y + 2.0

    # ---------- 表：ボタンと層
    y = title("（表）")
    h = s.para(x0, y, w, WEB_TEXT["start"], 8.5, color=SUB)
    y += h + 2.5
    s.bar(x0, y, w, 5.8, "右下のボタン（上から）", 9)
    y += 8.5
    btns = [("play", "btn_play"), ("image", "btn_image"), ("thermo", "btn_layer"),
            ("wire", "btn_wire"), ("zoom", "btn_zoom"), ("reset", "btn_reset")]
    for i, (k, key) in enumerate(btns):
        cy = y + 6.0 + i * 13.5
        icon(s, k, x0 + 6.5, cy)
        lab = WEB_TEXT[key]
        s.para(x0 + 16, cy - 4.2, w - 18, lab, 9.4, leading=12)
    y += 6 * 13.5 + 2.5
    h = s.para(x0, y, w, WEB_TEXT["buttons_note"], 8.8, color=SUB)
    y += h + 4.0
    y = section(y, "層を切り替える（データ層をオンにしてから）", ["layer_panel", "layers", "diurnal"], size=9.8, bullets=False)
    s.text(W / 2, H - 4.5, "B9 Web操作カード（表）　裏面は、地点を決める・値を読む", 6.2, SUB, align="c")
    c.showPage()

    # ---------- 裏：地点を決める・読む
    y = title("（裏）")
    y = section(y, "地点を決めてピンを立てる", ["coord_input", "pin_click", "narrow", "hover", "follow", "minimap"], size=9.0)
    y = section(y, "ピンカードを読む", ["jigsaw", "pin_card", "pin_card_diurnal", "cell", "temp_unit"], size=9.0)
    y = section(y, "知っておくこと", ["place_hint", "data_list"], fill=MARK, size=9.0)
    s.text(W / 2, H - 4.5, "B9 Web操作カード（裏）", 6.2, SUB, align="c")
    c.showPage()
    c.save()
    return path, y


# =====================================================================
# 出力と検査
# =====================================================================
OUTPUTS = {
    # ファイル名: (用紙(幅,高さ)mm, ページ数)
    "B1_個人用ワークシート.pdf": (B5, 2),
    "B2_班用ワークシート.pdf": (A3L, 2),
    "B3_ミッションカード.pdf": (A5L, 4),
    "B4a_地域カード.pdf": (A4P, 2),
    "B4b_地域カード_全球図.pdf": (A4L, 1),
    "B4c_クレーターカード.pdf": (A4P, 1),
    "B5a_黒板用記入表.pdf": (A3L, 1),
    "B5b_付箋用月全球図.pdf": (A3L, 1),
    "B8_班編成ペア確定シート_教員用.pdf": (A4P, 2),
    "B9_Web操作カード.pdf": (A5P, 2),
}


def get_regions():
    """表示名の基準：データ_地域（学習者版）と course/data/candidate_regions.csv。"""
    xr = load_regions_xlsx()
    cr = load_regions_csv()
    return xr, cr


def orig_regions_as_display():
    """元データ（data/candidate_regions.csv）の旧名を、表示名に直したもの（名前の対応は NAME_ALIAS）。"""
    inv = {v: k for k, v in NAME_ALIAS.items()}
    return [(inv.get(n, n), a, b, c, d) for n, a, b, c, d in load_regions_csv(CSV_ORIG_PATH)]


def build_all():
    xr, cr = get_regions()
    if xr != cr:
        print("警告：データ_地域 と course/data/candidate_regions.csv が一致しない。xlsx（データ_地域）を使う。")
    names = [r[0] for r in xr]
    missing = [n for n in names if n not in REGION_LINES]
    if missing:
        raise SystemExit("REGION_LINES にない地域名：" + str(missing))
    out = [build_b1(), build_b2(), build_b3(), build_b4a(xr), build_b4b(xr), build_b4c(load_craters_xlsx()), build_b5a(), build_b5b(), build_b8()]
    p9, yend = build_b9()
    out.append(p9)
    for p in out:
        print("出力：", os.path.relpath(p, ROOT))
    return out


def pdf_text(path):
    r = subprocess.run(["pdftotext", "-enc", "UTF-8", "-raw", path, "-"], capture_output=True)
    t = r.stdout.decode("utf-8", errors="replace")
    return re.sub(r"\s+", "", t)


def check_all():
    import pypdf
    ng = []
    ok = []

    def rep(cond, msg):
        (ok if cond else ng).append(msg)

    # --- 用紙・ページ数
    for fn, (size, pages) in OUTPUTS.items():
        p = os.path.join(OUT_DIR, fn)
        if not os.path.exists(p):
            rep(False, f"{fn}：ファイルがない")
            continue
        r = pypdf.PdfReader(p)
        rep(len(r.pages) == pages, f"{fn}：ページ数 {len(r.pages)}（期待{pages}）")
        for i, pg in enumerate(r.pages, 1):
            w = float(pg.mediabox.width) / 72 * 25.4
            h = float(pg.mediabox.height) / 72 * 25.4
            good = abs(w - size[0]) < 0.6 and abs(h - size[1]) < 0.6
            rep(good, f"{fn} p{i}：メディアボックス {w:.1f}×{h:.1f} mm（期待 {size[0]}×{size[1]}）")

    # --- 地域：データ_地域 と candidate_regions.csv
    xr, cr = get_regions()
    rep(xr == cr, "データ_地域（学習者版）と course/data/candidate_regions.csv の名前・範囲が一致")
    rep(orig_regions_as_display() == xr, "元データ data/candidate_regions.csv と範囲が一致（名前は表示名の対応 NAME_ALIAS で読み替え）")
    rep(len(xr) == 8, f"地域は8件（{len(xr)}件）")
    rep(all(n in REGION_LINES for n, *_ in xr), "REGION_LINES の地域名がデータ_地域のA列と完全一致（全件）")
    rep(set(REGION_LINES) == {n for n, *_ in xr}, "REGION_LINES に余分な地域名がない")

    # --- カード本文（データ構造）
    for name, purpose, hint in MISSIONS:
        body = name + purpose + MISSION_Q + (hint or "")
        body2 = body.replace(ICE_HINT, "")
        bad = [w for w in FORBIDDEN_CARD if w in body2]
        rep(not bad, f"ミッションカード『{name}』本文に禁止語なし{('（検出：' + ','.join(bad) + '）') if bad else ''}")
        rep(not FORBIDDEN_PATTERN.search(body2), f"ミッションカード『{name}』に「〜は／〜なら」型の対応文なし")
        for rn in REGION_LINES:
            if rn[:3] in body2 or "赤道" in body2 or "火砕" in body2:
                rep(False, f"ミッションカード『{name}』に地域名が入っている")
                break
        if name != "氷採掘":
            rep(ICE_HINT not in body and "氷" not in body and "冷た" not in body, f"『{name}』に氷のヒント・氷の語がない")
        else:
            rep(hint == ICE_HINT, "『氷採掘』のヒントは指定の1行だけ")
    for rn, line in REGION_LINES.items():
        bad = [w for w in FORBIDDEN_CARD + FORBIDDEN_MISSION_WORDS if w in line]
        rep(not bad, f"地域カード『{rn}』の1行に禁止語・ミッション語なし{('（検出：' + ','.join(bad) + '）') if bad else ''}")
        rep(not FORBIDDEN_PATTERN.search(line), f"地域カード『{rn}』の1行に対応文なし")

    # --- PDF本文
    def pt(fn):
        return pdf_text(os.path.join(OUT_DIR, fn))

    b3 = pt("B3_ミッションカード.pdf")
    t = b3.replace(ICE_HINT, "")
    bad = [w for w in FORBIDDEN_CARD if w in t]
    rep(not bad, f"B3 PDF本文に禁止語なし{('（検出：' + ','.join(bad) + '）') if bad else ''}")
    rep(not FORBIDDEN_PATTERN.search(t), "B3 PDF本文に「氷採掘は」等の対応文なし")
    rep(b3.count(ICE_HINT) == 1, "B3 PDFで氷のヒントが1回だけ出る")
    rep(all(m[0] in b3 for m in MISSIONS) and b3.count(MISSION_Q) == 4, "B3 PDFに4種のミッション名と問い（4回）")
    rep(not any(rn[:2] in b3 for rn in ["赤道", "火砕", "溶岩", "危機", "嵐の", "静かの"]), "B3 PDFに地域名なし")

    b4 = pt("B4a_地域カード.pdf")
    b4b = pt("B4b_地域カード_全球図.pdf")
    body = b4
    for n, la0, la1, lo0, lo1 in xr:
        rep(re.sub(r"\s+", "", n) in b4, f"B4a PDFに地域名『{n}』")
        rep(re.sub(r"\s+", "", n) in b4b, f"B4b PDFに地域名『{n}』")
        rs = "緯度" + rng(la0, la1)
        rep(rs in b4, f"B4a PDFに範囲『{rs}』（{n}）")
        rep(("経度" + lon_str(lo0, lo1)) in b4, f"B4a PDFに範囲『経度{lon_str(lo0, lo1)}』（{n}）")
        body = body.replace(re.sub(r"\s+", "", n), "")
    for w in FORBIDDEN_CARD + FORBIDDEN_MISSION_WORDS:
        rep(w not in body, f"B4a PDF本文（地域名を除く）に『{w}』なし")
    body_b = b4b
    for n, *_ in xr:
        body_b = body_b.replace(re.sub(r"\s+", "", n), "")
    for w in FORBIDDEN_CARD + FORBIDDEN_MISSION_WORDS:
        rep(w not in body_b, f"B4b PDF本文（地域名を除く）に『{w}』なし")

    # --- B5：付箋用の全球図に地域の範囲・地域名を載せない
    b5b = pt("B5b_付箋用月全球図.pdf")
    rep(not any(n[:3] in b5b for n, *_ in xr), "B5b（付箋用の全球図）に地域名なし（地域の範囲は載せない）")
    rep("表側" in b5b and "裏側" in b5b, "B5b に表側・裏側の区別がある")
    for lon in range(-180, 181, 10):
        if num(lon) not in b5b:
            rep(False, f"B5b に経度の目盛 {lon}")
            break
    else:
        rep(True, "B5b に経度の目盛（−180〜180、10°ごと）の数字がある")
    b5a = pt("B5a_黒板用記入表.pdf")
    rep(all(nm in b5a for _, nm, *_ in JIGSAW_SITES), "B5a に5地点の名前")
    rep(all(k in b5a for k in ["日較差", "地球の仰角", "太陽高度", "夜の最低温度", "出典"]), "B5a に4指標と出典の見出し")

    # --- B1：欄の突き合わせ
    b1 = pt("B1_個人用ワークシート.pdf")
    miss = [f for f in B1_FIELDS if re.sub(r"\s+", "", f) not in b1]
    rep(not miss, f"B1 に記録欄の名前がすべてある{('（ない：' + ','.join(miss) + '）') if miss else ''}")
    rep(b1.count("教員判定") >= 6 and all(k in b1 for k in ["教員（総括）", "知識・技能", "思考・判断・表現"]),
        "B1 に教員の判定チェック欄（知・思・態）")
    rep(all(re.sub(r"\s+", "", "「" + q + "」") in b1 for q in QUESTION_TYPES), "B1 に質問の型3つ")
    rep(re.sub(r"\s+", "", REASON_TYPE) in b1, "B1 に理由の型（2文＋弱点1つ）")
    rep(re.sub(r"\s+", "", K_CELSIUS) in b1, "B1 にKと℃の関係の1行")

    # --- B8
    b8 = pt("B8_班編成ペア確定シート_教員用.pdf")
    rep(all(re.sub(r"\s+", "", r) in b8 for r in RULES_B8), "B8 に規則（もとの3行＋過半数のときの規則）")
    rep(len(RULES_B8) <= 5 and all(any(k in r for r in RULES_B8) for k in ["同じ地域タイプ同士は組まない", "3班の輪", "第3時の冒頭"])
        and any("過半数" in r and "ミッションが違う" in r for r in RULES_B8), f"B8 の規則は{len(RULES_B8)}行（5行以内）で、もとの3規則と過半数の規則がある")
    rep("教員用" in b8, "B8 に「教員用」の表記")
    rep(len(B8_EXAMPLE) == 10, "B8 の記入例は10班")
    from collections import Counter
    mc = Counter(m for _, m, *_ in B8_EXAMPLE)
    rep(mc == {"氷採掘": 2, "電波天文台": 3, "太陽光発電": 3, "有人総合": 2}, f"B8 のミッション割当 {dict(mc)}")
    pair_ok = True
    d = {b: (t, p) for b, m, j, t, p, cr in B8_EXAMPLE}
    for b, (t, p) in d.items():
        if d[int(p)][1] != str(b) or d[int(p)][0] == t:
            pair_ok = False
    rep(pair_ok, "B8 の記入例：ペアが相互で、同じ地域タイプ同士がない")
    jc = Counter(j for _, m, j, t, p, cr in B8_EXAMPLE)
    rep(all(v == 2 for v in jc.values()) and set(jc) == set("ABCDE"), "B8 の記入例：ジグソーの地点は A〜E の5地点・各2班")
    rep(all(re.fullmatch(r"[①-⑧]", t) for _, m, j, t, p, cr in B8_EXAMPLE), "B8 の記入例：地域タイプは地域カードの番号（①〜⑧）")

    # --- B8：第2時の担当クレーター（AC8：ティコ・ラングレヌスの班は全班の30％以下）
    cc = Counter(cr for *_, cr in B8_EXAMPLE)
    side_n = sum(cc[k] for k in CRATER_SIDE)
    rep(set(cc) <= set(CRATER_NAMES) and CRATER_EXAMPLE_NAME not in cc, f"B8 の記入例：担当クレーターは4つのうちのどれか（教員の例のコペルニクスは担当にしない）{dict(cc)}")
    rep(side_n <= CRATER_SHARE_MAX * len(B8_EXAMPLE) + 1e-9, f"B8 の記入例：ティコ・ラングレヌスの班が全班の30％以下（{side_n}/{len(B8_EXAMPLE)}）")
    for n_, plan in B8_CRATER_PLAN.items():
        sn = sum(plan[k] for k in CRATER_SIDE)
        rep(sum(plan.values()) == n_ and sn <= CRATER_SHARE_MAX * n_ + 1e-9, f"B8 の班数別の目安（{n_}班）：合計が班数に一致し、ティコ・ラングレヌスが30％以下（{sn}/{n_}）")
    rep(all(k in b8 for k in CRATER_NAMES) and "第2時の担当クレーター" in b8 and "30％以下" in b8 and "ラングレヌスを割り当てない" in b8,
        "B8 に担当クレーターの列・割当の規則（30％以下・班が6以下のときの扱い）がある")
    rep(re.sub(r"\s+", "", B8_STAMP_NOTE) in b8 and "確認印の運用" in b8 and "抜き取り" in b8 and "未確認" in b8,
        "B8 p2 に確認印の運用の代案（班ごと・抜き取り、判断メモの回収は第2時の終わり。未確認の注記つき）")
    rep("教員用" in b8 and "映さない" in b8, "B8 の割当のページに「教員用（生徒には配らない・映さない）」の表記")

    # --- B9
    b9 = pt("B9_Web操作カード.pdf")
    rep("データ一覧" in b9 and "別のタブ" in b9, "B9 に「データ一覧は別のタブで開く」")
    rep("押さない" not in b9, "B9 に旧版の「データ一覧は押さない」が残っていない")
    rep("℃" in b9 and "表計算はK" in b9, "B9 に温度は℃・表計算はK")
    rep("「緯度・経度」欄" in b9 and "ピンを立てる" in b9, "B9 に緯度・経度の入力欄（ピンを立てる）")
    rep("地点名ヒント" in b9 and "全指標の表" in b9, "B9 に地点名ヒントのスイッチとピンカードの全指標の表")
    rep("プルダウン" in b9, "B9 に層の選択（左上凡例のプルダウン）")

    rep("ジグソー" in b9 and "A〜E" in b9 and "1日の温度差" in b9 and "24時間の差とは別" in b9, "B9 にジグソーの手順（A〜E・表の「1日の温度差」を読む）")
    rep("データのない地点（南極に近い点）" in b9 and "85°" not in b9, "B9 の傾斜の説明は実態に合う（85°以上、とは書かない）")
    rep("座標欄にそのピンの緯度・経度が入る" in b9 and "1100px" in b9, "B9 にピンの座標欄への反映と、幅約1100px未満の注")

    # --- S3：⑫の選択肢（検算の落とし穴）を紙に印字しない
    trap = ["傾斜は評価式に入っていない", "領域の箱は人が決めた", "極付近の温度カーブ", "表計算と位置がずれる", "日較差・夜の最低温度が大きくずれる"]
    for fn in ("B1_個人用ワークシート.pdf", "B2_班用ワークシート.pdf"):
        t = pt(fn)
        rep(not [w for w in trap if w in t], f"{fn}：⑫の選択肢の文を印字していない")
    rep("黒板に出た例から3つ選び" in b1, "B1 ⑫は枠3行＋「黒板に出た例から選ぶ」")

    # --- S4・S5：書き写しだけの欄を B2 に置かない（B2 に4バンド・密度の数値表がない）
    b2 = pt("B2_班用ワークシート.pdf")
    rep("日較差の平均[K]" not in b2 and "陸は海の何倍" not in b2 and "クレーターの数" not in b2, "B2 に、個人用B1③④へ写すだけの数値表（4バンド・密度）がない")
    rep("上位10の経度帯（言葉で）" not in b2 and "決めた1地点（付箋にも書く）" not in b2 and "選んだ地域名（ステップ4のC5" not in b2,
        "B2 に、個人用（⑦⑨）と重なる欄（言葉での変化・1地点・地域名）がない")
    rep(all(k in b1 for k in ["赤道帯", "中緯度帯", "高緯度帯", "極付近"]), "B1 ③に4バンドの日較差の欄（評価の記録欄は減らさない）")

    # --- S6：指標の並びと名前をWebのピンカードの表（1日の温度差・夜の最低温度・正午の太陽高度・地球の仰角）に合わせる
    def in_order(t, names):
        pos = [t.find(n) for n in names]
        return all(p >= 0 for p in pos) and pos == sorted(pos)
    names = ["1日の温度差", "夜の最低温度", "正午の太陽高度", "地球の仰角"]
    b5a_t = pt("B5a_黒板用記入表.pdf")
    rep(in_order(b5a_t, names) and "日較差" in b5a_t, "B5a の列の並びがWebのピンカードの表と同じ（1日の温度差・夜の最低温度・正午の太陽高度・地球の仰角。日較差を併記）")
    k = b2.find("指標（Webのピンカードの表の順）")
    rep(k >= 0 and in_order(b2[k:], ["1日の温度差（日較差）", "夜の最低温度", "正午の太陽高度", "地球の仰角"]), "B2 の記録表の並びがWebのピンカードの表と同じ")

    # --- S7：ジグソー地点は A〜E（地域カード①〜⑧と区別）
    rep([n for n, *_ in JIGSAW_SITES] == list("ABCDE"), "ジグソー地点の記号は A〜E")
    rep(not re.search(r"[①-⑤]", b5a_t), "B5a に丸数字（地域カードの番号と紛らわしい記号）がない")
    k1 = b2.find("ジグソーの5地点")
    rep(k1 >= 0 and not re.search(r"[①-⑤]", b2[k1:k1 + 160]), "B2 のジグソー地点の表が A〜E で、丸数字を使っていない")

    # --- S8：B5b の付箋の注（実寸から計算した値）
    b5b_t = pt("B5b_付箋用月全球図.pdf")
    rep("原寸" in b5b_t and "点シール" in b5b_t and "約25mm角" in b5b_t.replace(" ", ""), "B5b に付箋の大きさの注（原寸印刷・約25 mm角・点シール）")

    # =================================================================
    # 第2時の改訂（design_dai2ji_v2.md rev2。B1⑥・B2・B4c・B8・旧値の検査）
    # =================================================================
    # --- B1：⑥ 判断メモ（評価に入れない・画面の数字と理由・確認印）
    k6 = b1.find("⑥第2時")
    k6e = b1.find("B1個人用ワークシート（表）")
    seg6 = b1[k6:k6e] if k6 >= 0 and k6e > k6 else ""
    rep(re.sub(r"\s+", "", "⑥ " + L_MEMO) in b1 and "判断メモ" in b1 and "数字" in seg6, "B1 ⑥ が「判断メモ（画面の数字と理由を書く）」になっている")
    rep("数字を2つ" not in b1 and "数字2つ" not in b1, "B1 に旧い見出し「数字を2つ（入れる）」がない（欄の数と合わない言い方を残さない）")
    rep(all(w in seg6 for w in ["(a)線", "内側の最大", "使えなくなる月の面積", "(b)担当クレーター", "内部の夜の最低温度", "帯の9割が入る範囲",
                                 "内部は範囲の", "違う", "違わない", "教員確認印", "評価には入れない"]),
        "B1 ⑥ に (a)線・内側の最大・使えなくなる月の面積／(b)担当・内部・帯の9割が入る範囲・外内・違う違わない／確認印の枠／「評価には入れない」")
    rep(re.sub(r"\s+", "", L_LINE_REASON) in seg6 and re.sub(r"\s+", "", L_IO_REASON) in seg6,
        "B1 ⑥ に判断の理由を書く2つの欄（線の理由・内か外かの理由）がある")
    rep("(c)" not in seg6 and "自班の地域タイプ" not in seg6 and "第3時のシート" not in seg6 and "ステップ4" not in seg6 and "4b" not in seg6
        and "外側" not in seg6 and "内側" not in re.sub(r"ステップ1b.*?（B10）", "", seg6).replace("内側の最大", ""),
        "B1 ⑥ に旧(c)（自班の地域タイプと線の内外・第3時のシート）がない（地域タイプと線の内外の対応を紙に書かない）")
    rep("教員判定" not in seg6 and b1.count("教員判定") == 7, f"B1 ⑥ に教員判定欄を付けていない（教員判定は③④⑤⑧⑨⑩⑪の7つ。実際 {b1.count('教員判定')}）")
    rep("第2時のまとめ" not in b1 and "分かったこと" not in b1, "B1 に旧⑥「第2時のまとめ（分かったこと）」が残っていない")
    rep(all(re.sub(r"\s+", "", n) in b1 for n in [L_NICHI, L_DENSITY, L_SEARIKU, L_MEMO]), "B1 ③④⑤⑥の記録欄の名前がある（③④は不変）")
    # 画面の語と紙の語の一致（学習者版の実物）
    try:
        import openpyxl
        wbq = openpyxl.load_workbook(XLSX_PATH, data_only=True, read_only=True)
        w1b, w1c = wbq["ステップ1b_帯を刻む"], wbq["ステップ1c_地点と帯"]
        lab = {"1b A9": w1b["A9"].value, "1b A10": w1b["A10"].value, "1b A7": w1b["A7"].value, "1c D9": w1c["D9"].value,
               "1c E9": w1c["E9"].value, "1c A10": w1c["A10"].value, "1c A5": w1c["A5"].value}
        wbq.close()
        rep("21〜3時" in lab["1b A9"] and "使えなくなる月の面積" in lab["1b A10"] and "線" in lab["1b A7"] and "帯の9割が入る範囲" in lab["1c D9"]
            and "夜の最低温度" in lab["1c A10"] and "担当" in lab["1c A5"],
            "表計算の実物（1b：B9・B10・B7、1c：B10:E10・B5）の語が、紙の語（21〜3時・使えなくなる月の面積・線・帯の9割が入る範囲・夜の最低温度・担当）と一致")
    except Exception as e:                                           # noqa
        rep(False, f"表計算の実物の語の確認に失敗：{e}")

    # --- B2：第2時の欄（線・クレーター・根拠にした今日の数字）、旧い欄の削除
    rep(all(w in b2 for w in ["ステップ1b", "班の線", "刻むと見えたこと", "ステップ1c", "担当クレーター"]), "B2 に班の線・刻むと見えたこと・担当クレーターの欄")
    rep(b2.count("根拠にした今日の数字") >= 2 and "見直" not in b2, f"B2 の重みの案の表に「根拠にした今日の数字」の列（ステップ4・4bの2表。{b2.count('根拠にした今日の数字')}か所）。「見直す」の行はない")
    rep("極付近の値は信じてよい" not in b2 and b2.count("信じてよい") == 1, "B2 に旧い問い「極付近の値は信じてよい？」がない（残る「信じてよい？」は第3時の質問の型の1か所だけ）")
    rep(b1.count("信じてよい") == 1, "B1 の「信じてよい？」は第3時の質問の型（⑩）の1か所だけ")

    # --- B4c：クレーターカード
    b4c = pt("B4c_クレーターカード.pdf")
    cr_xlsx = load_craters_xlsx()
    rep(all(n in cr_xlsx for n in CRATER_NAMES) and CRATER_EXAMPLE_NAME in cr_xlsx, "B4c：担当4つ＋教員の例が、表計算『データ_地点比較』の地点名にある")
    rep(all(n in b4c for n in CRATER_NAMES) and CRATER_EXAMPLE_NAME not in b4c, "B4c にクレーター4つの名前（教員の例のコペルニクスは入れない）")
    rep(all(("緯度" + num(round(cr_xlsx[n][0], 2)) + "°") in b4c and ("経度" + num(round(cr_xlsx[n][1], 2)) + "°") in b4c
            and ("直径約" + num(round(cr_xlsx[n][2])) + "km") in b4c for n in CRATER_NAMES), "B4c の緯度・経度・直径が、『データ_地点比較』と一致")
    rep(all(w in b4c for w in ["ステップ1c", "B5", "⑥(b)", "夜の最低温度", "帯の9割が入る範囲", "同じ緯度の帯", "内部"]), "B4c に手順（ステップ1c・B5・⑥(b)）とことば（内部・同じ緯度の帯・帯の9割が入る範囲）")
    try:
        import openpyxl
        wbq = openpyxl.load_workbook(XLSX_PATH, data_only=True, read_only=True)
        rows_ = list(wbq[SITE_SHEET].iter_rows(values_only=True))
        wbq.close()
        ix = {h: i for i, h in enumerate(rows_[0])}
        leak = []
        for r in rows_[1:]:
            for col in ("inner_tmin", "band_tmin", "band_tmin_p5", "band_tmin_p95"):
                s_ = f"{round(float(r[ix[col]]), 1):.1f}"
                if s_ in b4c:
                    leak.append((r[0], col, s_))
        rep(not leak, f"B4c に、表計算の数字（内部・帯の夜の最低温度・9割の下端上端）が書かれていない{('（検出：' + str(leak) + '）') if leak else ''}")
    except Exception as e:                                           # noqa
        rep(False, f"B4c の数字の混入検査に失敗：{e}")

    # --- 新しい禁止語（設計 §7.1 の「出さない」語。答え・原因の語。紙の全PDF）
    NEW_BANNED = ["平らな地面", "斜面", "作り方", "本当の姿", "分かっていません", "データの限界", "信頼できない", "信頼できる", "岩が多い", "冷めにくい",
                  "極域のデータを信じ", "崩れ"]
    for fn in OUTPUTS:
        t_ = pt(fn)
        bad = [w for w in NEW_BANNED if w in t_]
        rep(not bad, f"{fn}：第2時の答え・原因の語なし{('（検出：' + ','.join(bad) + '）') if bad else ''}")
    # B1⑥・B2・B4c：判断の正解・判定の語
    for nm_, t_ in (("B1⑥", seg6), ("B2", b2), ("B4c", b4c)):
        bad = [w for w in ["判定", "若い", "新しい", "OK", "NG", "合格", "正解"] if w in t_]
        rep(not bad, f"{nm_}：判断の正解・判定の語なし{('（検出：' + ','.join(bad) + '）') if bad else ''}")

    # --- 旧い数字（295・280・236・161・差134 K）が、全PDFにない（AC6）
    for fn in OUTPUTS:
        t_ = pt(fn)
        old = re.findall(r"(?<![\d.])(?:295|280|236|161)(?![\d.])", t_) + re.findall(r"差は?約?134", t_)
        rep(not old, f"{fn}：旧い4バンドの値（295・280・236・161）・旧い差（134）なし{('（検出：' + ','.join(old) + '）') if old else ''}")
    rep("要注意" not in b1 and "要注意" not in b2, "B1・B2 に旧い呼び名「極付近（要注意）」がない")

    # --- ミッション語の走査（B3 と、教員用の B8 を除く全PDF）
    #   許容：ミッション名そのものが要る B3（ミッションカード）と、教員用の B8（班ごとのミッション割当を書く）。
    #   B1・B2 の「ミッション」欄は見出しの語だけ（ミッション名ではない）なので、走査の対象のままで引っかからない。
    SCAN_WORDS = ["電波", "氷採掘", "天文", "発電", "通信", "有人", "採掘", "太陽光"]
    for fn in OUTPUTS:
        if fn.startswith("B3_") or fn.startswith("B8_"):
            continue
        t = pt(fn)
        bad = [w for w in SCAN_WORDS if w in t]
        rep(not bad, f"{fn}：ミッション語の走査（電波・氷採掘・天文・発電・通信・有人・採掘・太陽光）{('（検出：' + ','.join(bad) + '）') if bad else ''}")

    # --- 全PDF：答え語
    for fn in OUTPUTS:
        t = pt(fn)
        bad = [w for w in ANSWER_PHRASES if w in t]
        rep(not bad, f"{fn}：答え語なし{('（検出：' + ','.join(bad) + '）') if bad else ''}")
        rep("sdomi" not in t and "学芸" not in t, f"{fn}：個人・学校を特定する語なし")

    for m in ok:
        print("OK  ", m)
    for m in ng:
        print("NG  ", m)
    print(f"\n検査：OK {len(ok)} 件／NG {len(ng)} 件")
    return not ng


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    if mode in ("all", "build"):
        build_all()
    if mode in ("all", "check"):
        sys.exit(0 if check_all() else 1)
