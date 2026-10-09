# -*- coding: utf-8 -*-
"""ガイド型（層2a）の Excel ブック course_moonbase.xlsx を組み立てる。

分析ステップ1〜5を、Excel の「見える数式」で行う。
- `course/build_course_data.py` が作った前処理済み CSV をシートに埋め込む
- 各ステップのシートには AVERAGEIFS / COUNTIFS / INDEX-MATCH などの数式と、
  グラフを置く（生徒は数値を書き換えて再計算する）
- ピボットテーブル・マクロ・Power Query は使わない（デスクトップ Excel の基本機能のみ）

    python course/build_course_data.py    # 先にこちら
    python course/build_course_xlsx.py    # 引数なし：教員版と学習者版を両方出す（どちらも openpyxl 出力。計算結果なし）
    python course/build_course_xlsx.py --recalc
                                          # ＋学習者版を Excel で再計算して保存した版にする（配付用・リポジトリ用）

出力:
  course/course_moonbase.xlsx                  教員版（原本。答え・要素の組合せ例を含む）
  course/student/course_moonbase_student.xlsx  学習者版（答えにあたる記述・セルを除いたもの）

配付用の学習者版（計算結果を保存した版）の作り方:
  openpyxl は計算結果を保存できないため、Excel の「保護ビュー」（メール添付・ダウンロード直後）では
  値が空欄に見える。そこで、配付・リポジトリ用の学習者版は次の手順で作る（Windows＋Excel が必要）。
    1. python course/build_course_data.py regions      # 地域名の表示名を反映（通常は不要。済み）
    2. python course/build_course_xlsx.py --recalc     # 教員版と学習者版を出し、学習者版を再計算版に置き換える
    3. python course/check_course_xlsx.py --all        # 答え語・要素名の語・個人情報・構造・Excel 実機の検査
  再計算は course/recalc_xlsx_with_excel.py（Excel を COM で動かし、全再計算→別名保存→
  利用者名 lastModifiedBy・保存先パス absPath・作成者を除去）。Excel で直接上書き保存すると
  利用者名と保存先パスが xlsx に入るので、リポジトリに置く版では行わない。
  教員版は openpyxl 版のまま（教員が開いて使う。開くと再計算される）。

第2時の再設計（docs/design_dai2ji_v2.md rev2。ステップ1の改修）:
  ・新シート ステップ1b_帯を刻む（帯の幅 B6・線 B7。棒グラフ2つ）、ステップ1c_地点と帯（クレーター B5。折れ線1つ）
  ・新データ表 データ_緯度行（90行）・データ_地点比較（5行）。学習者版でも見える。作り方は build_course_data.py
  ・ステップ1の4バンドは データ_緯度行 から求める（295・279・234・156 K）。旧B（B17:B18・24時間カーブ）・旧C は残す

v4.0 の改修（要件定義 A1〜A4）:
  A1 学習者版を同じスクリプトから生成（答えの除去・非表示・案内の書き換え）
  A2 入力検証（C5 のプルダウン、重みは 0〜5 の整数、状態表示セル）
  A3 上位10の同点重複の解消（教員版・学習者版とも）
  A4 上位10の範囲（緯度・経度／4b は日照率・傾斜・永久影までの距離）の自動表示

要素ごとの分析と重み（docs/design_yoso_v3.md rev2）の改修:
  ・ステップ4の C5 のプルダウンの先頭に「月全体（線の内側）」を足した（非表示列 M5:M13 の一覧。名前 AreaChoices・AllMoon・LineOK）。
    選ぶと、『緯度の絶対値が線以下のマス』（マス中心）だけを採点する。線は新しい黄色セル C4（学習者版は空欄）。
    正規化は データ_環境 J〜N（全球固定）のまま。O 列の『採点する行』の条件だけが変わる。
  ・学習者版の既定の重みは C8（電力＝太陽高度の行）＝1（ステップ4・4b）。上位10の表に『夜の最低温度』の列（H）を足した。
  ・旧い呼び名の『ミッション』は『要素』に置き換えた（教員版の組合せ例は月全体の例・4bの例。数値は _top_all / _top_4b で再計算）。

同点重複の原因（A3）:
  上位10の表は、順位 k のスコアを LARGE で、その行を MATCH(…,0) で引いていた。
  MATCH は同じ値が複数あると「最初の行」しか返さない。スコアは 0〜1 に直した指標
  （前処理で小数4桁に丸め済み）の重みつき平均なので、1つの指標だけに重みを置くと
  同点の行が大量に出る（太陽高度は緯度だけで決まる、地球の仰角は経度の対称な2点が
  同値、など）。同点の行は LARGE が同じ値を返し続け、MATCH が毎回同じ最初の行を返す
  ため、同じ地点が何度も並んだ。
  （同時に見つけた不具合：C5 が空欄だと、地域名が空欄の行〔データ_環境 の region が空の 13604 行〕が
  「C5 と一致」と判定されて採点されていた。C5 が空欄のときは採点しないようにした。）
  直し方：スコア列の隣に「順位用」列を足し、ROUND(スコア,10) から ROW()×1E-12 を
  引いて、同点を行番号の小さい順に解く（差は最大でも約1.5E-8。ふつうに起こる最小の
  スコア差 約4E-6 より十分小さいので、同点でない行の順位は変わらない）。表示の
  スコアは元のスコア列を ROUND(…,3) で読むので、見た目は変わらない。
"""
import argparse
import math
import pathlib
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
DATA = HERE / "data"
OUT_TEACHER = HERE / "course_moonbase.xlsx"
OUT_STUDENT = HERE / "student" / "course_moonbase_student.xlsx"

NAVY = "1A2F4B"
H1 = Font(name="Yu Gothic", size=15, bold=True, color="FFFFFF")
H2 = Font(name="Yu Gothic", size=12, bold=True, color=NAVY)
BODY = Font(name="Yu Gothic", size=10)
BOLD = Font(name="Yu Gothic", size=10, bold=True)
BLUE = Font(name="Yu Gothic", size=10, bold=True, color="1A6FB0")
HFILL = PatternFill("solid", fgColor=NAVY)
INFILL = PatternFill("solid", fgColor="FFF6E9")   # 生徒が入力するセル
WRAP = Alignment(wrap_text=True, vertical="top")

TIE_EPS = "0.000000000001"   # 1E-12。順位用列で ROW() に掛ける（A3）
ALL_MOON = "月全体（線の内側）"   # ステップ4 C5 のプルダウンの先頭（非表示列 M5。名前 AllMoon）
TEACHER_LINE = 70             # 教員版のステップ4 C4（線）の既定。学習者版は空欄
NORM5 = ["norm_sun_high", "norm_amp_low", "norm_earth_high", "norm_earth_low", "norm_night_warm"]
NORM4B = ["norm_illum", "norm_near_shadow", "norm_low_psf", "norm_low_slope"]
REG_FAR_EQ = "裏側・赤道（月の裏側の赤道帯）"   # 3コマ版の表示名（build_course_data.py の DISPLAY_NAMES と同じ）
N_REGIONS = 8                 # データ_地域 の行数（名前のプルダウンの範囲）

# 第2時（改訂設計 rev2）の新シート
S1B = "ステップ1b_帯を刻む"
S1C = "ステップ1c_地点と帯"
LR_HEADERS = ["lat_lo　緯度の絶対値の下 [°]", "lat_hi　緯度の絶対値の上 [°]", "n_cells　セル数（0.5°）",
              "swing_K　1日の温度差（最高−最低）の平均 [K]",
              "night_peak_pct　最高が21〜3時に出る地点の割合 [%]"]

# 学習者版で非表示にするシート（名前は変えない。要件 §3.2）
STUDENT_HIDDEN = ["ステップ3_月全体", "ステップ4b_南極", "ステップ5_まとめ",
                  "データ_北極", "データ_地質", "データ_クレーター年代"]


def _title(ws, text, row=1):
    ws.cell(row=row, column=1, value=text).font = H1
    ws.cell(row=row, column=1).fill = HFILL
    for c in range(2, 12):
        ws.cell(row=row, column=c).fill = HFILL
    ws.row_dimensions[row].height = 22


def _h2(ws, text, row):
    ws.cell(row=row, column=1, value=text).font = H2


def _note(ws, text, row, col=1, span=10):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = BODY
    cell.alignment = WRAP
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    ws.row_dimensions[row].height = 14 * (1 + text.count("\n"))


def _input(ws, cell, value):
    ws[cell] = value
    ws[cell].fill = INFILL
    ws[cell].font = BOLD


def _text_width(s):
    """表示幅の目安（全角=1.9、半角=1.0 の列幅単位）"""
    return sum(1.9 if ord(ch) > 0x2E80 else 1.0 for ch in s)


def _fit_notes(ws, slack=0.92):
    """結合セルの注（_note）の行の高さを、折り返しに必要な行数で決め直す。
    列幅を決めた後に呼ぶ。もとの高さより低くはしない。"""
    for rng in ws.merged_cells.ranges:
        if rng.min_row != rng.max_row:
            continue
        cell = ws.cell(row=rng.min_row, column=rng.min_col)
        if not isinstance(cell.value, str) or cell.value.startswith("="):
            continue
        total = 0.0
        for c in range(rng.min_col, rng.max_col + 1):
            w = ws.column_dimensions[get_column_letter(c)].width
            total += w if w else 8.43
        lines = 0
        for part in cell.value.split("\n"):
            lines += max(1, math.ceil(_text_width(part) / (total * slack)))
        h = 14 * lines + 2
        cur = ws.row_dimensions[rng.min_row].height or 0
        if h > cur:
            ws.row_dimensions[rng.min_row].height = h


def _first_rank(score, k=1):
    """順位用列（ROUND(スコア,10)−ROW()×1E-12）で上位 k 行の位置を返す（同点は行番号の小さい順）。score は NaN＝採点しない行"""
    import numpy as np
    s = np.asarray(score, dtype=float)
    rows = np.arange(2, len(s) + 2)
    p = np.where(np.isnan(s), np.nan, np.round(s, 10) - rows * float(TIE_EPS))
    ok = np.where(~np.isnan(p))[0]
    return list(ok[np.argsort(-p[ok], kind="stable")][:k])


def _top_all(env_df, line, w, k=1):
    """ステップ4『月全体（線の内側）』の上位 k 行 [(緯度, 経度)…]（教員版の例と検査用。O 列の式と同じ）"""
    import numpy as np
    m = (env_df["lat"].abs() <= line).values
    s = sum(wi * env_df[c].values for wi, c in zip(w, NORM5)) / max(1, sum(w))
    idx = _first_rank(np.where(m, s, np.nan), k)
    return [(float(env_df["lat"].values[i]), float(env_df["lon"].values[i])) for i in idx]


def _top_4b(ps_df, w, k=1):
    """ステップ4b の上位 k 行 [(緯度, 経度)…]"""
    s = sum(wi * ps_df[c].values for wi, c in zip(w, NORM4B)) / max(1, sum(w))
    idx = _first_rank(s, k)
    return [(float(ps_df["lat"].values[i]), float(ps_df["lon"].values[i])) for i in idx]


def _fmt_pt(p):
    return f"（{p[0]:.1f}, {p[1]:.1f}）"


def add_data_sheet(wb, csv_name, sheet_name, student=False):
    df = pd.read_csv(DATA / csv_name)
    if student and sheet_name == "データ_地域":
        df = df.iloc[:, :5]          # 名前と範囲（A〜E列）だけ。rationale・caveat は答えを含む
    if student and sheet_name == "参考":
        # 説明列から答えの言い回し（「＝陸のほうが古い」）を除く
        df["説明"] = df["説明"].str.replace("（＝陸のほうが古い）", "", regex=False)
    ws = wb.create_sheet(sheet_name)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    for c in range(1, len(df.columns) + 1):
        ws.cell(row=1, column=c).font = BOLD
        ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor="EEF1F6")
    ws.freeze_panes = "A2"
    ws.sheet_state = "visible"
    return ws, df


def _weight_validation(ws, cell_range):
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="5",
                        allow_blank=False, showErrorMessage=True, showInputMessage=True,
                        errorStyle="stop",
                        errorTitle="重みの入れかた",
                        error="重みは 0〜5 の整数で入れてください（0＝その指標は気にしない）。",
                        promptTitle="重み", prompt="0〜5 の整数（0＝気にしない）")
    ws.add_data_validation(dv)
    dv.add(cell_range)


def _status_format(ws, cell):
    """状態表示セルの色分け（【注意】＝赤、OK＝緑、南極の誘導＝橙）"""
    red = PatternFill("solid", bgColor="FADBD8", fgColor="FADBD8")
    green = PatternFill("solid", bgColor="E3F4E8", fgColor="E3F4E8")
    amber = PatternFill("solid", bgColor="FDEBD0", fgColor="FDEBD0")
    ws.conditional_formatting.add(cell, FormulaRule(
        formula=[f'LEFT({cell},2)="【注"'], fill=red, font=Font(color="A93226", bold=True)))
    ws.conditional_formatting.add(cell, FormulaRule(
        formula=[f'LEFT({cell},2)="OK"'], fill=green, font=Font(color="1E7B3A", bold=True)))
    ws.conditional_formatting.add(cell, FormulaRule(
        formula=[f'LEFT({cell},2)="南極"'], fill=amber, font=Font(color="9C5700", bold=True)))


# ----- 第2時（改訂設計 rev2）の新シート：ステップ1b・ステップ1c --------------------------------------
GREY_FILL = PatternFill("solid", bgColor="E4E4E4", fgColor="E4E4E4")
TEACHER_RED = Font(name="Yu Gothic", size=10, bold=True, color="A93226")
HEAD_FILL = PatternFill("solid", fgColor="EEF1F6")


def _line_expect(lr_df, line):
    """線（|緯度|の上限）の内側で、最高が21〜3時に来る割合の最大 [%] と、使えなくなる月の面積 [%]（式は B9・B10 と同じ）"""
    inside = lr_df[lr_df["lat_hi"] <= line]
    mx = float(inside["night_peak_pct"].max())
    area = (1 - math.sin(math.radians(line))) * 100
    return round(mx, 1), round(area, 1)


def _memo(ws, lines, row, span=5):
    """教員用メモ（教員版だけ）。1行ずつ結合セルで書く"""
    ws.cell(row=row, column=1, value="【教員用メモ】生徒には配らない（学習者版には出ない）").font = TEACHER_RED
    row += 1
    for t in lines:
        _note(ws, t, row, span=span)
        row += 1
    _fit_notes(ws, 0.62)       # 列幅が決まった後で、メモの行の高さを折り返しに合わせる
    return row


def _tidy_axes(chart, y_fmt):
    """軸タイトル・凡例が目盛や軸に重ならないようにし、目盛の書式（整数）と薄い補助線にそろえる"""
    for ax in (chart.x_axis, chart.y_axis):
        if ax.title is not None:
            ax.title.overlay = False
    if chart.legend is not None:
        chart.legend.overlay = False
    chart.y_axis.numFmt = NumFmt(formatCode=y_fmt, sourceLinked=False)
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="D9D9D9", w=9525))


def _head_cells(ws, row, col0, heads, height):
    for i, h in enumerate(heads):
        cell = ws.cell(row=row, column=col0 + i, value=h)
        cell.font = BOLD
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = height


def build_step1b(wb, student, lr_df, LR):
    """第2時 判断①：緯度の帯の幅（30・5・1°）を黄色セルで変えて、日較差と『最高が21〜3時に来る割合』の棒グラフを描き直す。
    もう1つの黄色セル＝線（使う緯度の上限）。画面に出す数字は2つだけ（線の内側の最大・使えなくなる月の面積）。合否は出さない。"""
    ws = wb.create_sheet(S1B)
    _title(ws, "ステップ1b：緯度の帯を刻んで、使う緯度の上限（線）を引く")
    _note(ws, "『データ_緯度行』は、月全体の0.5度のセル（259,200個）を、緯度の絶対値1度ごとの90行にまとめた表。"
              "各行に、1日の温度差（最高−最低）の平均と、1日のうち最高の温度が出る時刻が現地時間の21〜3時になるセルの割合［％］が入っている。"
              "月には大気がなく、地面をあたためるのは太陽だけ。"
              "黄色いセルで帯の幅を変えると、下の表と2つのグラフが、帯ごとの平均で描き直される（各行には、その行が入る帯の平均が入る）。",
          3, span=5)
    _h2(ws, "A. 帯の幅と線（黄色いセルを変える）", 5)
    ws["A6"] = "帯の幅 [°]（30・5・1 から選ぶ）"
    ws["A7"] = "使う緯度の上限（線）[°]（50〜90 の整数）"
    _note(ws, "B9 は、『データ_緯度行』の1°ごとの行の値から求める。帯の幅を 30 にするとグラフは30°ぶんの平均になって棒が小さく見え、"
              "幅を 1 にすると、グラフの棒が B9 と同じ1°ごとの値になる。", 8, span=5)
    ws["A9"] = "線の内側で、最高が21〜3時に来る割合の最大 [%]"
    ws["A10"] = "使えなくなる月の面積 [%]"
    for r in (6, 7, 9, 10):
        ws[f"A{r}"].font = BODY
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="center")
    _input(ws, "B6", 30)
    _input(ws, "B7", 90)
    dv = DataValidation(type="list", formula1='"30,5,1"', allow_blank=False, showErrorMessage=True,
                        showInputMessage=True, errorStyle="stop",
                        errorTitle="帯の幅", error="帯の幅は 30・5・1 のどれかを、一覧から選んでください。",
                        promptTitle="帯の幅 [°]", prompt="30・5・1 から選ぶ")
    ws.add_data_validation(dv)
    dv.add("B6")
    dv2 = DataValidation(type="whole", operator="between", formula1="50", formula2="90",
                         allow_blank=False, showErrorMessage=True, showInputMessage=True, errorStyle="stop",
                         errorTitle="線の入れかた", error="線は 50〜90 の整数（緯度の絶対値 [°]）で入れてください。",
                         promptTitle="線 [°]", prompt="50〜90 の整数。この緯度の絶対値までを使う")
    ws.add_data_validation(dv2)
    dv2.add("B7")
    ok_line = "AND(ISNUMBER($B$7),$B$7>=50,$B$7<=90,$B$7=INT($B$7))"
    ws["B9"] = f'=IF({ok_line},ROUND(SUMPRODUCT(MAX(({LR("B")}<=$B$7)*{LR("E")})),1),"")'
    ws["B10"] = f'=IF({ok_line},ROUND((1-SIN(RADIANS($B$7)))*100,1),"")'
    for c in ("B9", "B10"):
        ws[c].font = Font(name="Yu Gothic", size=12, bold=True, color="1A6FB0")
        ws[c].number_format = "0.0"
    ws["A11"] = (
        '=IFERROR(IF(AND($B$6<>30,$B$6<>5,$B$6<>1),"【注意】帯の幅は 30・5・1 のどれかにしてください",'
        f'IF(NOT({ok_line}),"【注意】線は 50〜90 の整数で入れてください",'
        '"OK　帯の幅 "&$B$6&"°　線 "&$B$7&"°")),"【注意】黄色いセルに数字以外が入っています")')
    ws["A11"].font = BOLD
    _status_format(ws, "A11")
    _note(ws, "気づき：帯の幅を 30 → 5 → 1 と変えると、2つのグラフはどう変わる？　"
              "線（使う緯度の上限）を何度に引く？　B9 と B10 の2つの数字を紙に書く。"
              "表の灰色の行は、線の外側。", 12, span=5)

    # 表：90行（緯度の絶対値の下を1度ずつ）。各行に、その行が入る帯の平均
    H0 = 14
    _head_cells(ws, H0, 1, ["緯度の絶対値の下 [°]", "1日の温度差の平均 [K]（帯の平均）",
                            "最高が21〜3時に出る地点の割合 [%]（帯の平均）", "帯の下 [°]", "帯の上 [°]"], 44)
    r0, r1 = H0 + 1, H0 + 90
    for i in range(90):
        r = r0 + i
        ws[f"A{r}"] = f"=データ_緯度行!A{i + 2}"
        ws[f"D{r}"] = f"=INT(A{r}/$B$6)*$B$6"
        ws[f"E{r}"] = f"=D{r}+$B$6"
        ws[f"B{r}"] = f'=AVERAGEIFS({LR("D")},{LR("A")},">="&D{r},{LR("B")},"<="&E{r})'
        ws[f"C{r}"] = f'=AVERAGEIFS({LR("E")},{LR("A")},">="&D{r},{LR("B")},"<="&E{r})'
        ws[f"B{r}"].number_format = "0.0"
        ws[f"C{r}"].number_format = "0.0"
        for c in "ABCDE":
            ws[f"{c}{r}"].font = BLUE if c in "BC" else BODY
    ws.conditional_formatting.add(
        f"A{r0}:E{r1}", FormulaRule(formula=[f"AND(ISNUMBER($B$7),$A{r0}+1>$B$7)"], fill=GREY_FILL,
                                    font=Font(color="888888")))

    def bar_chart(title, ycol, ytitle, ymax, major, color, anchor):
        bar = BarChart()
        bar.type = "col"
        bar.title = title
        bar.y_axis.title = ytitle
        bar.x_axis.title = "緯度の絶対値 [度]（0 が赤道、90 が極）"
        bar.add_data(Reference(ws, min_col=ycol, min_row=H0, max_row=r1), titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=1, min_row=r0, max_row=r1))
        bar.height, bar.width = 7.2, 16
        bar.x_axis.delete = False
        bar.y_axis.delete = False
        bar.legend = None
        bar.title.overlay = False
        bar.varyColors = False
        bar.gapWidth = 0
        bar.y_axis.scaling.min = 0
        bar.y_axis.scaling.max = ymax
        bar.y_axis.majorUnit = major
        bar.x_axis.tickLblSkip = 10
        bar.x_axis.tickMarkSkip = 10
        bar.series[0].graphicalProperties.solidFill = color
        bar.series[0].graphicalProperties.line.solidFill = color
        _tidy_axes(bar, "0")
        ws.add_chart(bar, anchor)

    # 追加の順（日較差→割合）は変えず、置く位置だけ入れ替える：B6・B7 を変えながら見る「割合」を最初の画面（G5）に、日較差を下（G15）に
    bar_chart("1日の温度差（最高−最低）の平均 [K]", 2, "温度差 [K]", 300, 50, "1A6FB0", "G15")
    bar_chart("最高が21〜3時に出る地点の割合 [%]", 3, "割合 [%]", 8, 2, "D9622B", "G5")

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 3
    for r in (6, 7, 9, 10):
        ws.row_dimensions[r].height = 30
    _fit_notes(ws, 0.72)

    if not student:
        ws["A13"] = f"【教員用】想定値と問い返しの基準は {r1 + 2} 行目から。"
        ws["A13"].font = TEACHER_RED
        rows = lr_df.set_index("lat_lo")
        seg = "、".join(f"{lo}〜{lo + 1}°＝{rows.loc[lo, 'night_peak_pct']:.1f}％（日較差 {rows.loc[lo, 'swing_K']:.1f} K）"
                        for lo in range(84, 90))
        lines = []
        for ln in (70, 80, 85, 86, 87, 88):
            mx, ar = _line_expect(lr_df, ln)
            lines.append(f"線{ln}° → 内側の最大 {mx}％、使えなくなる面積 {ar}％")
        memo = [
            "最高が21〜3時に来る割合（|緯度|の行ごと。北南を合わせる）：" + seg + "。"
            "60°未満〜84°の行は 0.0％。帯の幅を 30 → 5 → 1 と刻むと、帯の最大が 0.6％ → 3.4％ → 6.8％ と見えてくる。",
            "線を入れたときの想定値（内側＝行の上端が線以下）：" + "；".join(lines) + "。",
            "問い返しの基準：線が60°以下なら『60〜84°の行は割合が 0.0％。なぜ捨てる？　線70°だと月の何％が使えなくなる？（6.0％）』。"
            "線が88°以上なら『86〜87°で 6.8％、87〜88°で 4.2％ある。内側に入れて大丈夫？』。"
            "85〜87°の線は『1.0％と 6.8％の間で、どこから使わない？』と数字を指す。許容は決めない。",
            "窓（21〜3時）の根拠：自転軸の傾き（約1.5°。一般に知られた値で、同梱データでは未確認）を考えても、21〜3時は"
            "|緯度|約87.8°までなら夏でも太陽が地平線の下にある時間帯。88°以上の行は割合の意味が変わる。89°行で割合が下がるのは回復ではない（日較差も下げ止まっている）。",
            "教員の一言の要点：夜に最高が出るのは、平らな地面の考え方では説明できない。斜面の向きかもしれないし、データの作り方かもしれない。"
            "今日はどちらかを調べない。『緯度だけで場所を代表させる見方』をどこまで使うかを決める。",
            "旧デモ（ステップ1の B17・B18 に -86.75・0.25 を入れる）で24時間カーブを見せたら、0.25 に戻す。",
        ]
        _memo(ws, memo, r1 + 2)
    return ws


def build_step1c(wb, student, sc_df):
    """第2時 判断②：担当クレーターの内部と、同じ緯度の帯（月を1周）の24時間カーブ・夜の最低温度を比べる。判定は自動表示しない。"""
    ws = wb.create_sheet(S1C)
    _title(ws, "ステップ1c：クレーターの内部と、同じ緯度の帯を比べる")
    _note(ws, "『データ_地点比較』には、5つのクレーターについて、クレーターの内部と、同じ緯度の帯の、24時間の温度カーブ（セルの平均）と夜の最低温度が入っている。"
              "内部＝クレーターの中心を囲む四角の中にある0.5度のセル（四角は、中心から南北・東西に、半径の半分の長さ＋0.25度の範囲。東西の長さは緯度に合わせて補正）。"
              "帯＝クレーターと同じ緯度（±1.5度）を、月を1周したもの（クレーターから経度が直径1つ分以上はなれたセルだけ）。"
              "担当のクレーターを黄色いセル B5 の一覧から選ぶ。", 3, span=5)
    ws["A5"] = "担当のクレーター（B5 の一覧から選ぶ）"
    ws["A5"].font = BODY
    ws["A5"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[5].height = 30
    _input(ws, "B5", None if student else "コペルニクス")
    dv = DataValidation(type="list", formula1="SiteNames", allow_blank=True, showErrorMessage=True,
                        showInputMessage=True, errorStyle="stop",
                        errorTitle="クレーター名が違います", error="『データ_地点比較』にある名前を、一覧から選んでください。",
                        promptTitle="担当のクレーター", prompt="クリックして、一覧から1つ選ぶ")
    ws.add_data_validation(dv)
    dv.add("B5")
    M = 'MATCH($B$5,データ_地点比較!$A$2:$A$6,0)'
    ws["A6"] = (
        '=IFERROR(IF($B$5="","担当のクレーターを選んでください（B5 をクリックすると一覧が出ます）",'
        'IF(COUNTIF(SiteNames,$B$5)=0,"【注意】B5 の名前が『データ_地点比較』にありません。一覧から選んでください",'
        f'"OK　"&$B$5&"：内部 "&INDEX(データ_地点比較!$E$2:$E$6,{M})&" セル、帯 "&INDEX(データ_地点比較!$F$2:$F$6,{M})&" セル")),'
        '"【注意】B5 の名前を確かめてください")')
    ws["A6"].font = BOLD
    _status_format(ws, "A6")

    _h2(ws, "A. 夜の最低温度 [K]（セルごとの1日の最低温度の平均）", 8)
    _head_cells(ws, 9, 2, ["クレーターの内部", "同じ緯度の帯の平均", "帯の9割が入る範囲：下端", "帯の9割が入る範囲：上端"], 44)
    ws["A10"] = "夜の最低温度 [K]"
    ws["A10"].font = BOLD
    for c, col in zip("BCDE", ("BC", "BD", "BE", "BF")):
        ws[f"{c}10"] = (f'=IF($B$5="","",IFERROR(ROUND(INDEX(データ_地点比較!${col}$2:${col}$6,{M}),1),""))')
        ws[f"{c}10"].font = Font(name="Yu Gothic", size=12, bold=True, color="1A6FB0")
        ws[f"{c}10"].number_format = "0.0"
    _note(ws, "比べ方：クレーターの内部の夜の最低温度が、『帯の9割が入る範囲』（帯の中のセルの9割が、この下端と上端の間に入る）の内側か外側かを見る。"
              "紙には、内部の値と、帯の9割が入る範囲の2つの数字を書く。", 11, span=5)

    _h2(ws, "B. 24時間の温度カーブ（現地時間。セルの平均）", 13)
    _head_cells(ws, 14, 1, ["現地時間", "クレーター内部の平均 [K]", "同じ緯度の帯の平均 [K]"], 30)
    for h in range(24):
        r = 15 + h
        ws[f"A{r}"] = h
        ws[f"A{r}"].font = BODY
        for c, base in (("B", 7), ("C", 31)):     # inner_t_lt00 は G 列（7）、band_t_lt00 は AE 列（31）
            col = get_column_letter(base + h)
            ws[f"{c}{r}"] = (f'=IF($B$5="",NA(),IFERROR(INDEX(データ_地点比較!${col}$2:${col}$6,{M}),NA()))')
            ws[f"{c}{r}"].number_format = "0.0"
            ws[f"{c}{r}"].font = BLUE
    ws.conditional_formatting.add("B15:C38", FormulaRule(formula=["ISNA(B15)"], font=Font(color="FFFFFF")))
    line = LineChart()
    line.title = "1日の温度カーブ：クレーター内部と、同じ緯度の帯"
    line.y_axis.title = "温度 [K]"
    line.x_axis.title = "現地時間"
    line.add_data(Reference(ws, min_col=2, max_col=3, min_row=14, max_row=38), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=15, max_row=38))
    line.height, line.width = 8.5, 16
    line.x_axis.delete = False
    line.y_axis.delete = False
    line.legend.position = "b"
    line.title.overlay = False
    line.varyColors = False
    for sr, color, wd in zip(line.series, ("D9622B", "1A6FB0"), (28000, 22000)):
        sr.smooth = False
        sr.graphicalProperties.line.solidFill = color
        sr.graphicalProperties.line.width = wd
        sr.marker.symbol = "none"
    _tidy_axes(line, "0")
    ws.add_chart(line, "E13")
    _note(ws, "注：A の夜の最低温度は、セルごとの最低温度の平均。B のカーブはセルの平均なので、カーブの最低は A と少しちがう。"
              "朝と夕方の温度が急に変わる時間帯（6〜7時・17〜18時ごろ）は、同じ緯度の帯の中でも場所による差が大きい。夜と昼の平らな部分を見る。", 40, span=5)

    ws.column_dimensions["A"].width = 34
    for c in "BCDE":
        ws.column_dimensions[c].width = 22
    _fit_notes(ws, 0.72)

    if not student:
        sc = sc_df.set_index("site_name")
        rows = []
        for nm in sc.index:
            s_ = sc.loc[nm]
            d_ = s_["inner_tmin"] - s_["band_tmin"]
            rows.append(f"{nm}：内部 {s_['inner_tmin']:.1f}、帯の平均 {s_['band_tmin']:.1f}、差 {d_:+.1f}、"
                        f"帯の9割 {s_['band_tmin_p5']:.1f}〜{s_['band_tmin_p95']:.1f}（内部−上端 {s_['inner_tmin'] - s_['band_tmin_p95']:+.1f}）"
                        f"［内部 {int(s_['n_inner'])} セル・帯 {int(s_['n_band'])} セル］")
        memo = [
            "想定値（夜の最低温度 [K]）。" + "　".join(rows),
            "全体の比率（教員が別に数えた値。この表計算では再計算しない）：直径60〜200 km・|緯度|55°未満のクレーター778個のうち、帯との差が＋3 K以上は12個（約1.5％）。"
            "教員が選んだ5つのうち2つはその上位に入る。『クレーターは違うもの』という規則ではなく、例外を探す練習として扱う。",
            "教員の1例＝コペルニクス（担当にしない）。担当は4クレーター。違う側（ティコ・ラングレヌス）を担当する班は全班の30％以下にする。",
            "ラングレヌス班の見取り：『帯の9割が入る範囲の上端との差が小さい。何Kなら違うと言う？』と問い返す。",
            "ステップ1の旧デモ（B17・B18 に 9.25・-20.75）はコペルニクス内部の1セルで、夜の最低は 102.8 K。内部平均 103.4 K と少し違う（3度標本の1セル）。"
            "『データ_着陸地点』のコペルニクス・ティコは1点の値で、ここの内部平均とは別の数字。",
            "朝夕（6〜7時・17〜18時）の差は原因を調べていない。『夜と昼の平らな部分を見よう』とだけ言う。",
        ]
        _memo(ws, memo, 43)
    return ws


def build(student: bool = False):
    out = OUT_STUDENT if student else OUT_TEACHER
    out.parent.mkdir(exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.creator = None        # 作成者は空（個人名を入れない）
    wb.properties.title = ("月データでムーンベースの場所を決めよう（学習者版）" if student
                           else "月データでムーンベースの場所を決めよう（教員用・答え入り）")
    wb.calculation = CalcProperties(fullCalcOnLoad=True)

    # ----- データシート（埋め込み） -----
    _, t_df = add_data_sheet(wb, "temp_grid.csv", "データ_温度")
    _, c_df = add_data_sheet(wb, "craters_labeled.csv", "データ_クレーター")
    _, a_df = add_data_sheet(wb, "crater_ages_labeled.csv", "データ_クレーター年代")
    ps_ws, ps_df = add_data_sheet(wb, "polar_south_sites.csv", "データ_南極")
    add_data_sheet(wb, "polar_north_sites.csv", "データ_北極")
    lr_ws, lr_df = add_data_sheet(wb, "temp_lat_rows.csv", "データ_緯度行")      # 第2時（ステップ1・1b）。学習者版でも見える
    _, sc_df = add_data_sheet(wb, "site_compare.csv", "データ_地点比較")        # 第2時（ステップ1c）。学習者版でも見える
    assert len(lr_df) == 90 and len(sc_df) == 5
    for col, head in enumerate(LR_HEADERS, start=1):     # 見出しに単位・窓を書く（判断や原因の語は入れない）
        lr_ws.cell(row=1, column=col, value=head)
    for col, w in zip("ABCDE", (22, 22, 12, 20, 40)):
        lr_ws.column_dimensions[col].width = w
    _, g_df = add_data_sheet(wb, "geology_grid.csv", "データ_地質")
    env_ws, env_df = add_data_sheet(wb, "env_grid.csv", "データ_環境")
    _, reg_df = add_data_sheet(wb, "candidate_regions.csv", "データ_地域", student)
    _, ls_df = add_data_sheet(wb, "landing_sites.csv", "データ_着陸地点")
    _, ref_df = add_data_sheet(wb, "reference.csv", "参考", student)

    T_N = len(t_df) + 1
    C_N = len(c_df) + 1
    A_N = len(a_df) + 1
    PS_N = len(ps_df) + 1
    G_N = len(g_df) + 1
    E_N = len(env_df) + 1
    REG_N = len(reg_df) + 1
    assert len(reg_df) == N_REGIONS

    S3G = "ステップ3_月全体"
    S4R = "ステップ4_地域を選ぶ"
    S4B = "ステップ4b_スコア"

    # 地域名のプルダウン用の名前（データ_地域 の A 列）
    wb.defined_names["RegionNames"] = DefinedName(
        "RegionNames", attr_text=f"データ_地域!$A$2:$A${REG_N}")
    # ステップ4 C5 のプルダウン用の名前。非表示列 M5:M13＝先頭に『月全体（線の内側）』＋ データ_地域 の8件（M6:M13 はリンク）。
    #   AllMoon＝M5（定数の文字列はここだけ）／LineOK＝M14（C4 の線が 50〜90 の整数ならその値、そうでなければ空文字）
    wb.defined_names["AreaChoices"] = DefinedName(
        "AreaChoices", attr_text=f"{S4R}!$M$5:$M${5 + N_REGIONS}")
    wb.defined_names["AllMoon"] = DefinedName("AllMoon", attr_text=f"{S4R}!$M$5")
    wb.defined_names["LineOK"] = DefinedName("LineOK", attr_text=f"{S4R}!$M${6 + N_REGIONS}")
    # クレーター名のプルダウン用の名前（データ_地点比較 の A 列。ステップ1c の B5）
    wb.defined_names["SiteNames"] = DefinedName(
        "SiteNames", attr_text=f"データ_地点比較!$A$2:$A${len(sc_df) + 1}")
    LR_N = len(lr_df) + 1      # データ_緯度行 の最終行（91）
    LR = lambda col: f"データ_緯度行!${col}$2:${col}${LR_N}"

    # データ_環境 の列: A lat B lon C 区分 D age_index E temp_amp_K F night_min_K
    #   G noon_sun_elev_deg H earth_elev_deg I region
    #   J norm_sun_high K norm_amp_low L norm_earth_high M norm_earth_low N norm_night_warm
    #   → O スコア（表示用。選んだ範囲の行だけ）／P 順位用（同点を行番号で解く。A3）
    #   O の『採点する行』＝ C5 が『月全体（線の内側）』なら |緯度| が線以下のマス（マス中心）、
    #   地域名なら region が C5 と一致する行。J〜N は全球固定の 0〜1（範囲で正規化し直さない）。
    #   （全球傾斜 slope_deg は Python 版・data/site_environment.csv 側。表計算版は 5 指標）
    env_ws.cell(row=1, column=15, value="スコア").font = BOLD
    env_ws.cell(row=1, column=16, value="順位用（同点を行番号で解く）").font = BOLD
    for r in range(2, E_N + 1):
        env_ws.cell(row=r, column=15, value=(
            f'=IF(AND(SUM({S4R}!$C$8:$C$12)>0,'
            f'OR(AND({S4R}!$C$5=AllMoon,ISNUMBER(LineOK),ABS($A{r})<=LineOK),'
            f'AND({S4R}!$C$5<>"",{S4R}!$C$5<>AllMoon,$I{r}={S4R}!$C$5))),'
            f'({S4R}!$C$8*J{r}+{S4R}!$C$9*K{r}+{S4R}!$C$10*L{r}'
            f'+{S4R}!$C$11*M{r}+{S4R}!$C$12*N{r})'
            f'/MAX(1,{S4R}!$C$8+{S4R}!$C$9+{S4R}!$C$10+{S4R}!$C$11+{S4R}!$C$12),"")'))
        env_ws.cell(row=r, column=16, value=(
            f'=IF(O{r}="","",ROUND(O{r},10)-ROW()*{TIE_EPS})'))
    env_ws.column_dimensions["P"].width = 26

    # データ_南極 の列: A lat B lon C illum D psf E km_to_shadow F slope_deg
    #   G norm_illum H norm_near_shadow I norm_low_psf J norm_low_slope K n
    #   → L スコア／M 順位用（同点を行番号で解く。A3）
    ps_ws.cell(row=1, column=12, value="スコア").font = BOLD
    ps_ws.cell(row=1, column=13, value="順位用（同点を行番号で解く）").font = BOLD
    for r in range(2, PS_N + 1):
        ps_ws.cell(row=r, column=12, value=(
            f"=({S4B}!$C$8*G{r}+{S4B}!$C$9*H{r}+{S4B}!$C$10*I{r}+{S4B}!$C$11*J{r})"
            f"/MAX(1,{S4B}!$C$8+{S4B}!$C$9+{S4B}!$C$10+{S4B}!$C$11)"))
        ps_ws.cell(row=r, column=13, value=(
            f'=IF(SUM({S4B}!$C$8:$C$11)>0,ROUND(L{r},10)-ROW()*{TIE_EPS},"")'))
    ps_ws.column_dimensions["M"].width = 26

    # データ_温度 の列: A=lat B=lon C..Z=t_lt00..23  AA=t_mean_K AB=t_swing_K
    T_SWING = f"データ_温度!$AB$2:$AB${T_N}"
    T_LAT = f"データ_温度!$A$2:$A${T_N}"

    # ================= はじめに =================
    ws = wb.create_sheet("はじめに", 0)
    if student:
        _title(ws, "月データでムーンベースの場所を決めよう（Excel版）")
        _note(ws, "月の公開データ（温度・クレーター・極域の日照と傾斜・地質図）を Excel で分析して、"
                  "「月面基地をどこに建てるか」を自分で決めます。コードは書きません。"
                  "セルの数式を見て、黄色いセルの数字を書き換えて、結果を紙のワークシートに記録します。", 3)
        _h2(ws, "1. 使い方", 5)
        _note(ws, "① 結果を見る前に、まず紙のワークシートに『予想』を書く\n"
                  "② 書き換えてよいのは、指示されたところの黄色いセルだけ"
                  "（ステップ1の4つの緯度帯の黄色いセルは変えない。ほかのセルの数式は消さない）\n"
                  "③ 結果（青い数字）を読んで、紙のワークシートに記録する\n"
                  "④ 数字が空欄に見えるときは、画面上の『編集を有効にする』を押す\n"
                  "⑤ このファイルは上書き保存しない（閉じるときは『保存しない』）", 6)
        _h2(ws, "2. 使うシート", 12)
        for i, (nm, desc) in enumerate([
            ("ステップ1_温度", "（第2時）月の1日の温度は緯度でどう変わる？ 緯度の平均を、どこまで使う？"),
            (S1B, "（第2時）緯度の帯の幅を変えて、グラフを見て、使う緯度の上限（線）を決める"),
            (S1C, "（第2時）担当のクレーターの内部と、同じ緯度の帯の温度を比べる"),
            ("ステップ2_海と陸", "（第2時）『海』と『陸』でクレーターの数（密度）はどう違う？ 地質図と合っている？"),
            ("ステップ4_地域を選ぶ", "（第2時の終わり・第3時）採点する範囲（『月全体（線の内側）』、または地域タイプ）を選び、"
                                "5つの指標に重みをつけて、その範囲でいちばんよい場所を点数で決める"),
            ("ステップ4b_スコア", "（第3時）水を1番にした班：4つの指標に重みをつけて、点数で決める"),
            ("データ_地域", "地域タイプの名前と範囲（ステップ4の C5 の一覧に出る名前）"),
            ("データ_◯◯ / 参考", "分析のもとデータ（前処理済み）。作り方は course/build_course_data.py"),
        ]):
            ws.cell(row=13 + i, column=1, value=nm).font = BOLD
            ws.cell(row=13 + i, column=2, value=desc).font = BODY
        # 入力の順（ステップ4。C4 は画面ではC5の上にあるが、入れる順は C5 が先）
        ws.cell(row=21, column=1, value="ステップ4の入力の順").font = BOLD
        ws.cell(row=21, column=2, value="① C5（採点する範囲）→ ② C4（線。『月全体（線の内側）』を選んだときだけ）→ ③ 重み（C8〜C12）"
                ).font = BODY
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 62
        _fit_notes(ws)
        ws.row_dimensions[6].height = 92     # ①〜⑤の5行が印刷・表示で切れない高さ（⑤＝上書き保存しない）
    else:
        _title(ws, "【教員用】月データでムーンベースの場所を決めよう（Excel版）")
        ws["A2"] = ("教員用：答え・要素の組合せ例（1番・2番と1位）が入っています。生徒には配らないでください"
                    "（配付用は course/student/course_moonbase_student.xlsx）")
        ws["A2"].font = Font(name="Yu Gothic", size=10, bold=True, color="A93226")
        _note(ws, "月の公開データ（温度・クレーター・極域の日照と傾斜・地質図）を Excel で分析して、"
                  "「月面基地をどこに建てるか」を自分で決めます。コードは書きません。"
                  "セルの数式を見て、黄色いセルの数字を書き換えて、結果をワークシートに記録します。", 3)
        _h2(ws, "1. 基地の要素と、表計算の行（重みの置き方）", 5)
        ws["A6"], ws["B6"] = "要素", "区分・点の置き方・表計算の行"
        ws["A6"].font = ws["B6"].font = BOLD
        for i, (m, need) in enumerate([
            ("電力", "共通。必ず1点（1番・2番にはしない）。ステップ4＝太陽高度の行／4b＝日照率の行"),
            ("建設", "共通。1番にはしない（検算で傾斜を読む）。ステップ4に行なし／4b＝傾斜の行（水を1番にした班は2番〔2点〕にしてよい）"),
            ("温度", "班が分析。1番＝3点、2番＝2点。ステップ4＝日較差の行と夜の最低温度の行"
                   "（変換表：1番＝日較差2・夜1、2番＝日較差1・夜1）／4bに行なし"),
            ("通信", "班が分析。1番＝3点、2番＝2点。ステップ4＝地球の仰角（表側）の行（裏側の行は0）／4bに行なし"),
            ("水", "班が分析。1番＝3点。4b＝永久影までの距離の行（4bで1番にできるのは水だけ）／ステップ4に行なし"),
            ("（観測）", "教員用の発展。ステップ4の地球の仰角（裏側）の行＝通信の鏡像。授業では使わない"),
        ]):
            ws.cell(row=7 + i, column=1, value=m).font = BODY
            ws.cell(row=7 + i, column=2, value=need).font = BODY
            ws.cell(row=7 + i, column=2).alignment = WRAP
            ws.row_dimensions[7 + i].height = 30
        _note(ws, "同じデータでも、何を重視するかで答えが変わります。使うシートは1番で決まります："
                  "水を1番にした班＝ステップ4b、それ以外＝ステップ4で C5『月全体（線の内側）』を選び、C4 に線（第2時で引いた線）を入れる。"
                  "電力は1点で入れてあります（学習者版は C8＝1）。", 13)

        _h2(ws, "2. 進め方（各ステップ共通）", 15)
        _note(ws, "① まずワークシートに『予想』を書く（数式を見る前に）\n"
                  "② 黄色いセルの数字を書き換えて、結果（青いセル）を読む\n"
                  "③ ワークシートに結果と『気づいたこと』を書く", 16)

        _h2(ws, "3. シートの並び", 20)
        for i, (nm, desc) in enumerate([
            ("ステップ1_温度", "月の1日の温度は緯度でどう変わる？ 緯度の平均を、どこまで使う？"),
            (S1B, "緯度の帯の幅を変えて、グラフを見て、使う緯度の上限（線）を決める（判断①）"),
            (S1C, "担当のクレーターの内部と、同じ緯度の帯の温度を比べる（判断②）"),
            ("ステップ2_海と陸", "『海』と『陸』でクレーターの数・大きさ・年代はどう違う？ 地質図と合っている？"),
            ("ステップ3_月全体", "月全体で環境を見る（日較差・地球の仰角・太陽高度）。どこも『全部で一番』にはならない"),
            ("ステップ4_地域を選ぶ", "採点する範囲（月全体〔線の内側〕、または地域タイプ）を選び、その範囲でいちばんよい場所を点数で決める"),
            ("ステップ4b_南極", "（探索用）南極の日照・傾斜・永久影を細かく見る"),
            ("ステップ4b_スコア", "（水を1番にした班）南極の4つの指標に重みをつけて、点数で決める"),
            ("ステップ5_まとめ", "各班の1番・使った範囲・1位を見比べる。実在の計画は、どの要素を重んじたかで読む"),
            ("データ_◯◯ / 参考", "分析のもとデータ（前処理済み）。作り方は course/build_course_data.py"),
        ]):
            ws.cell(row=21 + i, column=1, value=nm).font = BOLD
            ws.cell(row=21 + i, column=2, value=desc).font = BODY
        _note(ws, "参考：同じ分析を Python（pandas）で書くと？ は notebooks/course_moonbase.ipynb にあります。"
                  "Excel でやったことと数値がそろうように作ってあります。", 32)
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 62
        _fit_notes(ws)
        ws.row_dimensions[16].height = 60    # ①〜③の3行が切れない高さ
    _fit_notes(ws)

    # ================= ステップ1：温度 =================
    ws = wb.create_sheet("ステップ1_温度")
    _title(ws, "ステップ1：月の温度は1日でどれくらい変わる？")
    if student:
        _note(ws, "月には大気がない。「1日の温度の較差（＝最高－最低）」を、緯度の絶対値の帯ごとに平均して比べる。"
                  "A の4つの帯は、『データ_緯度行』（月全体の0.5度のセルを、緯度の絶対値1度ごとの行にまとめた表）から求める。"
                  "B の24時間カーブは、『データ_温度』（3度ごとの世界地図。7200地点）の1地点。", 3)
    else:
        _note(ws, "月には大気がない＝温室効果も熱の運搬もない。太陽が当たる昼と、当たらない夜の差が大きい。"
                  "「1日の温度の較差（＝最高－最低）」を、緯度の絶対値の帯ごとに平均して比べる。"
                  "A の4つの帯は、『データ_緯度行』（月全体の0.5度のセル259,200個を、緯度の絶対値1度ごとの90行にまとめた表。"
                  "旧い版は3度標本で、値が少し違った）から求める。"
                  "B の24時間カーブ・C の着陸地点は、『データ_温度』（3度ごとの世界地図。7200地点＝0.5度のセルを3度おきにとった標本）の1地点。", 3)

    _h2(ws, ("A. 緯度帯ごとの1日の温度較差（黄色い緯度帯は変えずに、結果を読む）" if student
             else "A. 緯度帯ごとの1日の温度較差（黄色いセルに緯度を入れる）"), 6)
    ws["A7"], ws["B7"], ws["C7"], ws["D7"] = ("緯度の絶対値（下）", "緯度の絶対値（上）", "セル数（0.5度）",
                                              "1日の較差の平均 [K]")
    for c in "ABCD":
        ws[c + "7"].font = BOLD
    # 第2時（rev2）：4バンドは 0.5度の全セルを緯度の絶対値1度ごとにまとめた『データ_緯度行』から求める（北南を合わせる）
    bands = [(0, 6, "赤道"), (24, 36, "中緯度"), (54, 66, "高緯度"),
             (78, 90, "極付近" if student else "極付近（要注意）")]
    for i, (lo, hi, label) in enumerate(bands):
        r = 8 + i
        _input(ws, f"A{r}", lo)
        _input(ws, f"B{r}", hi)
        ws[f"C{r}"] = (f'=SUMIFS({LR("C")},{LR("A")},">="&A{r},{LR("B")},"<="&B{r})')
        ws[f"D{r}"] = (f'=ROUND(AVERAGEIFS({LR("D")},{LR("A")},">="&A{r},'
                       f'{LR("B")},"<="&B{r}),0)')
        ws[f"D{r}"].font = BLUE
        ws[f"E{r}"] = label
        ws[f"E{r}"].font = BODY
    _note(ws, "気づき：緯度が高くなると較差は？　赤道の較差（約○○K＝約○○℃）を、"
              "地球の砂漠の昼夜差（20〜30℃）と比べると？　"
              "帯の幅を変えたとき、グラフはどう変わる？（ステップ1b）　1行で言う", 13)

    _h2(ws, "B. 1地点の24時間カーブを見る（黄色いセルに地点を入れる）", 16)
    _input(ws, "B17", 0.25)
    _input(ws, "B18", 0.25)
    ws["A17"], ws["A18"] = "見たい地点の緯度", "見たい地点の経度"
    ws["A17"].font = ws["A18"].font = BODY
    _note(ws, "※ 緯度・経度は『データ_温度』にある値ちょうどを入れる（緯度は …-2.75, 0.25, 3.25, 6.25…／"
              "経度も同じきざみ）。見つからないと温度が 0 のまま。", 19, span=8)
    ws["A21"] = "現地時間"
    ws["B21"] = "温度 [K]"
    ws["A21"].font = ws["B21"].font = BOLD
    for h in range(24):
        r = 22 + h
        ws[f"A{r}"] = h
        col = get_column_letter(3 + h)   # t_lt00 は データ_温度 の C列
        ws[f"B{r}"] = (f'=SUMIFS(データ_温度!${col}$2:${col}${T_N},'
                       f'データ_温度!$A$2:$A${T_N},$B$17,'
                       f'データ_温度!$B$2:$B${T_N},$B$18)')
    line = LineChart()
    line.title = "選んだ地点の1日の温度カーブ"
    line.y_axis.title = "温度 [K]"
    line.x_axis.title = "現地時間"
    data = Reference(ws, min_col=2, min_row=21, max_row=45)
    cats = Reference(ws, min_col=1, min_row=22, max_row=45)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    line.height, line.width = 8, 15
    # openpyxl 3.1 の既定では軸が非表示になり、目盛が出ない。軸を表示し、系列を単色の直線にする
    line.x_axis.delete = False
    line.y_axis.delete = False
    line.legend = None
    line.title.overlay = False
    line.varyColors = False
    line.series[0].smooth = False
    line.series[0].graphicalProperties.line.solidFill = "1A6FB0"
    line.series[0].graphicalProperties.line.width = 22000
    _tidy_axes(line, "0")       # 軸タイトルが目盛に重なる（Excel 実機で確認）のを直す
    # G10：以前は D16 に置いて、注意書き（A19:H19）に重なっていた。24時間の表（A21:B45）の右に移す
    ws.add_chart(line, "D21")
    _note(ws, "気づき：いちばん暑い時刻・いちばん寒い時刻はいつ？　朝と夕方でカーブの形は左右対称？　"
              "緯度を -86.75（経度は 0.25 のまま）にして、カーブの形を見る（先生のデモ）。", 47)

    _h2(ws, "C. 実際に人が降りた場所の温度（『データ_着陸地点』より）", 50)
    ws["A51"], ws["B51"], ws["C51"], ws["D51"] = "着陸地点", "正午 [K]", "真夜中 [K]", "1日の差 [K]"
    for c in "ABCD":
        ws[c + "51"].font = BOLD
    for i, nm in enumerate(["Apollo 11", "Apollo 15", "Chang'e 4", "Chandrayaan-3"]):
        r = 52 + i
        ws[f"A{r}"] = nm
        ws[f"A{r}"].font = BODY
        m = f'MATCH("{nm}",データ_着陸地点!$A$2:$A${len(ls_df) + 1},0)'
        ws[f"B{r}"] = f'=INDEX(データ_着陸地点!$G$2:$G${len(ls_df) + 1},{m})'
        ws[f"C{r}"] = f'=INDEX(データ_着陸地点!$H$2:$H${len(ls_df) + 1},{m})'
        ws[f"D{r}"] = f'=INDEX(データ_着陸地点!$I$2:$I${len(ls_df) + 1},{m})'
        for c in "BCD":
            ws[f"{c}{r}"].font = BLUE
    _note(ws, "気づき：赤道の海（Apollo 11）と高緯度（Chandrayaan-3）で、1日の差はどう違う？", 56)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 14
    _fit_notes(ws)

    # ================= ステップ1b・1c（第2時の新シート） =================
    build_step1b(wb, student, lr_df, LR)
    build_step1c(wb, student, sc_df)

    # ================= ステップ2：海と陸 =================
    ws = wb.create_sheet("ステップ2_海と陸")
    _title(ws, "ステップ2：月の「海」と「陸」で何が違う？")
    if student:
        _note(ws, "月の黒い部分＝『海』（マリア）は、昔の溶岩でおおわれた低地。"
                  "『データ_クレーター』の各行に、その場所が海か陸かの区分がついている（円で囲む近似。作り方は『参考』シート）。"
                  "海と陸でクレーターの数（密度）を比べ、最後に USGS の公式地質図の年代と『答え合わせ』する。", 3)
    else:
        _note(ws, "月の黒い部分＝『海』（マリア）は、昔の溶岩でおおわれた低地。"
                  "『データ_クレーター』の各行に、その場所が海か陸かの区分がついている（円で囲む近似。作り方は『参考』シート）。"
                  "数（密度）・大きさ・年代を海と陸で比べ、最後に USGS の公式地質図と『答え合わせ』する。"
                  "（3コマ版の授業では B・C の区画は使わない。D は USGS の年代だけを読む）", 3)
    C_KUBUN = f"データ_クレーター!$D$2:$D${C_N}"
    C_DIAM = f"データ_クレーター!$C$2:$C${C_N}"
    A_KUBUN = f"データ_クレーター年代!$F$2:$F${A_N}"
    A_AGE = f"データ_クレーター年代!$D$2:$D${A_N}"

    _h2(ws, "A. クレーターの数と密度", 6)
    rows = [
        ("海のクレーター数", f'=COUNTIF({C_KUBUN},"海")'),
        ("陸のクレーター数", f'=COUNTIF({C_KUBUN},"陸")'),
        ("海の面積割合（円近似・参考シート）", "=参考!B2"),
        ("陸の面積割合（円近似）", "=参考!B3"),
        ("海の密度（数÷面積割合）", "=ROUND(B7/B9,0)"),
        ("陸の密度（数÷面積割合）", "=ROUND(B8/B10,0)"),
        ("陸は海の何倍こみあっている？", "=ROUND(B12/B11,1)"),
    ]
    for i, (label, formula) in enumerate(rows):
        r = 7 + i
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = formula
        ws[f"B{r}"].font = BLUE

    if not student:
        # B・C 区画（3コマ版では使わない。学習者版には作らない）
        _h2(ws, "B. クレーターの大きさ（直径の平均）", 16)
        ws["A17"], ws["B17"] = "海の直径の平均 [km]", f'=ROUND(AVERAGEIF({C_KUBUN},"海",{C_DIAM}),1)'
        ws["A18"], ws["B18"] = "陸の直径の平均 [km]", f'=ROUND(AVERAGEIF({C_KUBUN},"陸",{C_DIAM}),1)'

        _h2(ws, "C. クレーターの年代（1=最も古い 〜 5=最も新しい）", 20)
        ws["A21"], ws["B21"] = "海の年代の平均", f'=ROUND(AVERAGEIF({A_KUBUN},"海",{A_AGE}),2)'
        ws["A22"], ws["B22"] = "陸の年代の平均", f'=ROUND(AVERAGEIF({A_KUBUN},"陸",{A_AGE}),2)'
        for r in (17, 18, 21, 22):
            ws[f"A{r}"].font = BODY
            ws[f"B{r}"].font = BLUE

    # D 区画（学習者版では B. 答え合わせ。円近似と USGS の面積割合の行 29〜31 は作らない）
    _h2(ws, "B. 答え合わせ：USGS の公式地質図と比べる" if student
        else "D. 答え合わせ：USGS の公式地質図と比べる", 25)
    checks = [
        ("あなたの推定：海と陸、古いのはどっち？",
         "（A の密度から考える）" if student
         else "（C の年代の平均が小さいほう。3コマ版の授業では A の密度から考える）", None),
        ("地質図：海の相対年代の平均", "＝参考!B6", "=参考!B6"),
        ("地質図：陸の相対年代の平均", "＝参考!B7（大きいほど新しい）", "=参考!B7"),
    ]
    if not student:
        checks += [
            ("海の面積割合：あなたの円近似", "＝参考!B2", "=参考!B2"),
            ("海の面積割合：USGS 地質図", "＝参考!B4（文献値 約16%）", "=参考!B4"),
            ("円近似は USGS より大きい？小さい？その差は何%ポイント？",
             "=ROUND((参考!B2-参考!B4)*100,1)", "=ROUND((参考!B2-参考!B4)*100,1)"),
        ]
    for i, (label, hint, formula) in enumerate(checks):
        r = 26 + i
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        if formula:
            ws[f"B{r}"] = formula
            ws[f"B{r}"].font = BLUE
        ws[f"C{r}"] = hint
        ws[f"C{r}"].font = Font(name="Yu Gothic", size=9, color="777777")
    if student:
        _note(ws, "気づき：①海と陸で、クレーターの密度はどう違う？　"
                  "②あなたの推定は、地質図の年代と合っていた？", 33)
        # 使わない区画の行を隠す（B・C 区画の行は作らないで空けてある）
        for r in list(range(16, 23)) + [29, 30, 31]:
            ws.row_dimensions[r].hidden = True
    else:
        _note(ws, "気づき：①海のほうがクレーターが少ないのはなぜ？（隕石が落ちなかった？ 落ちた後に消えた？）"
                  "②あなたのクレーターからの推定は、地質図と合っていた？　"
                  "③円で海を囲むと面積割合が大きく出るのはなぜ？（海岸線の外側の陸も丸に入るから）", 33)
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 30
    _fit_notes(ws)

    # ================= ステップ3：月全体で環境を見る =================
    # 学習者版では非表示（3コマ版では使わない）。中身は答えに触れない書き方にしてある。
    ws = wb.create_sheet(S3G)
    _title(ws, "ステップ3：月全体で環境を見る")
    if student:
        _note(ws, "『データ_環境』は月全体を3度マスに区切った表。temp_amp_K=1日の温度差、"
                  "night_min_K=夜の最低温度、noon_sun_elev_deg=正午の太陽高度、"
                  "earth_elev_deg=地球の仰角（正なら地球が地平線の上、負なら下）。", 3)
    else:
        _note(ws, "基地の場所を決めるには、温度だけでなく『地球が見えるか（通信・電波静穏）』『太陽がどれだけ高いか』"
                  "も要る。『データ_環境』は月全体を3度マスに区切った表。temp_amp_K=1日の温度差、"
                  "night_min_K=夜の最低温度、noon_sun_elev_deg=正午の太陽高度、"
                  "earth_elev_deg=地球の仰角（正＝表側で通信できる／負＝裏側で電波が静か）。", 3)
    E_LAT = f"データ_環境!$A$2:$A${E_N}"
    E_AMP = f"データ_環境!$E$2:$E${E_N}"
    E_EARTH = f"データ_環境!$H$2:$H${E_N}"
    E_SUN = f"データ_環境!$G$2:$G${E_N}"
    E_REG = f"データ_環境!$I$2:$I${E_N}"

    _h2(ws, "A. 緯度帯ごとの1日の温度差（黄色いセルに緯度）", 6)
    ws["A7"], ws["B7"], ws["C7"], ws["D7"] = "緯度の下", "緯度の上", "地点数", "1日の温度差の平均 [K]"
    for c in "ABCD":
        ws[c + "7"].font = BOLD
    for i, (lo, hi, label) in enumerate([(-6, 6, "赤道"), (24, 36, "中緯度"),
                                         (54, 66, "高緯度"), (81, 90, "極付近")]):
        r = 8 + i
        _input(ws, f"A{r}", lo)
        _input(ws, f"B{r}", hi)
        ws[f"C{r}"] = f'=COUNTIFS({E_LAT},">="&A{r},{E_LAT},"<="&B{r})'
        ws[f"D{r}"] = f'=ROUND(AVERAGEIFS({E_AMP},{E_LAT},">="&A{r},{E_LAT},"<="&B{r}),0)'
        ws[f"D{r}"].font = BLUE
        ws[f"E{r}"] = label
        ws[f"E{r}"].font = BODY

    _h2(ws, "B. 地域タイプごとの環境（『データ_地域』の名前で引く）", 14)
    ws["A15"], ws["B15"], ws["C15"], ws["D15"], ws["E15"] = (
        "地域タイプ", "地点数", "日較差の平均 [K]", "太陽高度の平均 [度]", "地球の仰角の平均 [度]")
    for c in "ABCDE":
        ws[c + "15"].font = BOLD
    for i, nm in enumerate(["赤道の海（静かの海）", "中緯度の火砕丘（Aristarchus 高原）",
                            REG_FAR_EQ, "南極（Shackleton-de Gerlache）"]):
        r = 16 + i
        ws[f"A{r}"] = nm
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = f'=COUNTIF({E_REG},$A{r})'
        ws[f"C{r}"] = f'=ROUND(AVERAGEIF({E_REG},$A{r},{E_AMP}),0)'
        ws[f"D{r}"] = f'=ROUND(AVERAGEIF({E_REG},$A{r},{E_SUN}),0)'
        ws[f"E{r}"] = f'=ROUND(AVERAGEIF({E_REG},$A{r},{E_EARTH}),0)'
        for c in "BCDE":
            ws[f"{c}{r}"].font = BLUE
    if student:
        _note(ws, "気づき：地域タイプごとの日較差・太陽高度・地球の仰角を比べて、気づいたことを書く。", 22)
    else:
        _note(ws, "気づき：『温度が安定』『地球が見える』『日がよく当たる』が全部そろう地域はあった？　"
                  "南極は日較差が小さいが太陽高度は？　裏側は地球の仰角が負（＝地球が地平線の下）。"
                  "担当の要素で、いちばん大事な列はどれ？", 22)
    ws.column_dimensions["A"].width = 32
    for c in "BCDE":
        ws.column_dimensions[c].width = 16
    _fit_notes(ws)

    # ================= ステップ4：採点する範囲（月全体の線の内側、または地域タイプ）を選んで評価する =================
    ws = wb.create_sheet(S4R)
    _title(ws, "ステップ4：採点する範囲を選んで、重みをつけて評価する")
    if student:
        _note(ws, "採点する範囲を1つ選び（C5 をクリックして一覧から選ぶ。一覧の先頭は『月全体（線の内側）』、その下が『データ_地域』の地域タイプ）、"
                  "5つの指標を 0〜1 に直して重みをつけて合計する（＝あなたのスコア式）。"
                  "『月全体（線の内側）』を選んだときは、C4 に線（第2時で引いた線）を入れる。"
                  "『データ_環境』の O 列『スコア』が、選んだ範囲の行だけ自動で計算される。", 3)
    else:
        _note(ws, "採点する範囲を1つ選び（C5 の一覧から選ぶ。先頭は『月全体（線の内側）』、その下が『データ_地域』の地域タイプ）、"
                  "5つの指標を 0〜1 に直して重みをつけて合計する（＝あなたのスコア式）。"
                  "月全体を選んだときは、C4 の線（|緯度| の上限。マス中心が線以下のマスを採点する）を使う。0〜1 への直し方は全球固定で、線を変えても変わらない。"
                  "『データ_環境』の O 列『スコア』が、選んだ範囲の行だけ自動で計算される。", 3)
    # C4：線（新しい黄色セル。月全体を選ぶときだけ使う。学習者版は空欄で配る）
    ws["A4"] = ("線 [°]（C5 が『月全体（線の内側）』のときだけ。温度のデータを使う範囲として、第2時で引いた線。50〜90 の整数）" if student
                else "線 [°]（C5 が『月全体（線の内側）』のときだけ。温度のデータを使う範囲として、第2時で引いた線。50〜90 の整数）")
    ws["A4"].font = BODY
    ws["A4"].alignment = WRAP
    ws.merge_cells("A4:B4")
    ws.row_dimensions[4].height = 44
    _input(ws, "C4", None if student else TEACHER_LINE)
    dv_line = DataValidation(type="whole", operator="between", formula1="50", formula2="90",
                             allow_blank=True, showErrorMessage=True, showInputMessage=True, errorStyle="stop",
                             errorTitle="線の入れかた", error="線は 50〜90 の整数（緯度の絶対値 [°]）で入れてください。",
                             promptTitle="線 [°]", prompt="50〜90 の整数。温度のデータを使う範囲（この緯度の絶対値まで）")
    ws.add_data_validation(dv_line)
    dv_line.add("C4")
    ws["A5"] = ("採点する範囲（C5 をクリックして一覧から選ぶ）" if student
                else "採点する範囲（C5 の一覧から選ぶ。『月全体（線の内側）』か、『データ_地域』の name をそのまま）")
    ws["A5"].font = BODY
    ws["A5"].alignment = WRAP
    ws.merge_cells("A5:B5")
    ws.row_dimensions[5].height = 30 if student else 44
    _input(ws, "C5", None if student else "赤道の海（静かの海）")

    # 非表示列 M：プルダウンの一覧。M5＝月全体（定数の文字列はここだけ）、M6:M13＝データ_地域 のリンク、M14＝LineOK（式の中で名前で使う）
    ws["M5"] = ALL_MOON
    for i in range(N_REGIONS):
        ws[f"M{6 + i}"] = f"=データ_地域!A{2 + i}"
    ws[f"M{6 + N_REGIONS}"] = '=IFERROR(IF(AND(ISNUMBER($C$4),$C$4>=50,$C$4<=90,$C$4=INT($C$4)),$C$4,""),"")'
    ws.column_dimensions["M"].hidden = True

    # C5 は AreaChoices（先頭＝月全体、続けて データ_地域 の名前）のプルダウン
    dv = DataValidation(type="list", formula1="AreaChoices", allow_blank=True,
                        showErrorMessage=True, showInputMessage=True, errorStyle="stop",
                        errorTitle="範囲の名前が違います",
                        error="一覧にある名前（『月全体（線の内側）』か『データ_地域』の地域タイプ）を選んでください。",
                        promptTitle="採点する範囲", prompt="クリックして、一覧から1つ選ぶ")
    ws.add_data_validation(dv)
    dv.add("C5")

    # 状態表示セル（A6）。C5・C4・重みの状態を1行で出す。先頭の文は従来のまま（既存の検査が見る）
    sum_w = "SUM($C$8:$C$12)"
    bad_w = f"OR(MIN($C$8:$C$12)<0,MAX($C$8:$C$12)>5,SUMPRODUCT(--($C$8:$C$12<>INT($C$8:$C$12)))>0)"
    n_lat = f"COUNTIFS(データ_環境!$A$2:$A${E_N},\"<=\"&LineOK,データ_環境!$A$2:$A${E_N},\">=\"&-LineOK)"
    ws["A6"] = (
        '=IFERROR(IF($C$5="","地域タイプを選んでください（C5 をクリックすると一覧が出ます。月全体で採点するときは『月全体（線の内側）』）",'
        'IF(COUNTIF(AreaChoices,$C$5)=0,"【注意】C5 の地域名が『データ_地域』にありません。一覧から選んでください",'
        'IF($C$5=AllMoon,'
        'IF(NOT(ISNUMBER(LineOK)),"【注意】月全体で採点するには、C4 に線（50〜90 の整数。第2時で引いた線）を入れてください",'
        f'IF({sum_w}=0,"【注意】重みがすべて0です。どれかを1以上にしてください",'
        f'IF({bad_w},"【注意】重みは0〜5の整数で入れてください",'
        'IF(SUM($C$9:$C$12)=0,"【注意】太陽高度の行だけでは、同じ点数の場所が大量に出ます。ほかの行にも重みを置いてください",'
        f'"OK　月全体（線 "&LineOK&"° の内側）　採点対象 "&{n_lat}&" 行　（重みの合計 "&{sum_w}&"）"'
        '&IF($C$8=0,"　※太陽高度の行が0です",""))))),'
        'IF(LEFT($C$5,2)="南極","南極は『ステップ4b_スコア』を使います（ここのトップ10は参考にしない）",'
        f'IF({sum_w}=0,"【注意】重みがすべて0です。どれかを1以上にしてください",'
        f'IF({bad_w},'
        '"【注意】重みは0〜5の整数で入れてください",'
        f'"OK　採点対象 "&COUNTIF(データ_環境!$I$2:$I${E_N},$C$5)&" 行　（重みの合計 "&{sum_w}&"）")))))),'
        '"【注意】重みの欄に数字以外が入っています")')
    ws["A6"].font = BOLD
    _status_format(ws, "A6")

    _h2(ws, "A. あなたのスコア式：重み（黄色いセル。0なら気にしない。0〜5の整数）", 7)
    # 学習者版：向きだけ。「（発電）」「（通信できる）」などの用途は書かない。C8（電力＝太陽高度の行）は 1 で配る
    if student:
        w_rows = [
            ("太陽高度 noon_sun_elev_deg", "高いほどよい", 1),
            ("1日の温度差 temp_amp_K", "小さいほどよい", 0),
            ("地球の仰角 earth_elev_deg（表側）", "高いほどよい", 0),
            ("地球の仰角 earth_elev_deg（裏側）", "低いほどよい", 0),
            ("夜の最低温度 night_min_K", "高いほどよい", 0),
        ]
    else:
        w_rows = [
            ("太陽高度 noon_sun_elev_deg", "高いほどよい（電力）", 2),
            ("1日の温度差 temp_amp_K", "小さいほどよい（温度の安定）", 2),
            ("地球の仰角 earth_elev_deg（表側）", "高いほどよい（通信できる）", 0),
            ("地球の仰角 earth_elev_deg（裏側）", "低いほどよい（観測用。授業では使わない）", 0),
            ("夜の最低温度 night_min_K", "高いほどよい（夜に冷えすぎない）", 0),
        ]
    for i, (label, good, w) in enumerate(w_rows):
        r = 8 + i
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = good
        ws[f"B{r}"].font = BODY
        _input(ws, f"C{r}", w)
    _weight_validation(ws, "C8:C12")
    ws["A13"] = "重みの合計"
    ws["A13"].font = BOLD
    ws["B13"] = "=C8+C9+C10+C11+C12"
    ws["B13"].font = BLUE

    if not student:
        # 月全体の例（教員版のみ）。数値は env_grid.csv から _top_all で再計算（検査スクリプトは別に pandas で再現して突き合わせる）
        SETS = [("A 温度を1番", (1, 2, 0, 0, 1)), ("B 通信を1番", (1, 0, 3, 0, 0)),
                ("C 温度を1番＋通信を2番", (1, 2, 2, 0, 1)), ("D 通信を1番＋温度を2番", (1, 1, 3, 0, 1))]
        base = {nm: {ln: _top_all(env_df, ln, w)[0] for ln in range(50, 91)} for nm, w in SETS}
        _h2(ws, "B. 月全体の例（一例。C5＝月全体、C4＝線、重みを入れ直して使う。1位は再計算した値）", 15)
        for c, h in zip("ABCDEF", ("設定（電力は1点）", "重み（太陽・温度差・地球表・地球裏・夜）", "1位の緯度（線50〜88）",
                                   "1位の経度", "電力を0にすると（線70）", "線を90にすると")):
            ws[f"{c}16"] = h
            ws[f"{c}16"].font = BOLD
            ws[f"{c}16"].alignment = WRAP
        ws.row_dimensions[16].height = 44
        for i, (nm, w) in enumerate(SETS):
            r = 17 + i
            b = base[nm][70]
            assert all(base[nm][ln] == b for ln in range(50, 89)), "線50〜88で1位が変わった：月全体の例の文を見直す"
            ws[f"A{r}"] = nm
            ws[f"B{r}"] = "・".join(str(x) for x in w)
            ws[f"C{r}"], ws[f"D{r}"] = b
            ws[f"E{r}"] = _fmt_pt(_top_all(env_df, 70, (0,) + w[1:])[0])
            ws[f"F{r}"] = _fmt_pt(base[nm][90])
            for c in "ABCDEF":
                ws[f"{c}{r}"].font = BODY
        ws["E16"].alignment = WRAP
        # 線ごとの1位の動き（A・C の電力0は線で動く）と、教員の一言の材料
        a0 = {ln: _top_all(env_df, ln, (0, 2, 0, 0, 1))[0] for ln in range(50, 91)}
        c0 = {ln: _top_all(env_df, ln, (0, 2, 2, 0, 1))[0] for ln in range(50, 91)}

        def _runs(d):
            out, lo = [], 50
            for ln in range(51, 92):
                if ln == 91 or d[ln] != d[lo]:
                    out.append(f"{lo}〜{ln - 1}：{_fmt_pt(d[lo])}")
                    lo = ln
            return "／".join(out)

        _note(ws, "線ごとの1位（電力1点）：A＝" + _runs(base["A 温度を1番"]) + "。B・C・D は線50〜90で同じ。"
                  "試行(a) 1番を替える：A⇄B（"
                  + _fmt_pt(base["A 温度を1番"][70]) + "⇄" + _fmt_pt(base["B 通信を1番"][70]) + "）、C⇄D（"
                  + _fmt_pt(base["C 温度を1番＋通信を2番"][70]) + "⇄" + _fmt_pt(base["D 通信を1番＋温度を2番"][70]) + "）。", 21)
        _note(ws, "電力を0にすると、A は線で動く：" + _runs(a0) + "。C は" + _runs(c0)
                  + "。B・D は動かない。電力1点と夜1点が、線の効きを消している。"
                  "上位10のスコアは重みが違う班どうしで比べない。", 22)

    _h2(ws, ("B. スコアの高い順トップ10（C5・重みを変えると入れ替わる）" if student
             else "C. スコアの高い順トップ10（C5・重みを変えると入れ替わる）"), 24)
    E_SCORE = f"データ_環境!$O$2:$O${E_N}"
    E_RANK = f"データ_環境!$P$2:$P${E_N}"     # 順位用（同点を行番号で解く）
    ws["A25"], ws["B25"], ws["C25"], ws["D25"], ws["E25"], ws["F25"], ws["G25"], ws["H25"] = (
        "順位", "スコア", "緯度", "経度", "日較差 [K]", "太陽高度 [度]", "地球の仰角 [度]", "夜の最低温度 [K]")
    for c in "ABCDEFGH":
        ws[c + "25"].font = BOLD
    for k in range(1, 11):
        r = 25 + k
        ws[f"A{r}"] = k
        m = f'MATCH(LARGE({E_RANK},A{r}),{E_RANK},0)'
        ws[f"B{r}"] = f'=IFERROR(ROUND(INDEX({E_SCORE},{m}),3),"")'
        ws[f"C{r}"] = f'=IFERROR(INDEX(データ_環境!$A$2:$A${E_N},{m}),"")'
        ws[f"D{r}"] = f'=IFERROR(INDEX(データ_環境!$B$2:$B${E_N},{m}),"")'
        ws[f"E{r}"] = f'=IFERROR(INDEX(データ_環境!$E$2:$E${E_N},{m}),"")'
        ws[f"F{r}"] = f'=IFERROR(INDEX(データ_環境!$G$2:$G${E_N},{m}),"")'
        ws[f"G{r}"] = f'=IFERROR(INDEX(データ_環境!$H$2:$H${E_N},{m}),"")'
        ws[f"H{r}"] = f'=IFERROR(INDEX(データ_環境!$F$2:$F${E_N},{m}),"")'
        for c in "BCDEFGH":
            ws[f"{c}{r}"].font = BLUE
    # 上位10の範囲。月全体では上位10が離れた場所に散らばるので日付変更線の文は出さず、散らばりの注意と、
    # 1位が線のいちばん外側のマスのときの確認の一文を出す（地域モードは従来どおり）
    ws["A36"] = "上位10の範囲"
    ws["A36"].font = BOLD
    ws["B36"] = (
        '=IF(COUNT(C26:C35)=0,"",'
        'IF($C$5=AllMoon,'
        '"緯度 "&TEXT(MIN(C26:C35),"0.0")&" 〜 "&TEXT(MAX(C26:C35),"0.0")&" °　／　経度 "&TEXT(MIN(D26:D35),"0.0")&" 〜 "&TEXT(MAX(D26:D35),"0.0")&" °"'
        '&IF(MAX(D26:D35)-MIN(D26:D35)>180,"　（経度が広く散らばっています。10行を1つずつ見ます）","")'
        '&IF(ABS(C26)>=LineOK-3,"　【確認】1位が線のいちばん外側のマスにあります。線を少し内側にして確かめます",""),'
        '"緯度 "&TEXT(MIN(C26:C35),"0.0")&" 〜 "&TEXT(MAX(C26:C35),"0.0")&" °　／　"&'
        'IF(MAX(D26:D35)-MIN(D26:D35)>180,"経度は日付変更線をまたぐので範囲を出しません",'
        '"経度 "&TEXT(MIN(D26:D35),"0.0")&" 〜 "&TEXT(MAX(D26:D35),"0.0")&" °")))')
    ws["B36"].font = BLUE
    if student:
        _note(ws, "気づき：選んだ範囲のトップ10は、どんな特徴をもっている？　"
                  "重みを変えると、トップ10はどう変わった？（変える前の結果は、紙に写しておく）", 37)
        for r in range(15, 23):      # 学習者版には教員版の『月全体の例』の表は作らない。行を隠す
            ws.row_dimensions[r].hidden = True
    else:
        _note(ws, "気づき（教員用）：月全体の上位10は、温度を1番にした設定（A）では離れた複数の場所に分かれ、通信を1番にした設定（B・D）では1か所に集まる"
                  "（北と南で同点のときは、表示は南が先）。B36 に【確認】が出たら、1位が線のそばにある＝選んだのは重みではなく線。"
                  "線を5°動かして確かめる。スコアは、重みが違う班どうしで比べない。", 37)
    ws.column_dimensions["A"].width = 34
    for c in "BCDEFG":
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["C"].width = 20      # C5 の『月全体（線の内側）』が黄色いセルに収まる幅
    ws.column_dimensions["H"].width = 16
    if not student:
        ws.column_dimensions["B"].width = 34      # 重みの行の『高いほどよい（…）』が切れない幅
        ws.column_dimensions["E"].width = 22
        ws.column_dimensions["F"].width = 20
        ws.row_dimensions[4].height = 38
        ws.row_dimensions[5].height = 38
    _fit_notes(ws)

    # ================= ステップ4b：南極の日照と傾斜 =================
    # 学習者版では非表示（3コマ版では使わない）
    ws = wb.create_sheet("ステップ4b_南極")
    _title(ws, "ステップ4b：南極でよい場所（日当たり・平ら）はどれくらい？")
    _note(ws, "極では太陽が地平線近くを回るだけなので『昼夜』がない。地形の高いところは年中日が当たり、"
              "クレーターの底は年中影（永久影）。基地には日当たりだけでなく『地面が平ら（傾斜が小さい）』ことも要る。"
              "『データ_南極』の average_illumination_percent は年間日照率、slope_deg は傾斜[度]。"
              "※どちらも絶対値は他の資料と単純比較しない（「暗い/明るい」「平ら/急」の順番だけ信じる）。", 3)
    PS_ILLUM = f"データ_南極!$C$2:$C${PS_N}"
    PS_SLOPE = f"データ_南極!$F$2:$F${PS_N}"

    _h2(ws, "A. しきい値より日照率が高い地点の数（黄色いセルにしきい値[%]）", 6)
    ws["A7"], ws["B7"], ws["C7"] = "しきい値 [%]", "その値以上の地点数", "全体に占める割合"
    for c in "ABC":
        ws[c + "7"].font = BOLD
    for i, th in enumerate([20, 30, 35, 40]):
        r = 8 + i
        _input(ws, f"A{r}", th)
        ws[f"B{r}"] = f'=COUNTIF({PS_ILLUM},">="&A{r})'
        ws[f"C{r}"] = f'=TEXT(B{r}/{PS_N - 1},"0.0%")'
        ws[f"B{r}"].font = BLUE

    _h2(ws, "B. 傾斜がしきい値より小さい（平らな）地点の数（黄色いセルに傾斜[度]）", 13)
    ws["A14"], ws["B14"], ws["C14"] = "傾斜 [度] 以下", "その地点数", "割合"
    for c in "ABC":
        ws[c + "14"].font = BOLD
    for i, th in enumerate([5, 8, 10, 15]):
        r = 15 + i
        _input(ws, f"A{r}", th)
        ws[f"B{r}"] = f'=COUNTIF({PS_SLOPE},"<="&A{r})'
        ws[f"C{r}"] = f'=TEXT(B{r}/{PS_N - 1},"0.0%")'
        ws[f"B{r}"].font = BLUE

    _h2(ws, ("C. 3つの条件を同時に満たす地点は何地点？" if student
             else "C. 「日当たりがよい」かつ「平ら」かつ「永久影のそば」は何地点？"), 20)
    _input(ws, "B21", 30)
    _input(ws, "B22", 10)
    _input(ws, "B23", 20)
    ws["A21"], ws["A22"], ws["A23"] = "日照率 ≧ [%]", "傾斜 ≦ [度]", "永久影まで ≦ [km]"
    for r in (21, 22, 23):
        ws[f"A{r}"].font = BODY
    ws["A24"] = "3つとも満たす地点数"
    ws["A24"].font = BOLD
    ws["B24"] = (f'=COUNTIFS({PS_ILLUM},">="&B21,{PS_SLOPE},"<="&B22,'
                 f'データ_南極!$E$2:$E${PS_N},"<="&B23)')
    ws["B24"].font = BLUE

    # ヒストグラム（傾斜）
    _h2(ws, "D. 傾斜のヒストグラム（2度きざみ）", 27)
    ws["A28"], ws["B28"] = "階級（以上）", "地点数"
    ws["A28"].font = ws["B28"].font = BOLD
    for i, lo in enumerate(range(0, 30, 2)):
        r = 29 + i
        ws[f"A{r}"] = lo
        ws[f"B{r}"] = f'=COUNTIFS({PS_SLOPE},">="&A{r},{PS_SLOPE},"<"&(A{r}+2))'
    bar = BarChart()
    bar.title = "南極の傾斜の分布"
    bar.y_axis.title = "地点数"
    bar.x_axis.title = "傾斜 [度] 以上"
    bar.add_data(Reference(ws, min_col=2, min_row=28, max_row=43), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=29, max_row=43))
    bar.height, bar.width = 8, 14
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.legend = None
    bar.title.overlay = False
    bar.series[0].graphicalProperties.solidFill = "1A6FB0"
    ws.add_chart(bar, "D27")
    _note(ws, "気づき：日当たりのよい地点と、平らな地点は『同じ場所』？　"
              "C で3つとも満たす地点はどれくらい残った？　"
              "いちばん平らな地点（傾斜が最小）を『データ_南極』で探すと、そこは日が当たる？", 46)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    _fit_notes(ws)

    # ================= ステップ4b：南極のスコア =================
    ws = wb.create_sheet(S4B)
    _title(ws, "ステップ4b：南極で、重みをつけて、いちばんよい場所を点数で決める")
    if student:
        _note(ws, "（水を1番にした班が使う）『データ_南極』の4つの指標を 0〜1 の点数に直して、"
                  "重みをつけて合計する（＝あなたのスコア式）。重みは黄色いセルで変える（0〜5の整数）。"
                  "地域名の入力（ステップ4の C5）は要らない。", 3)
    else:
        _note(ws, "（水を1番にした班向け。南極のシートを使う）『データ_南極』の4つの指標を 0〜1 の点数に直して、"
                  "重みをつけて合計する（＝あなたのスコア式）。重みは黄色いセルで変える。", 3)
    _h2(ws, "A. あなたのスコア式：重み（黄色いセル。0なら「気にしない」。0〜5の整数）", 6)
    ws["A7"], ws["B7"], ws["C7"] = "指標", "どちらが良い？", "重み"
    for c in "ABC":
        ws[c + "7"].font = BOLD
    weights = [
        ("日照率 (illum)", "高いほどよい", 1 if student else 3),
        ("永久影までの距離 (km_to_shadow)", "近いほどよい", 0),
        ("永久影率 (permanent_shadow_fraction)", "低いほどよい", 0),
        ("傾斜 (slope_deg)", "低いほどよい", 0 if student else 2),
    ]
    for i, (label, good, w) in enumerate(weights):
        r = 8 + i
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BODY
        ws[f"B{r}"] = good
        ws[f"B{r}"].font = BODY
        _input(ws, f"C{r}", w)
    _weight_validation(ws, "C8:C11")
    ws["A12"] = "重みの合計"
    ws["A12"].font = BOLD
    ws["B12"] = "=C8+C9+C10+C11"
    ws["B12"].font = BLUE
    # A2：状態表示セル（A13）
    sum_w4 = "SUM($C$8:$C$11)"
    ws["A13"] = (
        f'=IFERROR(IF({sum_w4}=0,"【注意】重みがすべて0です。どれかを1以上にしてください",'
        'IF(OR(MIN($C$8:$C$11)<0,MAX($C$8:$C$11)>5,SUMPRODUCT(--($C$8:$C$11<>INT($C$8:$C$11)))>0),'
        '"【注意】重みは0〜5の整数で入れてください",'
        f'"OK　採点対象 "&COUNT(データ_南極!$L$2:$L${PS_N})&" 地点　（重みの合計 "&{sum_w4}&"）")),'
        '"【注意】重みの欄に数字以外が入っています")')
    ws["A13"].font = BOLD
    _status_format(ws, "A13")

    if not student:
        # 水を1番にした班の例（教員版のみ）。電力＝日照率の行は1点。数値は polar_south_sites.csv から _top_4b で再計算
        _h2(ws, "B. 水を1番にした班の例（一例。上の黄色いセルに入れ直して使う。1位は再計算した値）", 14)
        for c, h in zip("ABCDEFG", ("設定（電力＝日照率は1点）", "日照", "永久影まで", "永久影率", "傾斜", "1位の緯度", "1位の経度")):
            ws[f"{c}15"] = h
            ws[f"{c}15"].font = BOLD
        import numpy as np
        ex4b = [("水を1番", (1, 3, 0, 0)), ("水を1番＋傾斜を2番", (1, 3, 0, 2)),
                ("（試行）電力を0にする", (0, 3, 0, 0)), ("（試行）電力を0＋傾斜を2番", (0, 3, 0, 2))]
        for i, (m, w) in enumerate(ex4b):
            r = 16 + i
            ws[f"A{r}"] = m
            ws[f"A{r}"].font = BODY
            for c, v in zip("BCDE", w):
                ws[f"{c}{r}"] = v
                ws[f"{c}{r}"].font = BODY
            ws[f"F{r}"], ws[f"G{r}"] = _top_4b(ps_df, w)[0]
            ws[f"F{r}"].font = ws[f"G{r}"].font = BODY
        s0 = sum(wi * ps_df[c].values for wi, c in zip((0, 3, 0, 0), NORM4B)) / 3
        n_tie = int((np.round(s0, 10) == np.round(s0, 10).max()).sum())
        _note(ws, f"水を1番にした班は、このシートを使う（4b で1番にできるのは水だけ）。傾斜（建設）は2番〔2点〕にしてよい。"
                  f"電力を0にした（0,3,0,0）は {n_tie} 地点が同点（スコア1.000）で、1位は行の順で決まっただけ＝試行にしない。", 20)

    _note(ws, "『データ_南極』シートの右端（L 列）に『スコア』列があり、A の重みで自動計算される"
              "（M 列は、同点の地点を行の順に並べるための列）。"
              "0〜1の点数（norm_illum など）は前処理済み（作り方は『参考』と build_course_data.py）。", 22)

    _h2(ws, ("B. スコアの高い順トップ10（重みを変えると入れ替わる）" if student
             else "C. スコアの高い順トップ10（重みを変えると入れ替わる）"), 24)
    PS_SCORE = f"データ_南極!$L$2:$L${PS_N}"
    PS_RANK = f"データ_南極!$M$2:$M${PS_N}"   # 順位用（同点を行番号で解く）
    ws["A25"], ws["B25"], ws["C25"], ws["D25"], ws["E25"], ws["F25"], ws["G25"] = (
        "順位", "スコア", "緯度", "経度", "日照率 [%]", "傾斜 [度]", "永久影まで [km]")
    for c in "ABCDEFG":
        ws[c + "25"].font = BOLD
    for k in range(1, 11):
        r = 25 + k
        ws[f"A{r}"] = k
        m = f'MATCH(LARGE({PS_RANK},A{r}),{PS_RANK},0)'
        ws[f"B{r}"] = f'=IFERROR(ROUND(INDEX({PS_SCORE},{m}),3),"")'
        ws[f"C{r}"] = f'=IFERROR(INDEX(データ_南極!$A$2:$A${PS_N},{m}),"")'
        ws[f"D{r}"] = f'=IFERROR(INDEX(データ_南極!$B$2:$B${PS_N},{m}),"")'
        ws[f"E{r}"] = f'=IFERROR(INDEX(データ_南極!$C$2:$C${PS_N},{m}),"")'
        ws[f"F{r}"] = f'=IFERROR(INDEX(データ_南極!$F$2:$F${PS_N},{m}),"")'
        ws[f"G{r}"] = f'=IFERROR(INDEX(データ_南極!$E$2:$E${PS_N},{m}),"")'
        for c in "BCDEFG":
            ws[f"{c}{r}"].font = Font(name="Yu Gothic", size=10, color="1A6FB0")
    # A4：上位10の範囲（日照率・傾斜・永久影までの距離）
    ws["A36"] = "上位10の範囲"
    ws["A36"].font = BOLD
    ws["B36"] = (
        '=IF(COUNT(E26:E35)=0,"",'
        '"日照率 "&TEXT(MIN(E26:E35),"0.0")&" 〜 "&TEXT(MAX(E26:E35),"0.0")&" %　／　"&'
        '"傾斜 "&TEXT(MIN(F26:F35),"0.0")&" 〜 "&TEXT(MAX(F26:F35),"0.0")&" °　／　"&'
        '"永久影まで "&TEXT(MIN(G26:G35),"0.0")&" 〜 "&TEXT(MAX(G26:G35),"0.0")&" km")')
    ws["B36"].font = BLUE
    if student:
        _note(ws, "気づき：トップ10の日照率・傾斜・永久影までの距離は、重みとどう対応している？　"
                  "重みを変えると順位はどう動いた？（変える前の結果は、紙に写しておく）", 37)
        for r in range(14, 21):      # 学習者版には教員版の『水を1番にした班の例』の表は作らない。行を隠す
            ws.row_dimensions[r].hidden = True
    else:
        _note(ws, "気づき：トップの場所は、どんな特徴（日照・傾斜・永久影までの距離）？　"
                  "日照を重くすると傾斜は？　両方を同時に満たす場所はあった？　重みを変えると順位はどう動く？", 37)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    for c in "CDEFG":
        ws.column_dimensions[c].width = 13
    _fit_notes(ws)

    # ================= ステップ5：まとめ =================
    ws = wb.create_sheet("ステップ5_まとめ")
    if student:
        # 学習者版では非表示の殻だけ残す（シート名は変えない）。実在計画の表は教員版で全体共有に使う
        _title(ws, "ステップ5：まとめ")
        _note(ws, "（配付版では内容を省いてある。全体共有で使う資料は教員が見せる）", 3)
    else:
        _title(ws, "ステップ5：各班の1番・使った範囲・1位と、実在の計画を見比べる")
        _note(ws, "各班が、1番に置いた要素・使った範囲（月全体／4b）とトップの場所（緯度・経度）を下の表に書き写す。"
                  "同じ月・同じデータなのに、1番に置いた要素によって答えが分かれたはず。", 3)
        ws["A6"], ws["B6"], ws["C6"], ws["D6"], ws["E6"] = (
            "班の選択（1番・2番）", "使った範囲（月全体／4b）", "トップの緯度", "トップの経度", "線 [°]")
        for c in "ABCDE":
            ws[c + "6"].font = BOLD
        for i, m in enumerate(["温度を1番にした班", "通信を1番にした班", "温度を1番・通信を2番にした班", "水を1番にした班"]):
            r = 7 + i
            ws[f"A{r}"] = m
            ws[f"A{r}"].font = BODY
            for c in "BCDE":
                _input(ws, f"{c}{r}", "")

        _h2(ws, "B. 実在の計画は、どの要素を重んじたか（出典は実施時に確認）", 12)
        ws["A13"], ws["B13"], ws["C13"], ws["D13"] = "計画", "場所", "重んじた要素・要件", "参考：地球の仰角"
        for c in "ABCD":
            ws[c + "13"].font = BOLD
        for i, (nm, where, req, ee) in enumerate([
            ("Artemis III（有人・氷）", "南極", "永久影の氷＋近くの日照尾根", "約 0°（地平線すれすれ）"),
            ("LCRT（月裏側電波望遠鏡・構想）", "裏側", "地球の電波が届かないこと", "負（地球が地平線の下）"),
            ("Apollo 11（赤道・実績）", "表側の海", "アボート容易・通信良好", "約 +66°（ほぼ真上）"),
            ("Chang'e 4/6（裏側・実績）", "裏側", "裏側の地質サンプル", "負"),
        ]):
            r = 14 + i
            ws[f"A{r}"] = nm
            ws[f"B{r}"] = where
            ws[f"C{r}"] = req
            ws[f"D{r}"] = ee
            for c in "ABCD":
                ws[f"{c}{r}"].font = BODY

        _h2(ws, "C. NASA Artemis III の南極候補地（『データ_着陸地点』より。氷・有人の着陸候補）", 20)
        ws["A21"], ws["B21"], ws["C21"], ws["D21"], ws["E21"] = (
            "候補地", "緯度", "経度", "傾斜 [度]", "日照率 [%]")
        for c in "ABCDE":
            ws[c + "21"].font = BOLD
        LSN = len(ls_df) + 1
        for i, nm in enumerate(["Artemis III: Malapert Massif", "Artemis III: Haworth",
                                "Artemis III: Nobile Rim 1", "Artemis III: de Gerlache Rim 2"]):
            r = 22 + i
            ws[f"A{r}"] = nm.replace("Artemis III: ", "")
            ws[f"A{r}"].font = BODY
            m = f'MATCH("{nm}",データ_着陸地点!$A$2:$A${LSN},0)'
            ws[f"B{r}"] = f'=INDEX(データ_着陸地点!$C$2:$C${LSN},{m})'
            ws[f"C{r}"] = f'=INDEX(データ_着陸地点!$D$2:$D${LSN},{m})'
            ws[f"D{r}"] = f'=INDEX(データ_着陸地点!$L$2:$L${LSN},{m})'
            ws[f"E{r}"] = f'=INDEX(データ_着陸地点!$M$2:$M${LSN},{m})'
            for c in "BCDE":
                ws[f"{c}{r}"].font = BLUE

        _note(ws, "考察：①1番に置いた要素が違うと、基地の場所はどう違った？　同じ1番でも別の場所を選んだ班は、何を見て選んだ？　"
                  "②1位が線のそばにあった班は、重みと線のどちらで場所が決まった？（線を5°動かして確かめる）　"
                  "③赤道に基地を置くなら、ステップ1の1日約290Kの較差にどう対処する？　"
                  "④このデータで『信じてよいか怪しいこと』は？（日照率・傾斜の絶対値／earth_elev_deg は秤動を無視／"
                  "temp_amp_K は極で不確か）", 27)
    ws.column_dimensions["A"].width = 28
    for c in "BCDEF":
        ws.column_dimensions[c].width = 14
    _fit_notes(ws)

    # 並び順
    order = ["はじめに", "ステップ1_温度", S1B, S1C, "ステップ2_海と陸", S3G, S4R,
             "ステップ4b_南極", S4B, "ステップ5_まとめ", "データ_温度", "データ_緯度行", "データ_地点比較", "データ_クレーター",
             "データ_クレーター年代", "データ_環境", "データ_地域", "データ_南極", "データ_北極",
             "データ_地質", "データ_着陸地点", "参考"]
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

    if student:
        for nm in STUDENT_HIDDEN:
            wb[nm].sheet_state = "hidden"
    # 開いたとき「はじめに」だけが選ばれた状態に
    wb.active = 0
    for w in wb.worksheets:
        w.sheet_view.tabSelected = (w.title == "はじめに")

    wb.save(out)
    print(f"wrote {out}  ({out.stat().st_size / 1024:.0f} KB, {len(wb.sheetnames)} シート)")
    return out


def main():
    ap = argparse.ArgumentParser(description="教員版と学習者版の xlsx を出す")
    ap.add_argument("--recalc", action="store_true",
                    help="学習者版を Excel で再計算して保存した版にする（Windows＋Excel が必要。配付・リポジトリ用）")
    a = ap.parse_args()
    build(student=False)
    out = build(student=True)
    if a.recalc:
        import recalc_xlsx_with_excel as rx
        tmp = out.with_name(out.stem + ".openpyxl.tmp")
        out.replace(tmp)
        try:
            rx.recalc(tmp, out)
        finally:
            tmp.unlink(missing_ok=True)
        print(f"recalculated {out}  ({out.stat().st_size / 1024:.0f} KB)  ※計算結果入り・個人情報除去済み")


if __name__ == "__main__":
    main()
