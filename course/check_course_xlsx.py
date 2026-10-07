# -*- coding: utf-8 -*-
"""表計算ブック（教員版・学習者版）の検査。

    python course/check_course_xlsx.py              # 答え語の検査＋構造の検査（Excel なしで動く）
    python course/check_course_xlsx.py --pandas     # ＋ pandas による式の再現（4バンド・密度比・USGS年代・
                                                    #    上位10の同点解消の網羅）
    python course/check_course_xlsx.py --excel      # ＋ Excel 実機（COM）での再計算・入力規則・上位10の網羅・
                                                    #    状態表示・範囲表示の検査（Windows＋Excel が必要）
    python course/check_course_xlsx.py --all        # すべて

1. 答え語の検査（要件 §4-X-2）
   学習者版の「全セルの文字列」（非表示シートを含む。数式の文字列も含む）と、xlsx 内のすべての XML
   （共有文字列・チャート・入力規則の文言・名前・文書情報）から、答え語リスト ANSWER_WORDS の出現を探す。
   学習者版で 0 件であること、**教員版にはリストの全語が存在する**こと（検査が空振りでないこと）を確認する。
2. 構造の検査（要件 §3.2）
   シート名と順序が両版で同じ／学習者版で非表示にしたシートの一覧／黄色い入力セルの位置と既定値／
   C5 のプルダウンと重みの入力規則／学習者版の データ_地域 が A〜E 列だけ、など。
3. 値の検査（--pandas / --excel）
   同じ式を pandas で再現し、Excel 実機の再計算結果と突き合わせる。

終了コード: すべて合格なら 0、1つでも不合格なら 1。
元のファイルは Excel で開かない（コピーを一時フォルダに作って開く）。
"""
import argparse
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

import numpy as np
import pandas as pd
from openpyxl import load_workbook

HERE = pathlib.Path(__file__).resolve().parent
DATA = HERE / "data"
TEACHER = HERE / "course_moonbase.xlsx"
STUDENT = HERE / "student" / "course_moonbase_student.xlsx"
DRIVER = HERE / "excel_driver.ps1"

# ---------------------------------------------------------------------------
# 答え語のリスト。教員版の実物のセルから拾った「答えにあたる文・語」と、要件 §4-X-2 の語。
# 学習者版では 0 件、教員版では全語が 1 件以上あること。
# ---------------------------------------------------------------------------
ANSWER_WORDS = [
    # 要件 §4-X-2 で挙げられた語
    "第一候補", "電波天文に理想", "氷採掘は南極", "見えたら失格", "裏側に", "永久影のそば",
    # はじめに（企画書 G1／指導案 §6.5）
    "電波天文台は裏側に", "通信重視なら表側に", "半球ごと変わります", "地球の電波が届かないこと",
    "（＝月の裏側）", "氷がありそうなこと", "ずっと日が当たらない永久影", "基地に必要なこと",
    "どこも『全部で一番』", "各ミッションの答え",
    # ステップ1
    "極付近（要注意）",
    # ステップ2（企画書 G1 の A33、G3 の D26）
    "海岸線の外側の陸も丸に入る", "落ちた後に消えた", "隕石が落ちなかった", "C の年代の平均が小さいほう",
    # ステップ3（企画書 G1）
    "正＝表側で通信できる", "負＝裏側で電波が静か", "裏側は地球の仰角が負", "南極は日較差が小さいが太陽高度は",
    # ステップ4（企画書 G1 の B8〜B12、指導案 §6.5 の A15〜C21・A37）
    "（発電）", "（熱の安定）", "（通信できる）", "（電波が静か・天文台）", "（夜に冷えすぎない）",
    "経度180度あたりが上位に来る", "ミッションごとの選び方", "2・2・0・0・1", "0・2・0・3・0",
    "1・1・3・0・0", "1・2・1・0・1", "永久影は環境データに無い", "赤道の海（静かの海） か 南極",
    # ステップ4b_スコア
    "氷採掘・南極を選んだ班向け", "ミッションごとの重み", "氷採掘基地", "太陽光発電基地", "有人基地",
    # ステップ5
    "氷採掘は南極で合意", "実在の計画は、ミッションで半球がちがう", "Artemis III（有人・氷）",
    # データ_地域 の rationale / caveat
    "永久影に氷があり", "電波雑音が届かない", "電波静穏度が最も高い", "LCRT（月裏側電波望遠鏡）構想の対象域",
    # 参考
    "（＝陸のほうが古い）",
]

STUDENT_HIDDEN = ["ステップ3_月全体", "ステップ4b_南極", "ステップ5_まとめ",
                  "データ_北極", "データ_地質", "データ_クレーター年代"]
SHEETS = ["はじめに", "ステップ1_温度", "ステップ2_海と陸", "ステップ3_月全体", "ステップ4_地域を選ぶ",
          "ステップ4b_南極", "ステップ4b_スコア", "ステップ5_まとめ", "データ_温度", "データ_クレーター",
          "データ_クレーター年代", "データ_環境", "データ_地域", "データ_南極", "データ_北極",
          "データ_地質", "データ_着陸地点", "参考"]
YELLOW = "FFF6E9"   # 黄色い入力セルの色（アルファ部分は openpyxl 出力＝00、Excel 保存＝FF と違うので比べない）
REG_FAR_EQ = "裏側・赤道（月の裏側の赤道帯）"
S1, S2, S4, S4B = "ステップ1_温度", "ステップ2_海と陸", "ステップ4_地域を選ぶ", "ステップ4b_スコア"

RESULTS = []   # (合否, 項目, 詳細)


def record(ok, name, detail=""):
    RESULTS.append((bool(ok), name, detail))
    print(("[OK]   " if ok else "[NG]   ") + name + (f"  {detail}" if detail else ""))


# ---------------------------------------------------------------------------
# 1. 答え語の検査
# ---------------------------------------------------------------------------
def scan_cells(path):
    """全シート（非表示を含む）の全セルの文字列（数式は式の文字列）→ {語: [場所,…]}"""
    wb = load_workbook(path, read_only=True)       # read_only でも非表示シートを含めて全部読める
    texts = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value:
                    texts.append((ws.title, c.coordinate, c.value))
    hits = {}
    for w in ANSWER_WORDS:
        for sh, cell, t in texts:
            if w in t:
                hits.setdefault(w, []).append(f"{sh}!{cell}")
    return hits, len(texts)


def scan_xml(path):
    """xlsx の全部品（共有文字列・シート・チャート・名前・入力規則・文書情報）を文字列として探す"""
    hits = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if not n.endswith((".xml", ".rels")):
                continue
            t = z.read(n).decode("utf-8", errors="replace")
            for w in ANSWER_WORDS:
                if w in t:
                    hits.setdefault(w, []).append(n)
    return hits


def check_answers(teacher, student):
    print("\n== 1. 答え語の検査 ==")
    print(f"答え語リスト: {len(ANSWER_WORDS)} 語")
    hs, n_s = scan_cells(student)
    hx = scan_xml(student)
    record(not hs, f"学習者版：全セル（{n_s} 個の文字列セル、非表示シート含む）に答え語 0 件",
           "" if not hs else "; ".join(f"「{w}」→{v[:3]}" for w, v in hs.items()))
    record(not hx, "学習者版：xlsx 内の全 XML（共有文字列・チャート・入力規則・名前・文書情報）に答え語 0 件",
           "" if not hx else "; ".join(f"「{w}」→{v[:3]}" for w, v in hx.items()))
    ht, n_t = scan_cells(teacher)
    missing = [w for w in ANSWER_WORDS if w not in ht]
    record(not missing, f"教員版：リストの全 {len(ANSWER_WORDS)} 語が存在（検査が空振りでない。{n_t} 個の文字列セル中）",
           "" if not missing else "教員版に無い語: " + "、".join(missing))
    # 空振りでないことの追加確認：検査関数に教員版を食わせて検出数を出す
    print(f"       教員版での検出：{sum(len(v) for v in ht.values())} 箇所／{len(ht)} 語")
    # 非表示シートを実際に走査していること（学習者版の非表示シートの文字列数を数える）
    wbs = load_workbook(student, read_only=True)
    hidden_cnt = sum(1 for ws in wbs.worksheets if ws.sheet_state != "visible")
    record(hidden_cnt == len(STUDENT_HIDDEN), f"学習者版の非表示シート数 {hidden_cnt}（走査対象に含まれる）")


# ミッション名・用途を示す語（要件：地域・半球とミッションの対応を学習者版に書かない）。
# 学習者版の全セル（非表示シート・数式の文字列を含む）と、xlsx 内の全部品（定義名・チャート・入力規則の文言・
# 文書情報など）に出たら不合格。出た場合は場所と前後の文脈を表示する。
MISSION_WORDS = ["電波", "氷採掘", "天文", "発電", "通信", "有人", "採掘", "太陽光"]
# 許容リスト（最小限）。「地域名・データ列の値として出る語」で、ミッションとの対応を示さないもの。
# 形式：(場所の前方一致, その語を含む文脈の部分文字列, 理由)。許容は1件ずつ理由を書く。
MISSION_ALLOWED = [
    ("データ_着陸地点!B", "着陸(有人)", "実在の着陸地点の種別（有人／無人）。史実の属性で、ミッションと地域の対応ではない"),
    ("データ_着陸地点!B", "着陸候補(有人)", "同上（Artemis の着陸候補地の種別）"),
]


def _ctx(text, i, n):
    return text[max(0, i - 14): i + n + 14].replace(chr(10), " ")


def mission_hits(path):
    """[(場所, 語, 文脈)] を返す（許容リストに当たるものは除く）。セルと、セル以外の XML 部品の両方を見る"""
    hits = []
    wb = load_workbook(path, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    for w in MISSION_WORDS:
                        i = c.value.find(w)
                        while i >= 0:
                            hits.append((f"{ws.title}!{c.coordinate}", w, _ctx(c.value, i, len(w))))
                            i = c.value.find(w, i + 1)
    cell_ctx = {(w, ctx) for _, w, ctx in hits}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if not n.endswith((".xml", ".rels")) or n == "xl/sharedStrings.xml":
                continue
            t = z.read(n).decode("utf-8", errors="replace")
            for w in MISSION_WORDS:
                i = t.find(w)
                while i >= 0:
                    ctx = _ctx(t, i, len(w))
                    # セルの値（数式を除く）はセルの走査で拾い済み。ここでは重複を省く
                    if (w, ctx) not in cell_ctx:
                        hits.append((n, w, ctx))
                    i = t.find(w, i + 1)
    out, allowed = [], []
    for where, w, ctx in hits:
        if any(where.startswith(a) and sub in ctx for a, sub, _ in MISSION_ALLOWED):
            allowed.append((where, w, ctx))
            continue
        out.append((where, w, ctx))
    return out, allowed


def check_mission_words(teacher, student):
    print(chr(10) + "== 1b. ミッション語の検査 ==")
    print("語：" + "・".join(MISSION_WORDS) + f"　許容リスト {len(MISSION_ALLOWED)} 件")
    hs, al = mission_hits(student)
    if al:
        print(f"       許容した出現 {len(al)} 件：" + "、".join(sorted({c for _, _, c in al})) + "（データ_着陸地点 の種別列）")
    record(not hs, "学習者版：全セル・全部品にミッション語 0 件",
           "" if not hs else f"{len(hs)} 件。" + " | ".join(f"{w}@{a}『{c}』" for a, w, c in hs[:8]))
    for a, w, c in hs[:30]:
        print(f"       学習者版 {a}  「{w}」 …{c}…")
    ht, _ = mission_hits(teacher)
    found = {w for _, w, _ in ht}
    record(found == set(MISSION_WORDS), "教員版：全 8 語が検出される（検査が空振りでない）",
           f"検出 {len(ht)} 件／語 {sorted(found)}" + ("" if found == set(MISSION_WORDS) else f" 未検出 {sorted(set(MISSION_WORDS) - found)}"))


# 地域名（データ_地域 A列）の列挙と、ミッション語・答えの示唆がないことの確認
def check_region_names(path, label):
    wb = load_workbook(path, read_only=True)
    ws = wb["データ_地域"]
    names = [r[0].value for r in ws.iter_rows(min_row=2, max_row=9, max_col=1)]
    bad = [n for n in names if any(w in n for w in MISSION_WORDS)]
    record(not bad and len(names) == 8, f"{label}：地域名 8 件にミッション語がない", " / ".join(names))


PII_PATTERNS = ["C:\\Users", "OneDrive", "@", "lastModifiedBy", "absPath"] + [os.environ.get("USERNAME", "")]
PII_PATTERNS = [w for w in PII_PATTERNS if w]


def check_pii(path, label):
    """xlsx 内の全部品に、利用者名・ローカルパス・メールの断片、Excel が埋める最終更新者・絶対パスがないこと"""
    found = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            t = z.read(n).decode("utf-8", errors="replace")
            for w in PII_PATTERNS:
                if w in t:
                    found.setdefault(w, []).append(n)
    record(not found, f"{label}：個人名・ローカルパス・最終更新者の記録が xlsx 内にない",
           "" if not found else "; ".join(f"{w}→{v[:2]}" for w, v in found.items()))


# ---------------------------------------------------------------------------
# 2. 構造の検査
# ---------------------------------------------------------------------------
def fill_of(ws, addr):
    c = ws[addr]
    return str(c.fill.fgColor.rgb)[-6:] if c.fill and c.fill.fill_type == "solid" else None


def check_structure(teacher, student):
    print("\n== 2. 構造の検査 ==")
    wt = load_workbook(teacher)
    ws_ = load_workbook(student)
    record(wt.sheetnames == SHEETS and ws_.sheetnames == SHEETS,
           "シート名・順序が両版で同じ（既存の18シート）")
    record(all(s.sheet_state == "visible" for s in wt.worksheets), "教員版：非表示シートなし")
    hid = [s.title for s in ws_.worksheets if s.sheet_state != "visible"]
    record(sorted(hid) == sorted(STUDENT_HIDDEN), "学習者版の非表示シート", "、".join(hid))

    # 黄色い入力セル（要件 §3.2：位置を動かさない）
    inputs = {S4: ["C5", "C8", "C9", "C10", "C11", "C12"], S4B: ["C8", "C9", "C10", "C11"],
              S1: ["A8", "B8", "A9", "B9", "A10", "B10", "A11", "B11", "B17", "B18"]}
    for label, wb in (("教員版", wt), ("学習者版", ws_)):
        bad = [f"{sh}!{a}" for sh, lst in inputs.items() for a in lst if fill_of(wb[sh], a) != YELLOW]
        record(not bad, f"{label}：黄色い入力セルの位置（ステップ4 C5・C8〜C12、4b C8〜C11、ステップ1）", "、".join(bad))
    # 既定値
    t4, t4b = wt[S4], wt[S4B]
    s4, s4b = ws_[S4], ws_[S4B]
    record(t4["C5"].value == "赤道の海（静かの海）" and [t4[f"C{r}"].value for r in range(8, 13)] == [2, 2, 0, 0, 0]
           and [t4b[f"C{r}"].value for r in range(8, 12)] == [3, 0, 0, 2],
           "教員版：既定値（C5＝赤道の海、重み 2・2・0・0・0／4b 3・0・0・2）は従来どおり")
    record(s4["C5"].value in (None, "") and all(s4[f"C{r}"].value == 0 for r in range(8, 13))
           and all(s4b[f"C{r}"].value == 0 for r in range(8, 12)),
           "学習者版：C5 空欄、重みはすべて 0（ステップ4・4b）")

    # 入力規則
    for label, wb in (("教員版", wt), ("学習者版", ws_)):
        dvs = wb[S4].data_validations.dataValidation
        lst = [d for d in dvs if d.type == "list" and "C5" in str(d.sqref)]
        whole = [d for d in dvs if d.type == "whole" and "C8:C12" in str(d.sqref)
                 and d.formula1 == "0" and d.formula2 == "5"]
        dvs_b = wb[S4B].data_validations.dataValidation
        whole_b = [d for d in dvs_b if d.type == "whole" and "C8:C11" in str(d.sqref)
                   and d.formula1 == "0" and d.formula2 == "5"]
        ref = wb.defined_names.get("RegionNames")
        record(len(lst) == 1 and lst[0].formula1 == "RegionNames" and ref is not None
               and ref.attr_text == "データ_地域!$A$2:$A$9",
               f"{label}：C5 のプルダウン（名前 RegionNames＝データ_地域!A2:A9）")
        record(len(whole) == 1 and len(whole_b) == 1,
               f"{label}：重みの入力規則（ステップ4 C8:C12、4b C8:C11＝0〜5 の整数）")
    # 状態表示・範囲表示セルの存在
    for label, wb in (("教員版", wt), ("学習者版", ws_)):
        a6 = str(wb[S4]["A6"].value)
        a13 = str(wb[S4B]["A13"].value)
        b36 = str(wb[S4]["B36"].value)
        b36b = str(wb[S4B]["B36"].value)
        record(a6.startswith("=") and "南極" in a6 and a13.startswith("=") and b36.startswith("=")
               and "日付変更線" in b36 and b36b.startswith("=") and "日照率" in b36b,
               f"{label}：状態表示（ステップ4 A6・4b A13）と範囲表示（B36）の式がある")
    # 順位用列（同点解消）
    for label, wb in (("教員版", wt), ("学習者版", ws_)):
        e = wb["データ_環境"]
        p = wb["データ_南極"]
        record(str(e["P2"].value).startswith("=IF(O2") and "ROW()" in str(e["P2"].value)
               and "ROW()" in str(p["M2"].value) and e["O1"].value == "スコア" and p["L1"].value == "スコア",
               f"{label}：順位用列（データ_環境 P、データ_南極 M）に行番号のタイブレークがある")
        f = str(wb[S4]["B26"].value)
        record("LARGE(データ_環境!$P$2" in f and "INDEX(データ_環境!$O$2" in f,
               f"{label}：上位10は順位用列で引き、表示のスコアは元のスコア列（見た目は変わらない）")
    # 学習者版の データ_地域
    dr = ws_["データ_地域"]
    record(dr.max_column == 5 and dr["A1"].value == "name" and dr["E1"].value == "lon_max",
           "学習者版：データ_地域は A〜E 列だけ（rationale・caveat なし）", f"max_column={dr.max_column}")
    record(wt["データ_地域"].max_column == 7, "教員版：データ_地域は A〜G 列（rationale・caveat あり）")
    # 地域名とプルダウン一覧の一致
    names = [dr[f"A{r}"].value for r in range(2, 10)]
    env_regions = set(pd.read_csv(DATA / "env_grid.csv")["region"].dropna())
    record(set(names) == env_regions, "データ_地域 A2:A9 の名前が データ_環境 の region 列の値と一致", f"{len(names)} 件")
    # 学習者版の「はじめに」が指導案の流れと合うこと（使うシートの記載）
    intro = " ".join(str(c.value) for row in ws_["はじめに"].iter_rows() for c in row if c.value)
    need = ["ステップ1_温度", "ステップ2_海と陸", "ステップ4_地域を選ぶ", "ステップ4b_スコア", "上書き保存しない", "編集を有効にする"]
    record(all(n in intro for n in need) and "ステップ3" not in intro and "ステップ5" not in intro,
           "学習者版：『はじめに』は使うシート（1・2・4・4b_スコア）と操作の決まりだけを案内し、使わないシートを挙げない")
    # 条件付き書式・チャート（Excel 保存版でも保たれていること）
    for label, wb in (("教員版", wt), ("学習者版", ws_)):
        cf4 = [str(r.sqref) for r in wb[S4].conditional_formatting]
        cf4b = [str(r.sqref) for r in wb[S4B].conditional_formatting]
        record("A6" in cf4 and "A13" in cf4b, f"{label}：状態表示の条件付き書式（ステップ4 A6、4b A13）")
        record(len(wb[S1]._charts) == 1 and len(wb["ステップ4b_南極"]._charts) == 1,
               f"{label}：チャート（ステップ1・4b_南極 に各1個）")
    record("指示されたところの黄色いセルだけ" in intro and "4つの緯度帯" in intro,
           "学習者版：『はじめに』は「指示されたところの黄色いセルだけ」「ステップ1の4つの緯度帯は変えない」と書いている")
    h = " ".join(str(c.value) for row in ws_[S1]["A6:A6"] for c in row)
    record("変えずに" in h, "学習者版：ステップ1の見出し A6 は黄色い緯度帯を変えない案内（教員版は従来どおり）")
    # 文書情報に個人名がない
    with zipfile.ZipFile(student) as z:
        core = z.read("docProps/core.xml").decode("utf-8")
    with zipfile.ZipFile(teacher) as z:
        core_t = z.read("docProps/core.xml").decode("utf-8")
    ok_meta = all("lastModifiedBy" not in x and not re.search(r"<dc:creator>[^<]+</dc:creator>", x) for x in (core, core_t))
    record(ok_meta, "両版：文書情報に最終更新者・作成者がない（空）")


# ---------------------------------------------------------------------------
# 3. pandas による式の再現
# ---------------------------------------------------------------------------
NORM4 = ["norm_sun_high", "norm_amp_low", "norm_earth_high", "norm_earth_low", "norm_night_warm"]
NORM4B = ["norm_illum", "norm_near_shadow", "norm_low_psf", "norm_low_slope"]
EPS = 1e-12
STD4 = {  # 指導案 §8 の標準の重み（ステップ4）
    "電波天文（裏側・赤道）": (REG_FAR_EQ, (0, 2, 0, 3, 0)),
    "太陽光（赤道の海）": ("赤道の海（静かの海）", (2, 2, 0, 0, 1)),
    "有人総合（赤道の海）": ("赤道の海（静かの海）", (1, 2, 1, 0, 1)),
}
STD4B = {"氷採掘（4b）": (0, 3, 0, 2), "太陽光（4b）": (3, 0, 0, 2), "有人（4b）": (2, 2, 1, 2)}


class Expect:
    """pandas で表計算の式を再現する"""

    def __init__(self):
        self.env = pd.read_csv(DATA / "env_grid.csv")
        self.env["row"] = np.arange(2, len(self.env) + 2)
        self.ps = pd.read_csv(DATA / "polar_south_sites.csv")
        self.ps["row"] = np.arange(2, len(self.ps) + 2)
        self.temp = pd.read_csv(DATA / "temp_grid.csv")
        self.cr = pd.read_csv(DATA / "craters_labeled.csv")
        self.ref = pd.read_csv(DATA / "reference.csv")
        self.regions = list(pd.read_csv(DATA / "candidate_regions.csv")["name"])

    # --- ステップ1・2 ---
    def bands(self):
        out = []
        for lo, hi in [(-6, 6), (24, 36), (54, 66), (78, 90)]:
            m = (self.temp["lat"] >= lo) & (self.temp["lat"] <= hi)
            out.append(int(round(self.temp.loc[m, "t_swing_K"].mean())))
        return out

    def density(self):
        ns = int((self.cr["区分"] == "海").sum())
        nl = int((self.cr["区分"] == "陸").sum())
        fs = float(self.ref.loc[0, "値"])
        fl = float(self.ref.loc[1, "値"])
        ds, dl = round(ns / fs), round(nl / fl)
        return ns, nl, ds, dl, round(dl / ds, 1), dl / ds

    def usgs(self):
        return float(self.ref.loc[4, "値"]), float(self.ref.loc[5, "値"])

    # --- ステップ4 ---
    def score4(self, region, w):
        e = self.env
        m = (e["region"] == region).values
        sc = np.full(len(e), np.nan)
        if sum(w) > 0:
            s = np.zeros(m.sum())
            for wi, col in zip(w, NORM4):          # Excel の式と同じ順に足す
                s = s + wi * e.loc[m, col].values
            sc[m] = s / max(1, sum(w))
        return sc

    def top10_old(self, sc, rows, k=10):
        """変更前の方式（LARGE＋MATCH(…,0)）：同じ値は最初の行を返す"""
        vals = sorted(sc[~np.isnan(sc)], reverse=True)[:k]
        out = []
        for v in vals:
            out.append(int(rows[np.where(sc == v)[0][0]]))
        return out

    def top10_new(self, sc, rows, k=10):
        """変更後：順位用列 ROUND(スコア,10)−ROW()×1E-12 で LARGE＋MATCH"""
        p = np.where(np.isnan(sc), np.nan, np.round(sc, 10) - rows * EPS)
        ok = np.where(~np.isnan(p))[0]
        order = ok[np.argsort(-p[ok], kind="stable")][:k]
        return [int(rows[i]) for i in order]

    def top4(self, region, w, which="new"):
        sc = self.score4(region, w)
        rows = self.env["row"].values
        f = self.top10_new if which == "new" else self.top10_old
        r = f(sc, rows)
        e = self.env.set_index("row")
        return r, [(float(e.loc[x, "lat"]), float(e.loc[x, "lon"])) for x in r], sc

    # --- ステップ4b ---
    def score4b(self, w):
        p = self.ps
        sc = np.full(len(p), np.nan)
        s = np.zeros(len(p))
        for wi, col in zip(w, NORM4B):
            s = s + wi * p[col].values
        sc[:] = s / max(1, sum(w))
        return sc

    def top4b(self, w, which="new"):
        sc = self.score4b(w)
        rows = self.ps["row"].values
        if sum(w) == 0:
            return [], [], sc
        f = self.top10_new if which == "new" else self.top10_old
        r = f(sc, rows)
        e = self.ps.set_index("row")
        return r, [(float(e.loc[x, "lat"]), float(e.loc[x, "lon"])) for x in r], sc


def fmt1(x):
    """Excel の TEXT(x,"0.0") と同じ丸め（四捨五入＝0.5 は切り上げ。表示された小数どおりに丸める）"""
    from decimal import Decimal, ROUND_HALF_UP
    return str(Decimal(repr(float(x))).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))


def range_text4(pts):
    lat = [p[0] for p in pts]
    lon = [p[1] for p in pts]
    s = f"緯度 {fmt1(min(lat))} 〜 {fmt1(max(lat))} °　／　"
    if max(lon) - min(lon) > 180:
        return s + "経度は日付変更線をまたぐので範囲を出しません"
    return s + f"経度 {fmt1(min(lon))} 〜 {fmt1(max(lon))} °"


def range_text4b(ex, rows):
    e = ex.ps.set_index("row").loc[rows]
    a, b, c = e["average_illumination_percent"], e["slope_deg"], e["km_to_shadow"]
    return (f"日照率 {fmt1(a.min())} 〜 {fmt1(a.max())} %　／　傾斜 {fmt1(b.min())} 〜 {fmt1(b.max())} °　／　"
            f"永久影まで {fmt1(c.min())} 〜 {fmt1(c.max())} km")


def check_pandas(ex):
    print("\n== 3a. pandas による式の再現 ==")
    b = ex.bands()
    record(b == [295, 280, 236, 161], f"ステップ1の4バンドの日較差の平均 {b}（期待 295・280・236・161 K）")
    ns, nl, ds, dl, r1, r = ex.density()
    record((ns, nl) == (2232, 34145) and abs(r - 3.52) < 0.005,
           f"ステップ2：クレーター数 海{ns}・陸{nl}、密度 海{ds}・陸{dl}、倍率 {r:.3f}（表示は小数1桁で {r1}）")
    us, ul = ex.usgs()
    record(round(us, 2) == 3.16 and round(ul, 2) == 2.38, f"USGS の相対年代 海{us}→{round(us, 2)}・陸{ul}→{round(ul, 2)}")

    # --- 同点重複：全地域×単一指標（ステップ4） ---
    print("   [全地域×単一指標（重み 3）の上位10：異なる行数（変更前→変更後）]")
    bad_new, bad_old, same1 = [], 0, True
    names4 = ["太陽高度", "日較差", "仰角表", "仰角裏", "夜の最低"]
    ncases = 0
    for reg in ex.regions:
        n_rows = int((ex.env["region"] == reg).sum())
        for i in range(5):
            w = [0] * 5
            w[i] = 3
            sc = ex.score4(reg, w)
            rows = ex.env["row"].values
            old = ex.top10_old(sc, rows)
            new = ex.top10_new(sc, rows)
            want = min(10, n_rows)
            ncases += 1
            if len(set(old)) < want:
                bad_old += 1
            if len(set(new)) != want:
                bad_new.append((reg, names4[i], len(set(new))))
            if old[0] != new[0]:
                same1 = False
    record(not bad_new, f"ステップ4：{ncases} 通り（8地域×5指標）すべてで上位10が全部異なる行（変更後）",
           f"変更前に重複した通り数 {bad_old}／{ncases}" + (f"　不合格 {bad_new}" if bad_new else ""))
    record(same1, "単一指標の全 40 通りで、1位は変更前と同じ行")
    # 4b 単一指標
    bad_b, bad_ob = [], 0
    names4b = ["日照", "永久影まで", "永久影率", "傾斜"]
    for i in range(4):
        w = [0] * 4
        w[i] = 3
        sc = ex.score4b(w)
        rows = ex.ps["row"].values
        old = ex.top10_old(sc, rows)
        new = ex.top10_new(sc, rows)
        if len(set(old)) < 10:
            bad_ob += 1
        if len(set(new)) != 10:
            bad_b.append((names4b[i], len(set(new))))
        if old[0] != new[0]:
            bad_b.append(("1位が変わった", names4b[i]))
    record(not bad_b, "4b：単一指標 4 通りで上位10が全部異なる行（変更後）", f"変更前に重複した通り数 {bad_ob}／4")
    # 赤道の海（3,0,0,0,0）と 4b（0,3,0,0）（0,0,3,0）
    for lab, f in (("赤道の海 (3,0,0,0,0)", lambda: ex.top4("赤道の海（静かの海）", [3, 0, 0, 0, 0], "old")),):
        r, pts, sc = f()
        record(len(set(r)) < 10, f"{lab}：変更前の方式では上位10が {len(set(r))} 行しかない（不具合の再現）")
    for lab, w in (("4b (0,3,0,0)", [0, 3, 0, 0]), ("4b (0,0,3,0)", [0, 0, 3, 0])):
        r, pts, sc = ex.top4b(w, "old")
        record(len(set(r)) < 10, f"{lab}：変更前の方式では上位10が {len(set(r))} 行しかない（不具合の再現）")

    # --- 標準の重み ---
    print("   [標準の重み（指導案 §8）]")
    for lab, (reg, w) in STD4.items():
        ro, po, _ = ex.top4(reg, list(w), "old")
        rn, pn, _ = ex.top4(reg, list(w), "new")
        record(ro[0] == rn[0], f"{lab}：1位は変更前と同じ地点 {pn[0]}", f"変更前の1位 {po[0]}")
    for lab, w in STD4B.items():
        ro, po, _ = ex.top4b(list(w), "old")
        rn, pn, _ = ex.top4b(list(w), "new")
        record(ro[0] == rn[0], f"{lab}：1位は変更前と同じ地点 {pn[0]}", f"変更前の1位 {po[0]}")
    # --- A4 の期待範囲 ---
    reg, w = STD4["電波天文（裏側・赤道）"]
    rn, pn, _ = ex.top4(reg, list(w))
    lat = [p[0] for p in pn]
    lon = [p[1] for p in pn]
    record(max(abs(x) for x in lat) <= 4.5 and min(lon) >= 175.5 and max(lon) <= 178.5,
           f"A4 電波天文：上位10は緯度±4.5°以内・経度175.5〜178.5°（実測 緯度{min(lat)}〜{max(lat)}、経度{min(lon)}〜{max(lon)}）",
           range_text4(pn))
    rn, pn, _ = ex.top4b(list(STD4B["氷採掘（4b）"]))
    e = ex.ps.set_index("row").loc[rn]
    record((e["average_illumination_percent"] == 0).all() and (e["km_to_shadow"] == 0).all(),
           "A4 氷採掘（4b）：上位10は日照 0%・永久影までの距離 0 km", range_text4b(ex, rn))


# ---------------------------------------------------------------------------
# 3b. Excel 実機（COM）
# ---------------------------------------------------------------------------
def run_excel(job_steps, src, workdir, saveas=None):
    work = pathlib.Path(workdir)
    work.mkdir(parents=True, exist_ok=True)
    copy = work / ("copy_" + pathlib.Path(src).name)
    shutil.copyfile(src, copy)               # 元ファイルは開かない
    job = {"file": str(copy), "out": str(work / "out.json"), "steps": job_steps}
    jp = work / "job.json"
    jp.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(DRIVER),
                        "-JobFile", str(jp)], capture_output=True, text=True, timeout=1500)
    if r.returncode != 0:
        raise RuntimeError("Excel ドライバが失敗: " + (r.stderr or r.stdout)[:1000])
    return json.loads((work / "out.json").read_text(encoding="utf-8"))


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def top_from(grid):
    """C26:D35 の2列 → [(緯度, 経度)…]（空白は除く）"""
    return [(num(r[0]), num(r[1])) for r in grid if num(r[0]) is not None]


def check_excel(ex, path, label, workdir, student):
    print(f"\n== 3b. Excel 実機（{label}） ==")
    steps = [{"op": "info", "key": "info"}, {"op": "calc", "key": "first"}]
    g = lambda sh, addr, key, text=False: steps.append({"op": "get", "sheet": sh, "addr": addr, "key": key, "text": text})
    g(S1, "A8:D11", "bands")
    g(S2, "A7:B13", "s2")
    g(S2, "A26:B28", "s2d")
    g("ステップ4_地域を選ぶ", "A6", "st_init", True)
    steps.append({"op": "dv", "sheet": S4, "addr": "C5", "key": "dv_c5"})
    steps.append({"op": "dv", "sheet": S4, "addr": "C8", "key": "dv_w4"})
    steps.append({"op": "dv", "sheet": S4B, "addr": "C8", "key": "dv_w4b"})
    # 入力規則の動作（拒否されるか）
    tests_w = [("w_ok3", 3), ("w_ok0", 0), ("w_ok5", 5), ("w_ng-1", -1), ("w_ng9", 9), ("w_ng2.5", 2.5), ("w_ngabc", "abc")]
    for k, v in tests_w:
        steps.append({"op": "testvalid", "sheet": S4, "addr": "C8", "value": v, "key": k})
        steps.append({"op": "testvalid", "sheet": S4B, "addr": "C8", "value": v, "key": k + "_b"})
    for i, nm in enumerate(ex.regions):
        steps.append({"op": "testvalid", "sheet": S4, "addr": "C5", "value": nm, "key": f"c5_ok{i}"})
    steps.append({"op": "testvalid", "sheet": S4, "addr": "C5", "value": "でたらめ", "key": "c5_ng"})
    steps.append({"op": "testvalid", "sheet": S4, "addr": "C5", "value": "赤道の海", "key": "c5_ng2"})

    cases = []   # (key, 種別, 地域, 重み)
    # ステップ4：全地域×単一指標（重み3）
    for ri, reg in enumerate(ex.regions):
        for i in range(5):
            w = [0] * 5
            w[i] = 3
            cases.append((f"s4_{ri}_{i}", 4, reg, w))
    # 標準の重みと、重みが全部0／1の例
    for lab, (reg, w) in STD4.items():
        cases.append((f"std4_{lab}", 4, reg, list(w)))
    cases.append(("s4_all1", 4, REG_FAR_EQ, [1, 1, 1, 1, 1]))
    cases.append(("s4_zero", 4, "赤道の海（静かの海）", [0, 0, 0, 0, 0]))
    # 4b
    for i in range(4):
        w = [0] * 4
        w[i] = 3
        cases.append((f"s4b_single{i}", "4b", None, w))
    for lab, w in STD4B.items():
        cases.append((f"std4b_{lab}", "4b", None, list(w)))
    cases.append(("s4b_zero", "4b", None, [0, 0, 0, 0]))
    cases.append(("s4b_over", "4b", None, [3, 0, 0, 9]))

    for key, kind, reg, w in cases:
        if kind == 4:
            steps.append({"op": "set", "sheet": S4, "addr": "C5", "value": reg})
            for j, v in enumerate(w):
                steps.append({"op": "set", "sheet": S4, "addr": f"C{8 + j}", "value": v})
            steps.append({"op": "calc", "key": key})
            g(S4, "C26:D35", key + "_top")
            g(S4, "B26:B35", key + "_score", True)
            g(S4, "A6", key + "_st", True)
            g(S4, "B36", key + "_rng", True)
        else:
            for j, v in enumerate(w):
                steps.append({"op": "set", "sheet": S4B, "addr": f"C{8 + j}", "value": v})
            steps.append({"op": "calc", "key": key})
            g(S4B, "C26:D35", key + "_top")
            g(S4B, "A13", key + "_st", True)
            g(S4B, "B36", key + "_rng", True)
    # 各地域を C5 に入れたときの状態表示（重みは標準）
    for ri, reg in enumerate(ex.regions):
        steps.append({"op": "set", "sheet": S4, "addr": "C5", "value": reg})
        for j, v in enumerate([2, 2, 0, 0, 0]):
            steps.append({"op": "set", "sheet": S4, "addr": f"C{8 + j}", "value": v})
        steps.append({"op": "calc", "key": f"reg{ri}"})
        g(S4, "A6", f"reg{ri}_st", True)
        g(S4, "C26:C35", f"reg{ri}_lat")
    for key, val in (("c5_blank", None), ("c5_bad", "でたらめ")):
        steps.append({"op": "set", "sheet": S4, "addr": "C5", "value": val})
        steps.append({"op": "calc", "key": key})
        g(S4, "A6", key + "_st", True)
        g(S4, "C26:C35", key + "_lat", True)
    steps.append({"op": "chart", "sheet": S1, "key": "chart1"})
    steps.append({"op": "chart", "sheet": "ステップ4b_南極", "key": "chart2"})

    out = run_excel(steps, path, workdir)
    print(f"   Excel {out.get('excel_version')}：開く {out['open_ms']} ms、初回再計算 {out.get('calc_ms_first')} ms、"
          f"検査全体 {out['total_ms'] / 1000:.0f} 秒")

    errs = out.get("errors") or []
    record(not errs, f"{label}：Excel の操作（値の入力・再計算・読み取り）にエラーなし", "" if not errs else str(errs[:3]))
    # --- 開いて計算できたか／シート ---
    vis = {n: v for n, v, _ in out["info"]["sheets"]}
    want_hidden = set(STUDENT_HIDDEN) if student else set()
    got_hidden = {n for n, v in vis.items() if v != -1}
    record(got_hidden == want_hidden, f"{label}：Excel で開けて、非表示シートが期待どおり", "、".join(sorted(got_hidden)))
    # --- 値 ---
    bands = [int(r[3]) for r in out["bands"]]
    record(bands == ex.bands() == [295, 280, 236, 161], f"{label}：ステップ1の4バンド（Excel）{bands}")
    s2 = {r[0]: r[1] for r in out["s2"]}
    ns, nl, ds, dl, r1, r = ex.density()
    vals = [x[1] for x in out["s2"]]
    record(vals[0] == ns and vals[1] == nl and vals[4] == ds and vals[5] == dl and abs(vals[6] - r1) < 1e-9,
           f"{label}：ステップ2（Excel）海{vals[0]}・陸{vals[1]}・密度 {vals[4]}/{vals[5]}・倍率 {vals[6]}（小数1桁表示。厳密な比 {r:.3f}）")
    us, ul = ex.usgs()
    d = out["s2d"]
    record(abs(d[1][1] - us) < 1e-9 and abs(d[2][1] - ul) < 1e-9,
           f"{label}：USGS 年代（Excel）海{round(d[1][1], 2)}・陸{round(d[2][1], 2)}")
    # --- 入力規則 ---
    dv = out["dv_c5"]
    record(dv.get("type") == 3 and dv.get("formula1") == "=RegionNames" and dv.get("in_cell_dropdown") is True
           and dv.get("alert_style") == 1 and dv.get("show_error") is True,
           f"{label}：C5 のプルダウン（Excel：リスト、数式 {dv.get('formula1')}、停止メッセージ）")
    ok_c5 = all(out[f"c5_ok{i}"] is True for i in range(len(ex.regions)))
    record(ok_c5 and out["c5_ng"] is False and out["c5_ng2"] is False,
           f"{label}：C5 の入力規則：8 地域名は受理、「でたらめ」「赤道の海」（名前の一部）は拒否")
    for lab, suf in (("ステップ4 C8", ""), ("4b C8", "_b")):
        okv = [out[f"w_ok{n}{suf}"] for n in (3, 0, 5)]
        ngv = [out[f"w_ng{n}{suf}"] for n in ("-1", "9", "2.5", "abc")]
        record(all(v is True for v in okv) and all(v is False for v in ngv),
               f"{label}：重み（{lab}）の入力規則：0・3・5 は受理、-1・9・2.5・abc は拒否")
    wdv = out["dv_w4"]
    record(wdv.get("type") == 1 and wdv.get("formula1") == "0" and wdv.get("formula2") == "5" and wdv.get("operator") == 1,
           f"{label}：重みの入力規則（Excel：整数、0〜5、{wdv.get('formula1')}〜{wdv.get('formula2')}）")
    # --- 上位10：pandas と一致、10行が全部異なる ---
    bad_dist, bad_eq = [], []
    n_dist = 0
    for key, kind, reg, w in cases:
        if kind == 4:
            rows, pts, sc = ex.top4(reg, w)
            want = min(10, int((ex.env["region"] == reg).sum())) if sum(w) > 0 else 0
        else:
            if sum(w) == 0:
                rows, pts = [], []
            else:
                rows, pts, sc = ex.top4b(w)
            want = 10 if sum(w) > 0 else 0
        got = top_from(out[key + "_top"])
        if kind == "4b" and key == "s4b_over":
            continue  # 重み 9 は入力規則の外。状態表示の検査で扱う
        if len(set(got)) != len(got) or len(got) != want:
            bad_dist.append((key, len(got), len(set(got))))
        if got != pts:
            bad_eq.append(key)
        if key.startswith("s4_") or key.startswith("s4b_single"):
            n_dist += 1
    record(not bad_dist, f"{label}：Excel の上位10が全部異なる地点（全地域×単一指標 40＋4b 単一指標 4＋標準の重み ほか）",
           "" if not bad_dist else str(bad_dist[:5]))
    record(not bad_eq, f"{label}：Excel の上位10（緯度・経度、全 {len(cases) - 1} 通り）が pandas の再現と一致",
           "" if not bad_eq else str(bad_eq[:8]))
    # 標準の重みの1位
    for lab, (reg, w) in STD4.items():
        got = top_from(out[f"std4_{lab}_top"])
        ro, po, _ = ex.top4(reg, list(w), "old")
        record(got[0] == po[0], f"{label}：{lab} の1位（Excel）{got[0]}は変更前の1位 {po[0]} と同じ")
    for lab, w in STD4B.items():
        got = top_from(out[f"std4b_{lab}_top"])
        ro, po, _ = ex.top4b(list(w), "old")
        record(got[0] == po[0], f"{label}：{lab} の1位（Excel）{got[0]}は変更前の1位 {po[0]} と同じ")
    # --- 表示のスコアの見た目 ---
    sc_rows = out["std4_電波天文（裏側・赤道）_score"]
    reg, w = STD4["電波天文（裏側・赤道）"]
    _, _, sc = ex.top4(reg, list(w))
    rows, _, _ = ex.top4(reg, list(w))
    exp_scores = [f"{round(sc[x - 2], 3):.3f}".rstrip("0").rstrip(".") for x in rows]
    got_scores = [r[0] for r in sc_rows]
    ok_sc = all(abs(float(a) - round(sc[x - 2], 3)) < 1e-9 for a, x in zip(got_scores, rows))
    record(ok_sc, f"{label}：表示のスコアは ROUND(元のスコア,3) と一致（見た目は変わらない）", "/".join(got_scores[:5]) + "…")
    # --- A4 範囲表示 ---
    for lab, (reg, w) in STD4.items():
        rows, pts, _ = ex.top4(reg, list(w))
        txt = out[f"std4_{lab}_rng"][0][0]
        record(txt == range_text4(pts), f"{label}：A4 範囲表示 {lab}", txt)
    for lab, w in STD4B.items():
        rows, pts, _ = ex.top4b(list(w))
        txt = out[f"std4b_{lab}_rng"][0][0]
        record(txt == range_text4b(ex, rows), f"{label}：A4 範囲表示 {lab}", txt)
    # 日付変更線をまたぐ領域（南極エイトケン・南極）は経度の範囲を出さない
    key = "s4_%d_%d" % (ex.regions.index("裏側・南極エイトケン（フォン・カルマン）"), 0)
    # --- 状態表示 ---
    st = lambda k: out[k + "_st"][0][0]
    record(st("c5_blank").startswith("地域タイプを選んでください") and all(r[0] == "" for r in out["c5_blank_lat"]),
           f"{label}：C5 が空欄 → 「{st('c5_blank')}」、上位10は空白")
    record(st("c5_bad").startswith("【注意】C5 の地域名が") and all(r[0] == "" for r in out["c5_bad_lat"]),
           f"{label}：C5 が不一致 → 「{st('c5_bad')}」、上位10は空白")
    record(st("s4_zero").startswith("【注意】重みがすべて0"), f"{label}：ステップ4 重みが全部0 → 「{st('s4_zero')}」")
    record(st("s4b_zero").startswith("【注意】重みがすべて0"), f"{label}：4b 重みが全部0 → 「{st('s4b_zero')}」")
    record(st("s4b_over").startswith("【注意】重みは0〜5の整数"), f"{label}：4b 重みが範囲外（貼り付けで 9）→ 「{st('s4b_over')}」")
    ok_regs = []
    for ri, nm in enumerate(ex.regions):
        t = out[f"reg{ri}_st"][0][0]
        n = int((ex.env["region"] == nm).sum())
        if nm.startswith("南極"):
            ok_regs.append(t.startswith("南極は『ステップ4b_スコア』を使います"))
        else:
            ok_regs.append(t.startswith(f"OK　採点対象 {n} 行"))
    record(all(ok_regs), f"{label}：8 地域を C5 に入れたときの状態表示（南極は「南極は『ステップ4b_スコア』を使います…」、他は OK＋行数）",
           f"南極：{out['reg6_st'][0][0]}")
    # 日付変更線
    key = f"s4_{ex.regions.index('裏側・南極エイトケン（フォン・カルマン）')}_3"
    rng = out[key + "_rng"][0][0]
    rows, pts, _ = ex.top4("裏側・南極エイトケン（フォン・カルマン）", [0, 0, 0, 3, 0])
    record(rng == range_text4(pts), f"{label}：日付変更線をまたぐ領域の範囲表示：{rng}")
    # --- G10 チャートの位置 ---
    ch = out["chart1"][0]
    record(ch[1] == "$D$21", f"{label}：ステップ1のチャートは D21 起点（注意書き A19:H19 に重ならない）", f"{ch[1]}〜{ch[2]}")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--teacher", default=str(TEACHER))
    ap.add_argument("--student", default=str(STUDENT))
    ap.add_argument("--pandas", action="store_true")
    ap.add_argument("--excel", action="store_true")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--workdir", default=None, help="Excel に渡すコピーと一時ファイルの置き場所")
    a = ap.parse_args()
    if a.all:
        a.pandas = a.excel = True

    check_answers(a.teacher, a.student)
    check_mission_words(a.teacher, a.student)
    check_region_names(a.teacher, "教員版")
    check_region_names(a.student, "学習者版")
    check_pii(a.teacher, "教員版")
    check_pii(a.student, "学習者版")
    check_structure(a.teacher, a.student)
    ex = None
    if a.pandas or a.excel:
        ex = Expect()
    if a.pandas:
        check_pandas(ex)
    if a.excel:
        work = a.workdir or tempfile.mkdtemp(prefix="course_xlsx_check_")
        for label, path, student in (("教員版", a.teacher, False), ("学習者版", a.student, True)):
            check_excel(ex, path, label, os.path.join(work, "t" if not student else "s"), student)

    ng = [r for r in RESULTS if not r[0]]
    print(f"\n合計 {len(RESULTS)} 項目：合格 {len(RESULTS) - len(ng)}、不合格 {len(ng)}")
    for _, n, d in ng:
        print("  不合格：", n, d)
    return 1 if ng else 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
