# -*- coding: utf-8 -*-
"""（任意）Excel で再計算して「計算結果を保存した版」を作り、個人名とローカルパスを取り除く。

    python course/recalc_xlsx_with_excel.py 入力.xlsx 出力.xlsx

なぜ必要か：build_course_xlsx.py（openpyxl）が出す xlsx には計算結果が入っていない。
Excel の「保護ビュー」（メール添付・ダウンロード直後）では再計算されないため、値が空欄に見える
（実測：COM の ProtectedViewWindows で開くと D8 が空欄。Excel で保存し直した版は 295 と出る）。
Excel で開いて保存し直すと値が入る。

ただし Excel で保存すると、次の個人情報が xlsx に入る（実測）：
  - docProps/core.xml の lastModifiedBy（Office に登録された利用者名）
  - xl/workbook.xml の absPath（保存したときのフォルダの絶対パス＝ユーザー名を含む）
このスクリプトは、Excel（COM）で「開く→全再計算→別名で保存」したあと、この2つを xlsx の XML から
取り除く（ZIP を書き直す）。元ファイルは上書きしない。Excel が必要（Windows）。
配付前に check_course_xlsx.py --student 出力.xlsx で答え語 0 件を、PII の検査（同ファイル）で
個人名・パスが無いことを確かめる。
"""
import argparse
import os
import pathlib
import re
import shutil
import sys
import tempfile
import zipfile

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import check_course_xlsx as chk  # noqa: E402


def scrub(src, dst):
    """lastModifiedBy・absPath・作成者（creator）を除去する"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                t = data.decode("utf-8")
                t = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "", t, flags=re.S)
                t = re.sub(r"<cp:lastModifiedBy\s*/>", "", t)
                t = re.sub(r"<dc:creator>.*?</dc:creator>", "", t, flags=re.S)
                t = re.sub(r"<dc:creator\s*/>", "", t)
                data = t.encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                t = data.decode("utf-8")
                t = re.sub(r"<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>\s*<x15ac:absPath[^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>",
                           "", t, flags=re.S)
                t = re.sub(r"<x15ac:absPath[^>]*/>", "", t)
                data = t.encode("utf-8")
            zout.writestr(item, data)


def recalc(src, dst):
    """src（openpyxl が出した xlsx）を Excel で開いて全再計算し、個人情報を除いて dst に保存する"""
    src, dst = pathlib.Path(src).resolve(), pathlib.Path(dst).resolve()
    if src == dst:
        raise ValueError("入力と出力が同じです。元ファイルは上書きしません。")
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="recalc_"))
    try:
        saved = tmp / "excel_saved.xlsx"
        out = chk.run_excel([{"op": "calc", "key": "x"}, {"op": "saveas", "path": str(saved)}], str(src), str(tmp / "w"))
        if out.get("errors"):
            raise RuntimeError("Excel の操作に失敗: " + str(out["errors"]))
        dst.parent.mkdir(parents=True, exist_ok=True)
        scrub(saved, dst)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return dst


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    a = ap.parse_args()
    dst = recalc(a.src, a.dst)
    print(f"wrote {dst} ({dst.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
