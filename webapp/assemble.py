# -*- coding: utf-8 -*-
"""webapp/content/ を組み立てる（JupyterLite に渡す教材一式）。

- notebooks/ の .ipynb に、ブラウザ用のブートストラップセルを先頭に追加する
- moonkit.py / moonkit_ml.py と、教材で使う data/ ・ assets/ をコピーする
- docs/ の md・pdf も参照用にコピーする
- Diviner の gz を展開して、素の csv もダウンロードできるようにする
- データカタログ webapp/data.html を、実データから生成する

GitHub Actions（.github/workflows/deploy-pages.yml）から呼ばれる。ローカルでも
`python webapp/assemble.py && jupyter lite build --contents webapp/content` で確認できる。
"""
import argparse
import csv
import gzip
import html
import json
import pathlib
import shutil
import warnings

import nbformat

warnings.filterwarnings("ignore", message="Cell is missing an id field")

ROOT = pathlib.Path(__file__).resolve().parent.parent
WEBAPP = ROOT / "webapp"

# ブラウザ版に含めるノートブック
NOTEBOOKS = [
    "course_moonbase.ipynb",
    "course_moonbase_polar.ipynb",
    "course_moonbase_ml.ipynb",
    "petit_inquiry_start.ipynb",
    "explore.ipynb",
    "explore_advanced.ipynb",
    "explore_clustering.ipynb",
]
# 一緒に置く Python ファイル
PYFILES = ["moonkit.py", "moonkit_ml.py"]
# 教材が読み込むデータ
DATA = [
    "craters_subset.csv", "craters_3d.csv", "deepcraters.csv",
    "diviner_global.csv.gz", "lola_polar_illumination.csv", "maria_boundaries.csv",
    "moon_geology_grid.csv", "landing_sites.csv", "diviner_nighttime.csv.gz",
    "isochron_reference.csv", "site_environment.csv", "candidate_regions.csv",
    "lunar_pits.csv",
]
ASSETS = ["NotoSansJP-Regular.ttf", "lroc_color_2k.jpg"]
DOCS = [
    "petit_inquiry_brief.md", "petit_inquiry_helpersheet.pdf",
    "petit_inquiry_mentor_notes.md", "teacher_guide_course.md",
    "worksheet_course.pdf",
]
# course/ から配布物へコピーするもの（層2a の表計算版・正本）
COURSE_FILES = ["course_moonbase.xlsx"]

# ブートストラップが fetch するファイルの一覧（content/ からの相対パス）
FETCH_LIST = (
    PYFILES
    + [f"data/{d}" for d in DATA]
    + [f"assets/{a}" for a in ASSETS]
)

# データカタログ（data.html）に載せる情報。ファイル名 → (見出し, 説明, ライセンス, 出典)
CATALOG = {
    "craters_subset.csv": (
        "クレーターの緯度経度・直径・形",
        "月クレーターDB（Robbins 2018）から直径8km以上を抽出。離心率・扁平率で「どれくらい丸いか」がわかる。",
        "パブリックドメイン（NASA PDS / USGS）",
        "Robbins, S. J. (2018), JGR Planets 123. NASA PDS Annex / USGS Astrogeology"),
    "craters_3d.csv": (
        "クレーターの直径と深さ",
        "3D形態情報つきカタログ（Wang ほか 2021、査読付き）から直径10km以上。複数のグローバルDEMからガウスフィットで深さを算出。"
        "教材側で直径10km以上に抽出・列名変更・CSV化した（値は未加工）。",
        "CC BY 4.0 (https://creativecommons.org/licenses/by/4.0/)",
        "Wang, Y., Wu, B., Xue, H., Li, X. & Ma, J. (2021), JGR Planets 126, e2020JE006728. データ: Zenodo 10.5281/zenodo.4983248"),
    "deepcraters.csv": (
        "クレーターの推定地質年代",
        "嫦娥1・2号データから機械学習で抽出し、地質年代（1=最古〜5=最新）を推定したカタログ。Aged サブセット（Age列つき）を抽出した。",
        "CC BY 4.0 (https://creativecommons.org/licenses/by/4.0/)",
        "Yang, C., Zhao, H., Bruzzone, L. ほか (2020), Nature Communications 11, 6358. データ: figshare 10.6084/m9.figshare.12768539"),
    "diviner_global.csv.gz": (
        "月面の温度（現地時間0〜23時の1日のカーブ）",
        "LRO Diviner の瞬間温度マップ24枚を現地時間に位相合わせして日変化カーブに再構成したもの。全球0.5度グリッド。"
        "信頼できるのは概ね緯度±70度より低い範囲（極付近は位相整列が退化して振幅が過大に出る）。",
        "パブリックドメイン（NASA）",
        "UCLA Diviner チーム level-4 製品。背景: Williams, J.-P. ほか (2017), Icarus 283, 300–325"),
    "lola_polar_illumination.csv": (
        "月の南極・北極の日照率・永久影率・傾斜",
        "LOLA 標高データから、日照条件（horizon 法, Mazarico 2011）と傾斜（GDR 240m 基線, LDSM）を約1kmブロックで整理。南北緯82.96〜90度。"
        "日照率・傾斜とも絶対値は他文献と単純比較せず、「明るい／暗い」「平ら／急」の順序を使うこと。極点0.7度以内の傾斜は欠測。",
        "パブリックドメイン（NASA）",
        "LRO LOLA Team (NASA GSFC / MIT)。日照: Mazarico ほか (2011), Icarus 211。傾斜: LOLA GDR LDSM (PDS)"),
    "maria_boundaries.csv": (
        "月の海・大洋 23件の中心座標と半径",
        "USGS / IAU 地名辞典から抽出。教材のステップ2で「海」と「陸」を分けるのに使う（円近似）。",
        "パブリックドメイン", "USGS / IAU Gazetteer of Planetary Nomenclature"),
    "moon_geology_grid.csv": (
        "月全体の相対地質年代・海陸の区分（1度グリッド）",
        "USGS の月統合地質図（1:5,000,000）を1度グリッド（64,800点）に空間結合。相対年代（Pre-Nectarian〜Copernican、age_index 1〜5）と海／陸。"
        "クレーターの数から推定した新旧の「答え合わせ」に使う。面積重みつきの海の割合は約16%（文献値と整合）。",
        "パブリックドメイン（USGS）",
        "Fortezzo, C. M. ほか (2020), Unified Geologic Map of the Moon, USGS SIM 3316"),
    "landing_sites.csv": (
        "実在の月着陸地点・Artemis III 候補地・参照地形（29件）",
        "Apollo・Luna・Surveyor・Chang'e・Chandrayaan-3 の着陸座標、Artemis III の候補領域の座標、Shackleton/Tycho/Copernicus など。"
        "グリッドの分析結果を「実際の場所」に結びつけるための手キュレーション表。",
        "パブリックドメイン（NASA / USGS 座標）",
        "Wagner, R. V. ほか (2017), Coordinates of anthropogenic features on the Moon, Icarus 283 ／ NASA Artemis III candidate regions (2024)"),
    "diviner_nighttime.csv.gz": (
        "月面の夜の最低温度と、その緯度平均からのずれ（異常）",
        "Diviner の夜の最低温度（temp_min_K）と、緯度平均を引いた異常（temp_min_anomaly_K）。異常が正＝夜も冷めにくい＝岩が多い（熱物性の代理）。"
        "発展ノートブック（クラスタリング）で「周りと違うふるまいの場所」を探すのに使う。",
        "パブリックドメイン（NASA）",
        "UCLA Diviner チーム level-4 製品。背景: Williams, J.-P. ほか (2017), Icarus 283"),
    "isochron_reference.csv": (
        "クレーター密度 → 絶対年代 の参照表（発展）",
        "各年代（0.5〜4.3 Ga）で、直径8km/20km 以上のクレーターが 100万km² あたり何個あるはずかを、"
        "Neukum の生産関数＋月の編年関数（Neukum et al. 2001）から計算した表。教材側で計算した派生物。",
        "参考文献ベースの計算値（式は Neukum et al. 2001）",
        "Neukum, G., Ivanov, B. A., Hartmann, W. K. (2001), Space Science Reviews 96"),
    "site_environment.csv": (
        "月ぜんたいの環境指標（1度グリッド。海陸・温度・地球の見えかた）",
        "月全体を同じ尺度で比べられるようにした全球グリッド。区分（海／陸）・相対年代・"
        "1日の温度の最大/最小/差・夜の最低温度・正午の太陽高度・地球の仰角（正＝表側、負＝裏側）。"
        "既存の月データ（USGS 地質図・Diviner 温度）の結合と幾何計算だけで作った派生物（新規観測なし）。",
        "パブリックドメイン（USGS / NASA 由来）＋幾何計算",
        "moon_geology_grid.csv（USGS SIM 3316）＋ diviner_global/diviner_nighttime（UCLA Diviner）＋ 幾何計算。tools/build_site_environment.py"),
    "candidate_regions.csv": (
        "8つの地域の緯度経度の範囲",
        "地域名と緯度経度の範囲（最小・最大）の表。site_environment から地域ごとに値を切り出すための手キュレーション表。",
        "パブリックドメイン（USGS 座標）＋各ミッション文献",
        "USGS / IAU Gazetteer ＋ 各ミッション・構想（Artemis III / LCRT / Chang'e 4,6 ほか）"),
    "lunar_pits.csv": (
        "溶岩チューブの天窓（縦孔）7件",
        "溶岩チューブの天井が崩れてできた縦孔。Marius Hills・静かの海・"
        "Mare Ingenii（裏側）など既知の縦孔の座標・開口径・深さ。座標は文献値を手キュレーション（全アトラスの再配布はしない）。",
        "座標は文献値（LROC 由来。実質パブリックドメイン）",
        "Wagner, R. V. & Robinson, M. S. (2014), Icarus 237 ／ Robinson, M. S. ほか (2012), Planet. Space Sci. 69 ／ Lunar Pit Atlas (LROC/ASU)"),
}


def bootstrap_source() -> str:
    tmpl = (WEBAPP / "bootstrap_template.py").read_text(encoding="utf-8")
    files_repr = "[\n        " + ",\n        ".join(json.dumps(f) for f in FETCH_LIST) + ",\n    ]"
    return tmpl.replace("__FILE_LIST__", files_repr)


def _fmt_bytes(n: int) -> str:
    x = float(n)
    for unit in ("B", "KB", "MB"):
        if x < 1024 or unit == "MB":
            return f"{x:.0f} {unit}" if unit == "B" else f"{x:.1f} {unit}"
        x /= 1024
    return f"{x:.1f} MB"


def _preview_rows(path: pathlib.Path, gz: bool, n: int = 6):
    opener = gzip.open if gz else open
    with opener(path, "rt", encoding="utf-8", newline="") as f:
        r = csv.reader(f)
        header = next(r)
        rows = [next(r, None) for _ in range(n)]
    return header, [x for x in rows if x is not None]


def _table_html(header, rows) -> str:
    th = "".join(f"<th>{html.escape(c)}</th>" for c in header)
    trs = ""
    for row in rows:
        tds = "".join(f"<td>{html.escape(c)}</td>" for c in row)
        trs += f"<tr>{tds}</tr>"
    return f'<div class="tablewrap"><table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table></div>'


def write_data_catalog(content: pathlib.Path) -> None:
    """webapp/data.html を実データから生成する。"""
    sections = []
    for fname, (title, desc, lic, src) in CATALOG.items():
        gz = fname.endswith(".gz")
        srcfile = ROOT / "data" / fname
        header, rows = _preview_rows(srcfile, gz)
        size = srcfile.stat().st_size
        n_rows = sum(1 for _ in (gzip.open(srcfile, "rt", encoding="utf-8") if gz
                                 else open(srcfile, "rt", encoding="utf-8"))) - 1

        # data.html はサイト直下、CSV 実体も /data/ に直接置く（JupyterLite に依存しない）
        dl = f'<a class="dl" href="./data/{fname}" download>{fname}（{_fmt_bytes(size)}）</a>'
        if gz:
            plain = fname[:-3]  # .gz を外した名前（assemble が content に展開して置く）
            dl = (f'<a class="dl" href="./data/{plain}" download>{plain}（展開版・約42 MB）</a>'
                  f'<a class="dl alt" href="./data/{fname}" download>{fname}（{_fmt_bytes(size)}・gzip圧縮）</a>')

        sections.append(f"""
  <section>
    <h2>{html.escape(title)}</h2>
    <p class="desc">{html.escape(desc)}</p>
    <p class="meta">{n_rows:,} 行 ／ {len(header)} 列 ／ ライセンス：{html.escape(lic)}<br>
       出典：{html.escape(src)}</p>
    <p class="dls">{dl}</p>
    <details><summary>先頭の数行を見る</summary>
    {_table_html(header, rows)}
    </details>
  </section>""")

    template = (WEBAPP / "data.template.html").read_text(encoding="utf-8")
    (WEBAPP / "data.html").write_text(
        template.replace("<!--SECTIONS-->", "\n".join(sections)), encoding="utf-8")
    print(f"wrote {WEBAPP / 'data.html'}")


def stage_downloads(dl_dir: pathlib.Path) -> None:
    """データ一覧ページ（/data/）から直接落とせる CSV を用意する。
    Diviner は gz と展開版の両方を置く。JupyterLite とは独立。"""
    if dl_dir.exists():
        shutil.rmtree(dl_dir)
    dl_dir.mkdir(parents=True)
    for name in DATA:
        shutil.copy2(ROOT / "data" / name, dl_dir / name)
        if name.endswith(".gz"):
            with gzip.open(ROOT / "data" / name, "rb") as fi, \
                 open(dl_dir / name[:-3], "wb") as fo:
                shutil.copyfileobj(fi, fo)
    print(f"staged downloads -> {dl_dir}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=str(WEBAPP / "content"),
                    help="JupyterLite 用 content の出力先（既定：webapp/content）")
    ap.add_argument("--dl-out", default=str(WEBAPP / "_dl"),
                    help="データ一覧の DL 用フォルダの出力先（既定：webapp/_dl）")
    args = ap.parse_args()
    CONTENT = pathlib.Path(args.out).resolve()

    if CONTENT.exists():
        shutil.rmtree(CONTENT)
    (CONTENT / "data").mkdir(parents=True)
    (CONTENT / "assets").mkdir(parents=True)
    (CONTENT / "docs").mkdir(parents=True)

    boot = bootstrap_source()
    for name in NOTEBOOKS:
        nb = nbformat.read(ROOT / "notebooks" / name, as_version=4)
        cell = nbformat.v4.new_code_cell(boot, metadata={"tags": ["webapp-bootstrap"]})
        nb.cells.insert(0, cell)
        # JupyterLite の Pyodide カーネルは name="python"。合わせておかないと
        # 「No Kernel」になる。
        nb.metadata["kernelspec"] = {"name": "python", "display_name": "Python (Pyodide)"}
        for i, c in enumerate(nb.cells):   # 全セルに id を付ける（元ノートブックには無い）
            c.setdefault("id", f"{name[:-6]}-{i:02d}")
        nbformat.validate(nb)
        nbformat.write(nb, CONTENT / name)

    for name in PYFILES:
        shutil.copy2(ROOT / "notebooks" / name, CONTENT / name)
    for name in DATA:
        # JupyterLite にはノートブックが読む形（Diviner は gz のまま）だけを渡す。
        # 展開版の 42MB を content/ に置くと app/files/ が膨らむので置かない。
        shutil.copy2(ROOT / "data" / name, CONTENT / "data" / name)
    for name in ASSETS:
        shutil.copy2(ROOT / "notebooks" / "assets" / name, CONTENT / "assets" / name)
    for name in DOCS:
        src = ROOT / "docs" / name
        if src.exists():
            shutil.copy2(src, CONTENT / "docs" / name)
    for name in COURSE_FILES:
        src = ROOT / "course" / name
        if src.exists():
            shutil.copy2(src, CONTENT / "docs" / name)

    write_data_catalog(CONTENT)
    stage_downloads(pathlib.Path(args.dl_out).resolve())

    print(f"assembled -> {CONTENT}")
    for p in sorted(CONTENT.rglob("*")):
        if p.is_file():
            print(f"  {p.relative_to(CONTENT)}  ({p.stat().st_size:,} B)")


if __name__ == "__main__":
    main()
