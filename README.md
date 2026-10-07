# 月データ探索教材

高校生が、月の公開データ（クレーター・温度・極域日照）を使って
**情報Ⅰの範囲（散布図・基本統計量・条件分岐）でデータ分析を体験する**探究学習教材。

## ▶ ブラウザで開く（インストール不要）

**https://higonouneme.github.io/moon-data-lesson/**

| ページ | 内容 |
|---|---|
| `/` | 3Dの月球儀で有名な地点を見る入口ページ（Three.js） |
| `/app/` | 分析ノートブック本体（JupyterLite / Pyodide。Python もデータもブラウザ内で動く。初回だけ 1〜2 分） |
| `/data.html` | 教材で使う月データ（CSV）のダウンロード一覧 |

重い計算をしたいときは Colab / ローカルでも同じノートブックが動く（下記）。

- GDAL 等の専門環境は不要。
- 「総合的な探究の時間」や「情報Ⅰ」で、**データ分析を通した探究**として活用できる。
  実授業での試行と改善は今後の課題。
- 開発の経緯と設計判断は `docs/requirements_v1.5.md`（3層構成）／`v1.6.md`（弱点対処）／
  `v1.7.md`（Web アプリ化）／`v1.8.md`（3D入口ページ）／`detailed_design_v3.md`（傾斜・地質図・
  着陸地点の追加）／`requirements_v3.1.md`（次の開発課題）にまとめている
  （v1.0〜v1.4 の履歴も `docs/` に残す）。

---

## 3層構成

教材は共有のデータ・ヘルパーの上に、足場かけの強さが異なる2つの構成を載せた3層からなる。

```
┌───────────────────────────┬───────────────────────────┐
│ 層2a  ガイド型（探究講座）    │ 層2b  オープン型（プチ探究） │
│  1コマ完結・クラス全体      │  数週間・選択した班         │
│  強い足場                  │  弱い足場                  │
│  手順固定のノートブック＋    │  課題ブリーフ＋チートシート  │
│  ワークシート＋進行表       │  ＋指導者メモ＋出発点        │
│  ゴール：ムーンベース最適地  │  問い・提案は生徒が決める    │
│  （任意）ML比較・クラスタリング│  自由探索ツール explore     │
├───────────────────────────┴───────────────────────────┤
│ 層1  解析ヘルパー  moonkit.py（22関数・各数行・機械学習なし）│
├───────────────────────────────────────────────────────┤
│ 層0  共有データ基盤（CSV 13種・座標統一・月面画像・出典と再現手順）│
└───────────────────────────────────────────────────────┘
```

| | 層2a ガイド型 | 層2b オープン型 |
|---|---|---|
| 足場 | 強い（`# ★ここを変える` の数値だけ書き換え） | 弱い（データと道具のみ） |
| 時間 | 1コマ（150〜180分）で完走 | 数回に分けて（導入 → 追究 → 発信） |
| 対象 | クラス全体 | 選択した少人数のグループ |
| 設計根拠 | 「解析体験」型チュートリアル | 石田(2022) の主体性 |

---

## リポジトリ構成

```
repo/
├── data/                              層0：共有データ基盤（下表の13種＋保持のみ1種）
├── tools/                             data/*.csv を作る前処理スクリプト（README・SOURCES.md 参照）
├── course/                            層2a の Excel ブックとその前処理
│   ├── build_course_data.py           data/ の CSV → 授業サイズの前処理 CSV
│   ├── build_course_xlsx.py           course_moonbase.xlsx（18シート）を組み立てる
│   └── course_moonbase.xlsx           層2a：表計算版（生徒はコードを書かない）
├── notebooks/
│   ├── moonkit.py                     層1：解析ヘルパー（機械学習なし。region_type/earth_elevation 等）
│   ├── moonkit_ml.py                  任意ステップ6：モデル比較（scikit-learn）
│   ├── course_moonbase.ipynb          層2a：ガイド型（ステップ1〜5）
│   ├── course_moonbase_ml.ipynb       任意ステップ6のノートブック
│   ├── petit_inquiry_start.ipynb      層2b：オープン型の最小の出発点
│   ├── explore.ipynb                  自由探索ツール（変数選択式の散布図）
│   ├── explore_advanced.ipynb         発展編（numpy.polyfit のべき乗則フィット）
│   ├── explore_clustering.ipynb       発展編（クラスタリング。クレーターの形／夜の熱のふるまい）
│   └── assets/                        日本語フォント・月面背景画像
├── docs/
│   ├── requirements_v1.5.md           要件定義（3層構成・機械学習の限定解禁）
│   ├── requirements_v1.6.md           要件定義（模擬授業前の弱点対処）
│   ├── requirements_v1.7.md           要件定義（Web アプリ化）
│   ├── requirements_v1.8.md           要件定義（3D入口ページ）
│   ├── explorer_data_verification.md  3D入口ページ38地点データの検証記録
│   ├── requirements_v1.2〜1.4.md       履歴
│   ├── requirements_v2.1.md           データ検証の技術記録
│   ├── requirements_v3.1.md           次の開発課題（プチ探究の完成・クラスタリング 等）
│   ├── detailed_design_v3.md          Ver.3 のデータ追加（傾斜・地質図・着陸地点）と組み込み設計
│   ├── data_redistribution_review.md  二次配布可否レビュー（public 化して問題なし）
│   ├── prior_work_survey.md           先行実践・類似DB調査
│   ├── improvement_plan_v1.md         全体俯瞰・改善5点
│   ├── report_v4.html / .pdf          成果報告書（本文2ページ＋データ一覧表。v3.2 の是正を反映）
│   ├── worksheet_course.html / .pdf   層2a：生徒用ワークシート
│   ├── teacher_guide_course.md        層2a：指導者用の進行表
│   ├── petit_inquiry_brief.md         層2b：課題ブリーフ
│   ├── petit_inquiry_helpersheet.html / .pdf   層2b：ヘルパー チートシート
│   ├── petit_inquiry_mentor_notes.md  層2b：指導者用メモ
│   ├── petit_inquiry_rubric.md        層2b：評価の観点表
│   ├── petit_inquiry_exemplars.md     層2b：到達水準の例
│   ├── customize_guide.md             自校データへの差し替え手引き
│   └── requirements_v3.2.md ほか       要件定義（v1.2〜v3.2）・詳細設計
├── webapp/                            ブラウザ版の組み立て
│   ├── assemble.py                    notebooks+data+assets → JupyterLite 用 content/ ＋ DL 用 _dl/
│   ├── bootstrap_template.py          各ノートブックに足す「準備」セル
│   ├── data.template.html             データ一覧ページの枠
│   └── react/                         3D入口ページ（Vite / React / Three.js）
├── .github/workflows/deploy-pages.yml push のたびに GitHub Pages に2本立てでビルド・デプロイ
├── LICENSE                            Apache License 2.0（リポジトリ全体）
├── NOTICE                             第三者コンポーネントとデータの帰属
├── data/LICENSES.md                   各データセットのライセンスと出典
├── data/SOURCES.md                    各データの取得元URLと前処理手順
├── requirements.txt                   ローカル実行用
└── run_notebook.bat                   ローカルで Jupyter を起動（Windows）
```

---

## 使い方

- **主：ブラウザ**（上記の GitHub Pages URL）。インストール不要。`webapp/README.md` 参照。
- **従1：Google Colaboratory**。ノートブックの「Open in Colab」から起動し、冒頭セルで
  リポジトリを `git clone` する（`data/` と `notebooks/` が一緒に来る）。追加インストール不要。
- **従2：ローカル**（Jupyter / VSCode）。`pip install -r requirements.txt`。

どのノートブックでも、最初に `from moonkit import *` を実行してから使う。
ヘルパーの一覧は [`docs/petit_inquiry_helpersheet.pdf`](docs/petit_inquiry_helpersheet.pdf) を参照。

---

## データセット（層0・13種）

| `load()` キー | 内容 | 出典 |
|---|---|---|
| `'クレーター'` | 緯度経度・直径・形（離心率・扁平率）。36,377個（直径8km以上） | Robbins Lunar Crater Database, USGS [3] |
| `'クレーター深さ'` | 直径・深さ・深さ÷直径比。24,982個（直径10km以上）※深さは劣化後の“今の深さ” | Wang & Wu 2021 [4] |
| `'クレーター年代'` | 推定地質年代（1〜5）。18,996個 | DeepCraters, figshare [5] |
| `'温度'` | 地点ごとの現地時間0〜23時の温度カーブ。259,200地点（全球0.5度） | LRO Diviner, UCLA [6] |
| `'極域日照'` | 平均日照率・永久影率・**傾斜 slope_deg**。157,922地点（南北緯82.96〜90度。**極域専用**） | LOLA（日照 [7]／傾斜 GDR） |
| `'地質'` | 相対地質年代・海陸区分（age_index 1〜5、区分 海/陸）。1度グリッド64,800点 | USGS 統合地質図 [9] |
| `'着陸地点'` | 実在の着陸地点・Artemis III 候補地・参照地形。29件 | Wagner ほか 2017 [10] ほか |
| `'夜の温度'` | 夜の最低温度と、その緯度平均からのずれ（異常＝岩の多さの代理）。259,200地点 | LRO Diviner, UCLA [6] |
| `'アイソクロン'` | クレーター密度→絶対年代の参照表（Neukum PF＋編年関数）。発展編で使う | Neukum ほか 2001 [11] |
| `'環境'` | 月全体1度グリッドの環境指標（海陸・1日の温度差・夜の底・太陽高度・**地球の仰角**・**全球傾斜 slope_deg**）。64,800点。**南極以外も同じ土俵で基地候補として評価できる** | 地質図＋Diviner の結合＋幾何計算（`tools/build_site_environment.py`）＋ LOLA LDEM_16 の傾斜（`tools/build_global_slope.py`） |
| `'地域'` | 候補地域タイプ8件（赤道の海・裏側・南極・溶岩チューブ…）と緯度経度の箱。`region_type(df, 名前)` で切り出す | USGS/IAU 座標＋各ミッション文献 |
| `'縦孔'` | 溶岩チューブの天窓7件（放射線・熱の遮蔽＝長期滞在拠点の候補） | Wagner & Robinson 2014 ほか |
| （`maria_boundaries.csv`） | 23の海・大洋の中心座標と半径。**ステップ2の `near_maria` 円近似にのみ使用**（`load()` の選択肢には出さない） | USGS 地名辞典 |

取得元 URL・前処理・文献照合は [`data/SOURCES.md`](data/SOURCES.md)・[`docs/detailed_design_v3.md`](docs/detailed_design_v3.md)。
再現スクリプトは [`tools/`](tools/)（`build_slope.py`・`build_diviner_curve.py` は committed CSV を再生成できることを確認済み）。

**データの注意（教材に明記）**
- 温度カーブが信頼できるのは概ね **緯度 ±70度より低い**範囲（極付近は「昼夜」が成立しない）。
  極域の基地判断は温度ではなく LOLA の日照・傾斜データで行う。
- 日照率・傾斜の絶対値は他文献と単純比較しない。「明るい／暗い」「平ら／急」の順序のみ信頼する。
  傾斜は極点から 0.7 度より外だけ有効。
- 「海」の面積割合は `near_maria` の円近似（約19%）と USGS 地質図（約16%＝文献値）で違う。
  教材ではこの差自体を「近似の限界」の題材にする。
- `moon_ephemeris.csv` は `data/` に残すが教材からは参照しない（保持のみ）。
  `moon_earth_correlation.csv`（月齢×地震）は科学的妥当性への疑義により
  `docs/requirements_v1.4.md` §9 で除外し、2026-09-04 に `data/` からも削除した。

---

## 学習指導要領との対応

- 情報Ⅰ：(4) データの収集・整理・分析／(3) モデル化とシミュレーション
- 地学基礎：(2) 変動する地球　宇宙、太陽系と地球の誕生
- 任意ステップ6（機械学習の比較）は情報Ⅱの先取り・発展に位置づける。

## 明示的にスコープ外とするもの

- 生の衛星画像（GeoTIFF 等）の直接処理、GIS 操作（前処理済みグリッド CSV までは許容）
- 外部 API へのライブ接続、バックエンドサーバー（Web アプリは静的ファイルのみ）
- 機械学習は任意の発展（`moonkit_ml.py`）に限る。教師あり（`train`）＝ステップ6a、
  教師なしクラスタリング（`cluster`）＝ステップ6b／`explore_clustering.ipynb`。層1（`moonkit.py`）は機械学習を含まない。

---

## ライセンス

- **コード・ノートブック・ドキュメント**：Apache License 2.0（`LICENSE`）。
  3D入口ページ（`webapp/react/`）は Google AI Studio 生成の Apache-2.0 テンプレートを
  改変したもので、リポジトリ全体で同じライセンスに揃えている（`NOTICE` 参照）。
- **データセット**：コードとは別に、それぞれ元データのライセンスが適用される。
  `craters_3d.csv`（Wang ほか 2021）と `deepcraters.csv`（Yang ほか 2020）は **CC BY 4.0**
  なので、使う際は著者・出典・ライセンス（https://creativecommons.org/licenses/by/4.0/ ）と
  「教材用に前処理・抽出した」旨を明記すること。ほかは NASA / USGS / JPL 由来で実質
  パブリックドメイン（学術慣行として出典明記を推奨）。取得元と前処理は `data/LICENSES.md`
  ／`data/SOURCES.md` を参照。
- 3D入口ページが表示する約38の「代表地点」は観測データとは別の紹介用の概略データ
  （検証記録：`docs/explorer_data_verification.md`）。分析には `/data.html` の CSV を使う。

---

## 出典

1. 文部科学省, 高等学校学習指導要領（平成30年告示）解説 情報編 (2018)。
2. 石田光宏, 高等学校「課題探究型授業」における天文分野の調査結果, 天文教育 34(2) (2022)。
3. Robbins, S. J., *A New Global Database of Lunar Impact Craters >1–2 km*, JGR Planets, 123 (2018)。NASA PDS Annex / USGS Astrogeology。
4. Wang, Y., Wu, B., Xue, H., Li, X. & Ma, J., *An improved global catalog of lunar impact craters (≥1 km) with 3D morphometric information and updates on global crater analysis*, JGR Planets, 126, e2020JE006728 (2021)。データ: Zenodo 10.5281/zenodo.4983248 (CC BY 4.0)。
5. Yang, C., Zhao, H., Bruzzone, L. ほか, *Lunar impact crater identification and age estimation with Chang'E data by deep and transfer learning*, Nature Communications, 11, 6358 (2020)。データ: figshare 10.6084/m9.figshare.12768539 (CC BY 4.0)。
6. Williams, J.-P. ほか, *The global surface temperatures of the Moon as measured by the Diviner Lunar Radiometer Experiment*, Icarus, 283, 300–325 (2017)。データ配布：UCLA Diviner チーム（瞬間温度マップ24枚を現地時間に位相合わせして利用）。
7. Mazarico, E. ほか, *Illumination conditions of the lunar polar regions using LOLA topography*, Icarus, 211 (2011)。データ配布：LRO LOLA Team (NASA GSFC)。傾斜は LOLA GDR bidirectional slope (LDSM, 240m 基線, PDS)。
8. NASA Scientific Visualization Studio, *CGI Moon Kit*（散布図の背景画像・海の輪郭、パブリックドメイン）。
9. Fortezzo, C. M., Skinner, J. A., Hunter, M. A., *Unified Geologic Map of the Moon*, USGS Scientific Investigations Map 3316 (2020)。
10. Wagner, R. V. ほか, *Coordinates of anthropogenic features on the Moon*, Icarus, 283 (2017)／NASA Artemis III candidate regions (2024)。
11. Neukum, G., Ivanov, B. A., Hartmann, W. K., *Cratering records in the inner solar system in relation to the lunar reference system*, Space Science Reviews, 96 (2001)。生産関数・編年関数の式のみ使用（`tools/build_isochron.py`）。
