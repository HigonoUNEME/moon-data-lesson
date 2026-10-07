# 表計算 ⇔ Python 対応表

`requirements_v3.3.md` I3。本教材は「同じ CSV を、表計算（Excel／Google スプレッドシート）でも
Python（moonkit）でも同じ手順で扱える」ことを売りにしている。その対応関係を1枚にまとめる。

- 表計算版の正本：`course/course_moonbase.xlsx`（18シート）
- Python 版の正本：`notebooks/course_moonbase.ipynb` ＋ `notebooks/moonkit.py`
- どちらも読むデータは `data/*.csv`（座標系・列名は共通）

---

## 1. 基本操作の対応

| やりたいこと | 表計算 | Python（pandas / moonkit） |
|---|---|---|
| データを開く | シート「データ_〇〇」を見る | `df = load('温度')` など（`DATASETS` のキー） |
| 行数を数える | `=COUNTA(A:A)-1` | `len(df)` |
| ある条件の行だけ数える | `=COUNTIFS(範囲, ">="&下限, 範囲, "<="&上限)` | `df[df['lat'].between(下限, 上限)].shape[0]` |
| 条件つき平均 | `=AVERAGEIFS(値の範囲, 条件範囲, 条件)` | `df.loc[条件, '値の列'].mean()` |
| グループごとの平均 | `=AVERAGEIF(区分の範囲, "海", 値の範囲)` を区分ごとに | `df.groupby('区分')['値の列'].mean()` |
| グループごとの件数・平均をまとめて | 区分ごとに `COUNTIF` / `AVERAGEIF` を並べる | `summary_by(df, group='区分', value='diam_km')` |
| 緯度・経度でしぼる | `COUNTIFS` / オートフィルタで緯度列を範囲指定 | `region(df, lat=(-10, 10), lon=(...))` |
| 上位 N 件を出す | `=LARGE(範囲, 1)`, `=INDEX(名前列, MATCH(LARGE(範囲,k), 範囲, 0))` | `df.nlargest(N, 'スコア')` |
| ヒストグラム | 「挿入」→ヒストグラム、または `FREQUENCY` | `hist(df, '列名', vline=しきい値)` |
| 散布図（地図） | 「挿入」→散布図（X=経度, Y=緯度） | `scatter(df, 'lon', 'lat', color='列名')` |
| ある地点にいちばん近い行 | 補助列に距離 `=SQRT((緯度-la)^2+(経度-lo)^2)` → `MIN` | `nearest(df, la, lo)` |

---

## 2. スコア式（適地評価）の対応 ― ここが教材の核

「複数の指標を 0〜1 に正規化して、重みをつけて合計する」。表計算では**正規化列を前処理で入れてある**
（`norm_amp_low` など、`データ_環境` の J〜N 列 / `データ_南極` の右側）。Python では `site_score` が内部で正規化する。

| 段階 | 表計算 | Python |
|---|---|---|
| 正規化（min–max） | 前処理済み列を使う。式は `=(x - MIN(範囲)) / (MAX(範囲) - MIN(範囲))`。「低いほど良い」指標は `1 - それ` | `site_score` が内部で実施（`want` の `'低い'` 指定で反転） |
| 重みづけ合計 | `データ_環境!O列 = norm_sun*$C$8 + norm_amp*$C$9 + …`（重み `$C$8:$C$12` は「ステップ4_地域を選ぶ」の黄色いセル） | `want = {'noon_sun_elev_deg': ('高い', 2), 'temp_amp_K': ('低い', 2), ...}` |
| 地域でしぼる | `データ_環境!I列`（region ラベル）＝選んだ地域名 のときだけ O 列を計算：`=IF($I2=ステップ4_地域を選ぶ!$C$5, 合計, "")` | `region_type(load('環境'), my_region)` で切り出してから `site_score` |
| 上位を出す | `=LARGE(データ_環境!$O$2:$O$14401, k)` と `MATCH` で緯度経度を引く | `site_score(region, want, top=10)` |

**同じことをしている**：`ステップ4_地域を選ぶ` シートの C8:C12 の重み ＝ Python の `want` 辞書の重み。
表計算では「地域名を C5 に入れる」、Python では「`my_region` に入れる」。

---

## 3. ステップごとの対応

| ステップ | 表計算シート | Python セル（`course_moonbase.ipynb`） | 主な関数・式 |
|---|---|---|---|
| 1 温度 | `ステップ1_温度` | `region(load('温度'), lat=my_band)` → `diurnal_curve` → `daily_swing` → `summary` | `AVERAGEIFS` ⇔ `groupby().mean()` |
| 2 海と陸 | `ステップ2_海と陸` | `near_maria(...)`, `summary_by(..., group='区分')`, `load('地質')` で答え合わせ | `COUNTIF`/`AVERAGEIF` ⇔ `groupby` |
| 3 月全体 | `ステップ3_月全体` | `env = load('環境')`, `scatter`, `region_type` の比較ループ | `AVERAGEIFS`（緯度帯）＋ `AVERAGEIF`（地域ラベル） ⇔ `between` / `region_type` |
| 4 地域を選ぶ | `ステップ4_地域を選ぶ` ＋ `データ_環境!O列` | `region_type(load('環境'), my_region)` → `site_score(region, want, top=10)` | `IF`＋重み付き和＋`LARGE`/`MATCH` ⇔ `site_score` |
| 4b 南極（分岐） | `ステップ4b_南極`, `ステップ4b_スコア` ＋ `データ_南極` | 別ノートブック `course_moonbase_polar.ipynb`：`south_pole(load('極域日照'))`, `dist_to_permanent_shadow`, `site_score` | 同上（指標が日照率・傾斜・永久影距離に変わるだけ） |
| 5 まとめ | `ステップ5_まとめ` | `nearest(env, lat, lon)` を実在計画ごとに | `INDEX`/`MATCH` ⇔ `nearest` |
| 6 機械学習（任意） | （表計算版なし） | `course_moonbase_ml.ipynb` | 情報Ⅱ(3)。表計算では扱わない |

---

## 4. 表計算にできないこと（前処理で吸収している）

以下は表計算の関数では作れないので、**前処理スクリプトで列を計算して配布データに入れてある**。
作り方は `data/SOURCES.md` と `tools/` のスクリプト、`course/build_course_data.py` に明記。

| 列 | どのデータ | なぜ表計算で作れないか |
|---|---|---|
| `区分`（海/陸） | クレーター・環境 | 23の海の中心座標との距離判定（円近似）。moonkit の `near_maria` が担う |
| `km_to_shadow` / `permanent_shadow_fraction` | 極域日照 | 最近傍探索（k-d 木）。`dist_to_permanent_shadow` |
| `age_index` / `relative_age` | 地質 | USGS 地質図ポリゴンの空間結合 |
| `earth_elev_deg` | 環境 | 幾何計算（`90 − acos(cos lat · cos lon)`）。`earth_elevation` でも見られる |
| `temp_amp_K` / `night_min_K` | 環境 | Diviner 24枚の位相合わせ＋1°セル集計 |
| `norm_*`（0〜1の点数） | 環境・南極 | min–max 正規化。式は §2 のとおりで、表計算でも書けるが手間なので前処理済み |

---

## 5. 教室での使い分け（teacher_guide 参照）

- **ガイド型（一斉授業）**：情報Ⅰ「データの活用」が表計算を主たる道具と明記しているため、**表計算を主**とする。
  ノートブックは「同じことを Python でやるとこうなる」の対照として、余力のある班・教員のデモで使う。
- **プチ探究（テーマ選択制）**：生徒が自由に探索するので **Python（moonkit）を主**とする。
  表計算しか使えない生徒には `explore` 系の操作を xlsx でなぞれるよう helpersheet に対応表（本書の §1）を載せる。

この対応表は `teacher_guide_course.md` と `worksheet_course.html`、`petit_inquiry_helpersheet.html` から参照する。
