# -*- coding: utf-8 -*-
"""ガイド型（層2a）の Excel ブックが読む「前処理済みデータ」を作る。

Ver.1.9：ガイド型を表計算ベースに切り替える。表計算では作れない列
（海／陸の区分、永久影までの距離、地質年代の結合）を、ここで先に計算して
`course/data/` に CSV として書き出す。作り方＝根拠は本スクリプトと docs に残す
（＝ブラックボックスにしない）。

    python course/build_course_data.py

出力（すべて UTF-8 BOM つき。Excel でダブルクリックしても化けない）:
  course/data/temp_grid.csv           3度グリッドの温度（24時間カーブ＋平均・較差）
  course/data/craters_labeled.csv     クレーター＋海/陸の区分
  course/data/crater_ages_labeled.csv 年代つきクレーター＋海/陸の区分
  course/data/polar_south_sites.csv   南極の日照・永久影率・永久影までの距離・傾斜
  course/data/polar_north_sites.csv   北極（参考）
  course/data/geology_grid.csv        USGS 統合地質図（3度グリッド。海陸・相対年代の答え合わせ用）
  course/data/env_grid.csv            月全体の環境指標（3度グリッド。日較差・夜の底・太陽高度・地球の仰角）＋地域タイプのラベル
  course/data/landing_sites.csv       実在の着陸地点＋その場所の温度・地質・傾斜
  course/data/reference.csv           海の面積割合など（円近似と USGS の両方）
"""
import pathlib
import sys

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "notebooks"))
import moonkit as mk  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parent / "data"

# 3コマ版（表計算・紙教材）だけで使う地域名の表示名。元の data/candidate_regions.csv（ノートブック版など
# 広く使われる）は変えない。ここに「元の名前 → 3コマ版の表示名」を1つの対応表として置く。
# 括弧書きがミッション（電波天文）と地域の対応をそのまま書いていたため、地理的な説明に替える。
# env_grid.csv の region 列・candidate_regions.csv の name 列・表計算の C5 のプルダウン・紙教材の地域カードは
# すべてこの表示名に統一する。
DISPLAY_NAMES = {
    "裏側・赤道（電波天文の候補域）": "裏側・赤道（月の裏側の赤道帯）",
}


def display_names(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """列 col の地域名を3コマ版の表示名に替える（対応表にない名前はそのまま）"""
    df = df.copy()
    df[col] = df[col].replace(DISPLAY_NAMES)
    return df
OUT.mkdir(exist_ok=True)
LT = mk.LT_COLS


def _write(df: pd.DataFrame, name: str) -> None:
    path = OUT / name
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"  {name}: {len(df):,} 行 / {len(df.columns)} 列")


def temp_grid(step_deg: float = 3.0) -> None:
    """Diviner の全球 0.5度グリッドを step_deg 度に粗くする。
    各点に 1日の平均温度 t_mean_K と較差 t_swing_K（最大-最小）を付ける。"""
    d = mk.load("温度")
    # 3度グリッドの代表点だけ残す（緯度・経度が step の倍数 + 1.5 に近いもの）
    lat_keep = np.isclose((d["lat"] - 0.25) % step_deg, 0.0, atol=0.26)
    lon_keep = np.isclose((d["lon"] - 0.25) % step_deg, 0.0, atol=0.26)
    g = d[lat_keep & lon_keep].copy()
    g = mk.daily_swing(g)  # t_mean_K / t_swing_K / t_std_K を追加
    cols = ["lat", "lon", *LT, "t_mean_K", "t_swing_K"]
    g = g[cols].sort_values(["lat", "lon"])
    val_cols = [c for c in cols if c not in ("lat", "lon")]
    g[val_cols] = g[val_cols].round(1)   # 温度だけ丸める。lat/lon は元のまま（.25 / .75）
    _write(g, "temp_grid.csv")


def craters_labeled() -> None:
    c = mk.load("クレーター")[["lat", "lon", "diam_km"]].copy()
    c = mk.near_maria(c, scale=0.8)          # '区分' = 海 / 陸
    c["diam_km"] = c["diam_km"].round(2)
    _write(c.sort_values("diam_km", ascending=False), "craters_labeled.csv")


def crater_ages_labeled() -> None:
    a = mk.load("クレーター年代")
    lat_c = next(x for x in a.columns if x.lower() == "lat")
    lon_c = next(x for x in a.columns if x.lower() == "lon")
    diam_c = next((x for x in a.columns if x.lower().startswith("diam")), None)
    a = a.rename(columns={lat_c: "lat", lon_c: "lon", diam_c: "diam_km"})
    a = a[["lat", "lon", "diam_km", "Age", "Age_name"]].copy()
    a = mk.near_maria(a, scale=0.8)
    a["diam_km"] = a["diam_km"].round(2)
    _write(a.sort_values("Age"), "crater_ages_labeled.csv")


def polar_sites(sign: int, name: str) -> None:
    """南（sign=-1）／北（sign=+1）の極。永久影までの距離・傾斜を付けて、
    授業で扱える大きさ（3〜4千行）に空間ビンで粗くする。"""
    d = mk.load("極域日照")
    p = d[(d["lat"] * sign) > 0].copy()
    p = mk.dist_to_permanent_shadow(p, threshold=0.9)  # km_to_shadow（全点で計算）

    # 極からの角距離 0.2度 × 経度 3度 でビンにして平均（点群を均す）
    p["lat_bin"] = (p["lat"] * 5).round() / 5
    p["lon_bin"] = (p["lon"] / 3).round() * 3
    g = (p.groupby(["lat_bin", "lon_bin"], as_index=False)
           .agg(average_illumination_percent=("average_illumination_percent", "mean"),
                permanent_shadow_fraction=("permanent_shadow_fraction", "mean"),
                km_to_shadow=("km_to_shadow", "mean"),
                slope_deg=("slope_deg", "mean"),
                n=("lat", "size")))
    g = g.rename(columns={"lat_bin": "lat", "lon_bin": "lon"})
    g["average_illumination_percent"] = g["average_illumination_percent"].round(2)
    g["permanent_shadow_fraction"] = g["permanent_shadow_fraction"].round(3)
    g["km_to_shadow"] = g["km_to_shadow"].round(2)
    g["slope_deg"] = g["slope_deg"].round(2)
    g = g.dropna(subset=["slope_deg"])   # 極点付近の傾斜欠測ビンは落とす

    # 0〜1 に正規化した列（min-max）。Excel のステップ4 はこの列に重みを掛けるだけ。
    # 「よい向き」に合わせて、近い方がよい km_to_shadow・低い方がよい psf・低い方がよい slope は 1 から引く。
    def norm(s, invert=False):
        lo, hi = s.min(), s.max()
        z = (s - lo) / (hi - lo) if hi > lo else pd.Series(0.5, index=s.index)
        return (1 - z if invert else z).round(4)

    g["norm_illum"] = norm(g["average_illumination_percent"])
    g["norm_near_shadow"] = norm(g["km_to_shadow"], invert=True)
    g["norm_low_psf"] = norm(g["permanent_shadow_fraction"], invert=True)
    g["norm_low_slope"] = norm(g["slope_deg"], invert=True)
    g = g[["lat", "lon", "average_illumination_percent",
           "permanent_shadow_fraction", "km_to_shadow", "slope_deg",
           "norm_illum", "norm_near_shadow", "norm_low_psf", "norm_low_slope", "n"]]
    _write(g.sort_values("average_illumination_percent", ascending=False), name)


def geology_grid(step_deg: float = 3.0) -> None:
    """USGS 統合地質図（1度グリッド）を step_deg 度に粗くして、ステップ2の
    『答え合わせ』に使う。海／陸と相対年代（age_index 1〜5）。"""
    g = mk.load("地質")
    keep = (np.isclose(g["lat"] % step_deg, step_deg / 2, atol=0.6) &
            np.isclose((g["lon"] + 0.5) % step_deg, step_deg / 2, atol=0.6))
    out = g.loc[keep, ["lat", "lon", "relative_age", "age_index", "区分"]].copy()
    _write(out.sort_values(["lat", "lon"]), "geology_grid.csv")


def env_grid(step_deg: float = 3.0) -> None:
    """site_environment.csv（全球1度）を step_deg 度に粗くして、ステップ3〜4 で使う。
    各行に、それが入る候補地域タイプの名前（region）と、0〜1 正規化した列を付ける。
    """
    e = mk.load("環境")
    keep = (np.isclose(e["lat"] % step_deg, step_deg / 2, atol=0.6) &
            np.isclose((e["lon"] + 0.5) % step_deg, step_deg / 2, atol=0.6))
    # 表計算版のスコアは 5 指標に絞る（操作を簡単に保つ）。全球傾斜 slope_deg は
    # Python 版・入口アプリ・data/site_environment.csv で見られるが、Excel の env_grid には入れない
    # （build_course_xlsx.py が列位置でスコア列を差し込むため、列を増やすと壊れる）。
    g = e.loc[keep, ["lat", "lon", "区分", "age_index", "temp_amp_K", "night_min_K",
                     "noon_sun_elev_deg", "earth_elev_deg"]].copy()

    # 候補地域タイプのラベルを付ける（region_type と同じ箱）
    regions = display_names(pd.read_csv(mk._find("data", "candidate_regions.csv")), "name")
    g["region"] = ""
    for _, r in regions.iterrows():
        m = g["lat"].between(r["lat_min"], r["lat_max"])
        lo_min, lo_max = float(r["lon_min"]), float(r["lon_max"])
        if not (lo_min <= -179.9 and lo_max >= 179.9):
            if lo_max > 180:
                m &= (g["lon"] >= lo_min) | (g["lon"] <= lo_max - 360)
            else:
                m &= g["lon"].between(lo_min, lo_max)
        g.loc[m & (g["region"] == ""), "region"] = r["name"]

    def norm(s, invert=False):
        lo, hi = s.min(), s.max()
        z = (s - lo) / (hi - lo) if hi > lo else pd.Series(0.5, index=s.index)
        return (1 - z if invert else z).round(4)

    g["norm_sun_high"] = norm(g["noon_sun_elev_deg"])          # 太陽が高いほどよい
    g["norm_amp_low"] = norm(g["temp_amp_K"], invert=True)     # 日較差は小さいほどよい
    g["norm_earth_high"] = norm(g["earth_elev_deg"])           # 地球が見えるほどよい（通信）
    g["norm_earth_low"] = norm(g["earth_elev_deg"], invert=True)  # 地球が見えないほどよい（電波天文）
    g["norm_night_warm"] = norm(g["night_min_K"])              # 夜が冷えないほどよい

    for c in ("temp_amp_K", "night_min_K", "noon_sun_elev_deg", "earth_elev_deg"):
        g[c] = g[c].round(1)
    _write(g.sort_values(["lat", "lon"]), "env_grid.csv")

    # 候補地域の一覧もそのまま Excel に載せる
    _write(regions, "candidate_regions.csv")


def landing_sites_course() -> None:
    """実在の着陸地点に、その場所の温度・地質・（南極なら）傾斜を付ける。
    抽象的なグリッドの分析を『実際の場所』につなげる。"""
    ls = mk.load("着陸地点")
    temp, geo = mk.load("温度"), mk.load("地質")
    lola = mk.load("極域日照")
    rows = []
    for _, s in ls.iterrows():
        t = mk.nearest(temp, s["lat"], s["lon"])
        gg = mk.nearest(geo, s["lat"], s["lon"])
        rec = dict(name=s["name"], kind=s["kind"], lat=s["lat"], lon=s["lon"],
                   terrain=s["terrain"], year=s["year"],
                   temp_noon_K=round(float(t["temp_noon_K"])),
                   temp_midnight_K=round(float(t["temp_midnight_K"])),
                   t_swing_K=round(float(t["temp_noon_K"] - t["temp_midnight_K"])),
                   usgs_区分=gg["区分"], usgs_relative_age=gg["relative_age"])
        if abs(s["lat"]) >= 82.9:
            p = mk.nearest(lola, s["lat"], s["lon"])
            rec["slope_deg"] = round(float(p["slope_deg"]), 1) if pd.notna(p["slope_deg"]) else None
            rec["illum_percent"] = round(float(p["average_illumination_percent"]), 1)
        else:
            rec["slope_deg"] = None
            rec["illum_percent"] = None
        rows.append(rec)
    _write(pd.DataFrame(rows), "landing_sites.csv")


def reference() -> None:
    """ステップ2で「密度（面積あたりの数）」を出すのに要る、月全体に対する
    『海』の面積割合を2通りで見積もる：円近似（教材の near_maria）と USGS 地質図。"""
    lat = np.arange(-89.5, 90, 1.0)
    lon = np.arange(-179.5, 180, 1.0)
    la, lo = np.meshgrid(lat, lon, indexing="ij")
    grid = pd.DataFrame({"lat": la.ravel(), "lon": lo.ravel()})
    grid = mk.near_maria(grid, scale=0.8)
    w = np.cos(np.radians(grid["lat"]))          # 高緯度のセルは面積が小さい
    sea_circle = float((w[grid["区分"] == "海"].sum()) / w.sum())

    geo = mk.load("地質")
    wg = np.cos(np.radians(geo["lat"]))
    sea_usgs = float(wg[geo["区分"] == "海"].sum() / wg.sum())
    age_sea = float(np.average(geo.loc[geo["区分"] == "海", "age_index"],
                               weights=wg[geo["区分"] == "海"]))
    age_land = float(np.average(geo.loc[geo["区分"] == "陸", "age_index"],
                                weights=wg[geo["区分"] == "陸"]))

    ref = pd.DataFrame({
        "項目": ["海の面積割合(円近似)", "陸の面積割合(円近似)",
                 "海の面積割合(USGS地質図)", "陸の面積割合(USGS地質図)",
                 "USGS 海の相対年代の平均(age_index)", "USGS 陸の相対年代の平均(age_index)",
                 "永久影の判定しきい値", "永久影の点の数(南極)"],
        "値": [round(sea_circle, 4), round(1 - sea_circle, 4),
               round(sea_usgs, 4), round(1 - sea_usgs, 4),
               round(age_sea, 3), round(age_land, 3), 0.9,
               int((mk.load("極域日照").query("lat < 0")["permanent_shadow_fraction"] >= 0.9).sum())],
        "説明": [
            "near_maria(scale=0.8) で『海』になるセルの、緯度で重みづけした割合。円で囲むので大きめに出る",
            "1 - 海の面積割合(円近似)",
            "USGS 統合地質図のポリゴンを緯度で重みづけした割合。文献値 約16% と整合",
            "1 - 海の面積割合(USGS地質図)",
            "地質図の海セルの age_index(1古〜5新) の面積重み平均。約3.16",
            "地質図の陸セルの age_index の面積重み平均。約2.38（＝陸のほうが古い）",
            "permanent_shadow_fraction がこの値以上を『永久影』とした",
            "km_to_shadow の最近傍探索に使った永久影の点の数",
        ],
    })
    _write(ref, "reference.csv")


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "regions":
        # 地域名の表示名だけ作り直す（env_grid.csv と candidate_regions.csv のみ）
        print("地域名の表示名を反映:")
        env_grid()
        return
    print("前処理データを作成:")
    temp_grid()
    craters_labeled()
    crater_ages_labeled()
    polar_sites(-1, "polar_south_sites.csv")
    polar_sites(+1, "polar_north_sites.csv")
    geology_grid()
    env_grid()
    landing_sites_course()
    reference()
    print(f"-> {OUT}")


if __name__ == "__main__":
    main()
