# -*- coding: utf-8 -*-
"""3D入口ページに重ねる「データ層」テクスチャを作る（requirements_v3.3 の続き・地球の風 Phase1/2）。

data/site_environment.csv（全球1度グリッド）から、指定した列を等距円筒図法の
PNG（RGBA）に焼き込む。月本体のテクスチャ（lroc_color_2k.jpg）と同じ座標規約
（x=0 が経度-180°、y=0 が緯度+90°）で作るので、Three.js の SphereGeometry に
そのまま貼れる。欠測（例：slope_deg の極域）はアルファ0（透明）にし、月の実写が透けて見える。

    python webapp/react/scripts/gen_overlay_textures.py

出力：webapp/react/public/textures/overlay_<field>.png
      webapp/react/src/data/overlayLayers.generated.json（凡例用のメタデータ。
      色の凡例バーは実際に使ったカラーマップから CSS グラデーションを生成して埋め込む
      ＝ 見た目と数値が必ず一致する）
"""
import json
import os

import matplotlib
matplotlib.use("Agg")
import numpy as np
import pandas as pd
from scipy.ndimage import zoom

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
DATA = os.path.join(ROOT, "data")
REACT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEX_OUT = os.path.join(REACT, "public", "textures")
JSON_OUT = os.path.join(REACT, "src", "data", "overlayLayers.generated.json")
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_cache")
os.makedirs(TEX_OUT, exist_ok=True)

UPSCALE = 4  # 360x180 の1度グリッド -> 1440x720（滑らかに見せるための補間。データ自体は1度分解能のまま）

# age_index だけ、site_environment.csv（1度グリッド）ではなく専用の0.25度グリッドを使う
# （2026-09-18、「地質年代の解像度を上げたい」への対応）。tools/build_geology_grid.py --step 0.25 で
# 作る中間生成物（64MB超）。data/ には置かず（配布物ではないため）_cache/ に置きコミットしない。
# 無ければこのスクリプトが自動生成を促す（USGSのGISデータ224MBを毎回落とすのは重いので、
# 手動で1回だけ tools/build_geology_grid.py を実行してもらう）。
AGE_FINE_GRID = os.path.join(CACHE, "moon_geology_grid_fine.csv")
AGE_FINE_UPSCALE = 2  # 元データがすでに0.25度=1度グリッドの4倍細かいので、補間の倍率はここでは控えめでよい

# Phase2：site_environment.csv の6指標。極域日照（lola_polar_illumination）は範囲が
# |緯度|>=83°の細い帯だけで全球図には向かないため対象外（極域は shadow_sim.html 側で扱う）。
#
# 配色メモ（2026-09-18、見た目の指摘を受けて変更）：
#  - slope_deg は元 cividis（色覚バリアフリー向けの地味な単色系）だと細かい起伏がザラついて
#    見えるだけで「傾斜が急」という意味が伝わりにくかったので、危険度が直感的にわかる
#    YlOrRd（平ら=薄い黄、急=赤）に変更。
#  - age_index は元 Spectral（虹色の発散配色）だと1〜5という「古い→新しい」の順序が
#    直感的に読めなかった（虹色は「種類の違い」を表すには向くが「連続的な順序」には不向き）ので、
#    単色で明るさが単調に変わる cividis（暗い紫〜明るい黄）に変更。
LAYERS = [
    # 温度系は「熱い=赤、冷たい=青」の直感に合わせて coolwarm（発散配色）にする。
    dict(key="temp_amp_K", label="1日の温度差", unit="K", cmap="coolwarm", vmin=None, vmax=None,
         desc="1日のうち、いちばん暑い時刻といちばん寒い時刻の温度の差（最高−最低）。"),
    dict(key="night_min_K", label="夜の最低温度", unit="K", cmap="coolwarm", vmin=None, vmax=None,
         desc="夜のあいだにいちばん下がったときの温度。"),
    dict(key="noon_sun_elev_deg", label="正午の太陽高度", unit="°", cmap="viridis", vmin=0, vmax=90,
         desc="正午に太陽が地平線から何度の高さにあるか（=90-|緯度|の近似）。0〜90°。"),
    dict(key="earth_elev_deg", label="地球の仰角", unit="°", cmap="RdBu_r", vmin=-90, vmax=90,
         desc="月から見た地球の、地平線からの高さ（角度）。正＝地球が地平線の上にある（表側）、負＝地平線の下にある（裏側）。"),
    dict(key="slope_deg", label="全球の傾斜", unit="°", cmap="YlOrRd", vmin=0, vmax=20,
         desc="地面の傾き。0°が水平で、大きいほど急。データのない極付近は透明（欠測）。"),
    dict(key="age_index", label="地質年代", unit="", cmap="cividis", vmin=1, vmax=5,
         desc="USGS統合地質図の相対年代。1=最古、5=最新。"),
    dict(key="elev_m", label="標高", unit="m", cmap="terrain", vmin=None, vmax=None,
         desc="基準球（半径1737.4km）からの高さ。低いほど青、高いほど白（地球の地形図と同じ配色）。"),
]


def _grid(df: pd.DataFrame, col: str) -> np.ndarray:
    """行0=北緯89.5、列0=西経179.5 の (180, 360) グリッドにする。

    pivot_table は既定で「全部 NaN の行/列」を落とす（dropna=True）ため、
    slope_deg のように緯度帯まるごと欠測な列があると行数が減ってしまう。
    全緯度・全経度を明示して reindex し、形を (180, 360) に固定する。
    """
    lats = np.sort(df["lat"].unique())[::-1]   # 北緯89.5 → 南緯89.5
    lons = np.sort(df["lon"].unique())          # 西経179.5 → 東経179.5
    g = df.pivot_table(index="lat", columns="lon", values=col, dropna=False)
    g = g.reindex(index=lats, columns=lons)
    return g.values


def _wrap_zoom(arr: np.ndarray, order: int, mode: str, upscale: int) -> np.ndarray:
    """経度方向の周期性を考慮して zoom する（継ぎ目が出ないよう両端を回り込ませる）。"""
    n_lat, n_lon = arr.shape
    pad = 3
    padded = np.concatenate([arr[:, -pad:], arr, arr[:, :pad]], axis=1)
    up_padded = zoom(padded, (upscale, upscale), order=order, mode=mode)
    crop = pad * upscale
    return up_padded[:, crop:crop + n_lon * upscale]


def _to_rgba_png(grid: np.ndarray, cmap_name: str, vmin: float, vmax: float, path: str,
                  upscale: int = UPSCALE) -> None:
    """グリッドを補間して滑らかにし、カラーマップで着色して PNG にする。
    NaN（欠測）は別扱い：cubic 補間に NaN を渡すと周辺まで壊れるので、
    値は最近傍で埋めてから補間し、透明度（アルファ）だけ NaN マスクの最近傍拡大で決める。
    """
    nan_mask = np.isnan(grid)
    filled = grid
    if nan_mask.any():
        # 欠測は全体平均で仮埋め（cubic 補間が NaN で壊れるのを防ぐだけが目的。表示はアルファ0で消す）
        filled = np.where(nan_mask, np.nanmean(grid), grid)

    up = _wrap_zoom(filled, order=3, mode="nearest", upscale=upscale)
    norm = np.clip((up - vmin) / (vmax - vmin), 0, 1)
    rgba = (matplotlib.colormaps[cmap_name](norm) * 255).astype(np.uint8)  # (H, W, 4)

    if nan_mask.any():
        up_mask = _wrap_zoom(nan_mask.astype(np.float64), order=0, mode="nearest", upscale=upscale) > 0.5
        rgba[up_mask, 3] = 0

    from PIL import Image
    Image.fromarray(rgba, mode="RGBA").save(path)


def _gradient_css(cmap_name: str, n: int = 8) -> str:
    """凡例バー用の CSS linear-gradient を、実際に使うカラーマップから生成する。"""
    cmap = matplotlib.colormaps[cmap_name]
    stops = []
    for i in range(n):
        r, g, b, _ = cmap(i / (n - 1))
        stops.append(f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}")
    return "linear-gradient(to right, " + ", ".join(stops) + ")"


def main() -> None:
    df = pd.read_csv(os.path.join(DATA, "site_environment.csv"))

    age_fine = None
    if os.path.exists(AGE_FINE_GRID):
        age_fine = pd.read_csv(AGE_FINE_GRID)
        print(f"age_index: 高解像度グリッドを使用（{AGE_FINE_GRID}、{len(age_fine)}点）")
    else:
        print(f"age_index: 高解像度グリッドが無いため site_environment.csv の1度グリッドで代用します。\n"
              f"  上げたい場合は tools/build_geology_grid.py --step 0.25 "
              f"--out {AGE_FINE_GRID} --write を実行してください（要 USGS GISデータ224MB）。")

    meta = []
    for layer in LAYERS:
        col = layer["key"]
        upscale = UPSCALE
        source = "data/site_environment.csv"
        if col == "age_index" and age_fine is not None:
            grid = _grid(age_fine, col)
            upscale = AGE_FINE_UPSCALE
            source = "USGS Unified Geologic Map of the Moon（0.25°グリッド。data/moon_geology_grid.csv " \
                      "と同じ元データをtools/build_geology_grid.py --step 0.25で再処理した中間生成物）"
        else:
            grid = _grid(df, col)
        vmin = layer["vmin"] if layer["vmin"] is not None else float(np.nanmin(grid))
        vmax = layer["vmax"] if layer["vmax"] is not None else float(np.nanmax(grid))
        out_path = os.path.join(TEX_OUT, f"overlay_{col}.png")
        _to_rgba_png(grid, layer["cmap"], vmin, vmax, out_path, upscale=upscale)
        n_nan = int(np.isnan(grid).sum())
        print(f"wrote {out_path}  ({os.path.getsize(out_path)/1024:.0f} KB)  "
              f"range [{vmin:.1f}, {vmax:.1f}] {layer['unit']}  欠測 {n_nan} セル  "
              f"grid {grid.shape}")
        meta.append({
            "key": col, "label": layer["label"], "unit": layer["unit"],
            "desc": layer["desc"],
            "min": round(vmin, 1), "max": round(vmax, 1),
            "texture": f"textures/overlay_{col}.png",
            "gradientCss": _gradient_css(layer["cmap"]),
            "source": source,
        })

    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print("wrote", JSON_OUT)


if __name__ == "__main__":
    main()
