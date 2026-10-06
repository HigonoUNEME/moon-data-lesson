/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 *
 * 月球儀にカーソルを合わせたとき、「1日の温度（アニメーション）」レイヤーの値を出すための
 * 遅延読み込みルックアップ。JSON は webapp/react/scripts/gen_diurnal_lookup.py が
 * data/diviner_global.csv.gz から3度グリッドに間引いて生成する（各セルに現地時間0〜23時の
 * 実測カーブをそのまま持つ。色からの逆算はしない＝焼いたPNGフレームと同じ考え方）。
 *
 * PNG24枚（約8.6MB）と同様この層を選んだときだけ読みに行く（常にバンドルへ含めない）。
 */
type DiurnalLookupJson = { step: number; cols: string[]; cells: number[][] };

let cache: Promise<DiurnalLookupJson> | null = null;

/** 「1日の温度」レイヤーを選んだときに一度だけ呼ぶ。以後は同じ Promise を返す。 */
export function loadDiurnalLookup(): Promise<DiurnalLookupJson> {
  if (!cache) {
    cache = import('../data/diurnalLookup.generated.json').then(
      (m) => (m as unknown as { default: DiurnalLookupJson }).default ?? (m as unknown as DiurnalLookupJson)
    );
  }
  return cache;
}

function wrapLonDelta(a: number, b: number): number {
  let d = Math.abs(a - b) % 360;
  if (d > 180) d = 360 - d;
  return d;
}

/** lat, lon にいちばん近いセルの24時間カーブ（t00..t23）を返す。 */
function nearestCurve(data: DiurnalLookupJson, lat: number, lon: number): number[] | null {
  const lonN = ((lon + 180) % 360 + 360) % 360 - 180;
  let best: number[] | null = null;
  let bestD = Infinity;
  for (const c of data.cells) {
    const dLat = Math.abs(c[0] - lat);
    if (dLat > data.step * 2 && best) continue;
    const dLon = wrapLonDelta(c[1], lonN) * Math.cos((lat * Math.PI) / 180);
    const d = dLat * dLat + dLon * dLon;
    if (d < bestD) {
      bestD = d;
      best = c;
    }
  }
  return best ? best.slice(2) : null;
}

/** lat, lon にいちばん近いセルの24時間カーブ（現地時間0〜23時の温度[K]）。
 *  ピンカードの「24時間の最高・最低・差」用（自転の角度に関係なく、1日ぶんの全体を返す）。 */
export function diurnalCurveAt(data: DiurnalLookupJson, lat: number, lon: number): number[] | null {
  return nearestCurve(data, lat, lon);
}

/** lat, lon の地点で、太陽直下点の経度が subsolarLon のときの温度[K]。
 *  gen_diurnal_frames.py の位相補間（現地時間 = (12+(lon-太陽直下点経度)/15) mod 24 を
 *  整数時2点で線形補間）とまったく同じ式。 */
export function diurnalValueAt(
  data: DiurnalLookupJson, lat: number, lon: number, subsolarLon: number
): number | null {
  const curve = nearestCurve(data, lat, lon);
  if (!curve) return null;
  const loc = (((12 + (lon - subsolarLon) / 15) % 24) + 24) % 24;
  const lo = Math.floor(loc) % 24;
  const hi = (lo + 1) % 24;
  const frac = loc - Math.floor(loc);
  return curve[lo] * (1 - frac) + curve[hi] * frac;
}
