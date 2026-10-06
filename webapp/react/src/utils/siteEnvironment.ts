/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 *
 * data/site_environment.csv（全球1度グリッドの環境指標）を粗くした JSON を読み、
 * 「クリックした地点にいちばん近いセル」を返す。ノートブックの nearest() の入口版。
 * JSON は webapp/react/scripts/gen_site_env.py が生成する（手で編集しない）。
 */
import raw from '../data/siteEnvironment.generated.json';

export interface SiteEnv {
  lat: number;
  lon: number;
  tempAmp: number;      // 1日の温度差 [K]
  nightMin: number;     // 夜の最低温度 [K]
  noonSun: number;      // 正午の太陽高度 [度]
  earthElev: number;    // 地球の仰角 [度]（正=表側 / 負=裏側）
  terrain: string;      // 海 / 陸
  slopeDeg: number | null;  // 全球の傾斜 [度]（元データのないセルは欠測。実際には緯度±88.5° のセルだけ）
  ageIndex: number | null;  // 相対地質年代 1(古)〜5(新)
  elevM: number;        // 標高 [m]（基準球 R=1737.4km からの高さ。マイナス＝低地）
}

type EnvCell = [number, number, number, number, number, number, string, number | null, number | null, number];
const STEP: number = (raw as { step: number }).step;
const CELLS = (raw as unknown as { cells: EnvCell[] }).cells;

function wrapLonDelta(a: number, b: number): number {
  let d = Math.abs(a - b) % 360;
  if (d > 180) d = 360 - d;
  return d;
}

/** lat, lon（度、lon は -180..180）にいちばん近いグリッドセルを返す。 */
export function nearestEnv(lat: number, lon: number): SiteEnv | null {
  if (!CELLS.length) return null;
  const lonN = ((lon + 180) % 360 + 360) % 360 - 180;
  let best: EnvCell | null = null;
  let bestD = Infinity;
  for (const c of CELLS) {
    const dLat = Math.abs(c[0] - lat);
    if (dLat > STEP * 2 && best) continue;
    const dLon = wrapLonDelta(c[1], lonN) * Math.cos((lat * Math.PI) / 180);
    const d = dLat * dLat + dLon * dLon;
    if (d < bestD) {
      bestD = d;
      best = c;
    }
  }
  if (!best) return null;
  return {
    lat: best[0], lon: best[1], tempAmp: best[2], nightMin: best[3],
    noonSun: best[4], earthElev: best[5], terrain: best[6],
    slopeDeg: best[7], ageIndex: best[8], elevM: best[9]
  };
}

export function earthSideLabel(earthElev: number): string {
  if (earthElev > 10) return '表側（地球が地平線の上に見える）';
  if (earthElev < -10) return '裏側（地球が地平線の下で見えない）';
  return '表裏の境（地球が地平線すれすれ）';
}
