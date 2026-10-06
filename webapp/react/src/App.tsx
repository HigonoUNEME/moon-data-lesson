/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useCallback } from 'react';
import { LUNAR_DATASETS } from './data/lunarData';
import { LunarFeature, MoonViewerSettings } from './types';
import { MoonViewer3D } from './components/MoonViewer3D';
import { Header } from './components/Header';
import { CheckCircle2 } from 'lucide-react';

export default function App() {
  // 公開最小版：右パネル（データセット選択・CSVダウンロード）は「データ一覧」ページ
  // （./data.html）に譲り、3D月球儀だけを画面いっぱいに出す。
  // activeDatasetId はクリックした地点のカテゴリ追従にだけ使う（表示自体はもうしない）。
  const [activeDatasetId, setActiveDatasetId] = useState<string>('landing-sites');
  // 既定では何も選択しない（勝手にアポロ11号が「選択中」になっているのは不要という指摘への対応）。
  // ユーザーが月をクリックして選ぶまで null のまま。
  const [selectedFeature, setSelectedFeature] = useState<LunarFeature | null>(null);
  const [toast, setToast] = useState<string | null>(null);

  const [settings, setSettings] = useState<MoonViewerSettings>({
    autoRotate: false,
    // 授業で観察しやすい速さ（×1.0＝1周約60秒。以前は×0.35＝約170秒）
    rotationSpeed: 1.0,
    wireframe: false,
    lightIntensity: 2.2,
    moonRotationDeg: 0,
    selectedCategoryFilter: 'all',
    // データ層は既定でオフ（まず月そのものを見せる）。「1日の温度」は重ねたときの既定の層として残す
    showDataLayer: false,
    dataLayerKey: 'diurnal_temp',
    showMoonTexture: true,
    // 地点名ヒント（既知地点の名前・カテゴリ）は既定でオフ。必要なときだけ画面のスイッチでオンにする
    showFeatureHints: false
  });

  const handleUpdateSettings = useCallback((patch: Partial<MoonViewerSettings>) => {
    setSettings(prev => ({ ...prev, ...patch }));
  }, []);

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 3600);
  }, []);

  const activeDataset = useMemo(
    () => LUNAR_DATASETS.find(d => d.id === activeDatasetId) || LUNAR_DATASETS[0],
    [activeDatasetId]
  );

  // 3D月球儀は常設ピンを出さず、カーソルを合わせたときの名前判定・クリック選択のためだけに
  // 全データセットの地点を渡す（どのパネルを開いていても、近くの地点名が出るようにする）。
  const allFeatures = useMemo(() => {
    const seen = new Set<string>();
    const out: LunarFeature[] = [];
    for (const ds of LUNAR_DATASETS) {
      for (const f of ds.data) {
        if (!seen.has(f.id)) {
          seen.add(f.id);
          out.push(f);
        }
      }
    }
    return out;
  }, []);

  const handleSelectFeature = useCallback((feature: LunarFeature) => {
    setSelectedFeature(feature);
    const parent = LUNAR_DATASETS.find(d => d.data.some(x => x.id === feature.id));
    if (parent && parent.id !== activeDatasetId) setActiveDatasetId(parent.id);
  }, [activeDatasetId]);

  return (
    <div className="h-screen flex flex-col bg-slate-950 text-slate-100 font-sans overflow-hidden selection:bg-cyan-500/30">
      {toast && (
        <div
          id="global-toast"
          className="fixed top-16 left-1/2 -translate-x-1/2 z-50 flex items-center gap-2.5 bg-emerald-950/95 border border-emerald-500/50 text-emerald-200 px-4 py-2.5 rounded-xl shadow-2xl backdrop-blur-md text-xs font-medium"
        >
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          {toast}
        </div>
      )}

      <Header />

      {/* 公開最小版：3D月球儀を画面いっぱいに。データセット選択・CSVダウンロードは
          「データ一覧」ページ（Headerからリンク）に譲る。 */}
      <div className="flex-1 relative overflow-hidden">
        <MoonViewer3D
          features={allFeatures}
          selectedFeature={selectedFeature}
          onSelectFeature={handleSelectFeature}
          settings={settings}
          onUpdateSettings={handleUpdateSettings}
          fill
        />
      </div>
    </div>
  );
}
