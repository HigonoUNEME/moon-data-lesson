import React, { useMemo, useState } from 'react';
import { LunarDataset, LunarFeature } from '../types';
import { downloadDatasetAsCSV } from '../utils/csvExport';
import { nearestEnv, earthSideLabel } from '../utils/siteEnvironment';
import {
  Search,
  Download,
  Copy,
  MapPin,
  ChevronRight,
  ChevronDown,
  NotebookPen,
  Table2,
  Compass
} from 'lucide-react';

interface SidePanelProps {
  datasets: LunarDataset[];
  activeDatasetId: string;
  onSelectDatasetId: (id: string) => void;
  selectedFeature: LunarFeature | null;
  onSelectFeature: (feature: LunarFeature) => void;
  onToast: (msg: string) => void;
}

// MoonViewer3D のピン色に合わせる
const CAT_DOT: Record<string, string> = {
  missions: 'bg-emerald-400',
  craters: 'bg-amber-400',
  maria: 'bg-cyan-400',
  moonquakes: 'bg-rose-400',
  resources: 'bg-indigo-400'
};

// カード上部に固定で出す列（重複表示を避けるため、下のチップ生成からは外す）
const SKIP_KEYS = new Set(['nameJa', 'name', 'description', 'latitude', 'longitude', 'agency', 'year']);

function valueOf(feature: LunarFeature, key: string): string | number | undefined {
  if (key in feature) return (feature as unknown as Record<string, string | number>)[key];
  if (feature.attributes && key in feature.attributes) {
    return feature.attributes[key] as string | number;
  }
  return undefined;
}

function fmtCoord(n: number): string {
  return `${n > 0 ? '+' : ''}${n.toFixed(2)}°`;
}

const NOTEBOOKS = [
  { path: 'course_moonbase.ipynb', label: 'ガイド型：ムーンベースの最適地', hint: 'ステップ1〜5。クラス1コマ＋ワークシート', primary: true },
  { path: 'petit_inquiry_start.ipynb', label: 'オープン型：自分で問いを立てる', hint: '最小の出発点。課題ブリーフを見ながら' },
  { path: 'explore.ipynb', label: '道具：自由探索ツール', hint: 'コードなしでデータ・軸・色を選んで散布図' },
  { path: 'explore_clustering.ipynb', label: '発展：クラスタリングで仲間分け', hint: '「周りと違う場所」を探す' },
  { path: 'explore_advanced.ipynb', label: '発展：クレーター数から絶対年代', hint: 'べき乗則フィットと参照表' },
  { path: 'course_moonbase_polar.ipynb', label: 'ガイド型の分岐：南極を細かく見る', hint: '南極付近を、日照率・傾斜・永久影で見る' },
  { path: 'course_moonbase_ml.ipynb', label: '任意：機械学習でモデル比較', hint: '5モデルの決定境界・正解率' }
];

export const SidePanel: React.FC<SidePanelProps> = ({
  datasets,
  activeDatasetId,
  onSelectDatasetId,
  selectedFeature,
  onSelectFeature,
  onToast
}) => {
  const [query, setQuery] = useState('');
  // 一覧は3D月球儀のクリック選択と役割が重なるため、既定では畳んでおく（検索中・手動で開いたときだけ展開）
  const [listExpanded, setListExpanded] = useState(false);

  const activeDataset = useMemo(
    () => datasets.find(d => d.id === activeDatasetId) || datasets[0],
    [datasets, activeDatasetId]
  );

  // カード内に出す「見どころ」の列（座標・名前・説明以外の先頭2つ）
  const statCols = useMemo(
    () => activeDataset.columns.filter(c => !SKIP_KEYS.has(c.key)).slice(0, 3),
    [activeDataset]
  );

  const list = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return activeDataset.data;
    return activeDataset.data.filter(f =>
      f.nameJa.toLowerCase().includes(q) ||
      f.name.toLowerCase().includes(q) ||
      f.description.toLowerCase().includes(q) ||
      (f.agency?.toLowerCase().includes(q) ?? false)
    );
  }, [activeDataset, query]);

  const totalPoints = datasets.reduce((a, d) => a + d.data.length, 0);
  // 検索中は自動で一覧を出す（手動トグルは検索していないときだけ意味を持つ）
  const showList = listExpanded || query.trim().length > 0;

  // 選択中の地点の詳細（説明・タグ・環境データ・座標コピー）。一覧の行の下にも、
  // 一覧を畳んでいるときの単独カードとしても使う共通部品
  const renderFeatureDetail = (f: LunarFeature) => {
    const env = nearestEnv(f.latitude, f.longitude);
    return (
      <div className="flex flex-col gap-2">
        <p className="text-[11px] text-slate-300 leading-relaxed">{f.description}</p>

        <div className="flex flex-wrap gap-1.5 text-[10px]">
          <span className="px-2 py-0.5 rounded bg-slate-800 border border-slate-700 font-mono text-slate-200">
            {f.name}
          </span>
          {f.agency && (
            <span className="px-2 py-0.5 rounded bg-emerald-950/60 border border-emerald-500/30 text-emerald-300">
              {f.agency}
            </span>
          )}
          {f.year && (
            <span className="px-2 py-0.5 rounded bg-indigo-950/60 border border-indigo-500/30 text-indigo-300">
              {f.year}
            </span>
          )}
          {statCols.map(col => {
            const v = valueOf(f, col.key);
            if (v === undefined || v === '' || v === null) return null;
            return (
              <span key={col.key} className="px-2 py-0.5 rounded bg-slate-800 border border-slate-700 text-slate-300">
                {col.label.replace(/\s*\(.*$/, '')}: {String(v)}{col.unit ?? ''}
              </span>
            );
          })}
        </div>

        {env && (
          <div id="feature-env" className="rounded-lg bg-slate-900/60 border border-slate-700/70 p-2 flex flex-col gap-1">
            <div className="text-[10px] font-bold text-slate-400">
              この地点の環境（<code className="text-cyan-400">site_environment.csv</code> の最寄りマス）
            </div>
            <div className="grid grid-cols-2 gap-x-3 gap-y-0.5 text-[10px] text-slate-300 font-mono">
              <span>1日の温度差: {env.tempAmp.toFixed(0)} K</span>
              <span>夜の底: {env.nightMin.toFixed(0)} K</span>
              <span>正午の太陽高度: {env.noonSun.toFixed(0)}°</span>
              <span>地球の仰角: {env.earthElev > 0 ? '+' : ''}{env.earthElev.toFixed(0)}°</span>
              <span>地形: {env.terrain || '—'}</span>
            </div>
            <div className="text-[10px] text-slate-400">{earthSideLabel(env.earthElev)}</div>
          </div>
        )}

        <button
          onClick={() => {
            navigator.clipboard?.writeText(`${f.latitude}, ${f.longitude}`);
            onToast(`座標 ${f.latitude}, ${f.longitude} をコピーしました`);
          }}
          className="self-start flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 text-[11px] transition-colors"
        >
          <Copy className="w-3 h-3" />
          座標をコピー
        </button>
      </div>
    );
  };

  return (
    <div className="flex flex-col gap-5 p-4 sm:p-5 text-slate-200">
      {/* イントロ */}
      <p className="text-xs leading-relaxed text-slate-400">
        月の有名な地点 <strong className="text-slate-200">{totalPoints}件</strong> を見て回る紹介ページです。
        地点をえらぶと、左の月がその場所に向き、その場所の環境
        （1日の温度差・地球の仰角・地形＝<code className="text-cyan-400">site_environment.csv</code> の最寄りマス）も出ます。
        分析につかう観測データ（クレーター・温度・極域日照など）は
        <a href="./data.html" className="text-cyan-400 hover:text-cyan-300"> データ一覧</a>
        からダウンロードできます。
      </p>

      {/* ① データを選ぶ */}
      <section className="flex flex-col gap-2">
        <h2 className="flex items-center gap-2 text-xs font-bold text-slate-300">
          <span className="flex w-5 h-5 items-center justify-center rounded-full bg-slate-800 text-cyan-400 text-[11px] font-black">1</span>
          データを選ぶ
        </h2>
        <div className="flex flex-col gap-1.5">
          {datasets.map(ds => {
            const active = ds.id === activeDatasetId;
            return (
              <button
                key={ds.id}
                id={`dataset-pick-${ds.id}`}
                onClick={() => { onSelectDatasetId(ds.id); setQuery(''); setListExpanded(false); }}
                className={`w-full text-left rounded-xl border px-3 py-2.5 transition-colors ${
                  active
                    ? 'bg-cyan-950/40 border-cyan-500/50'
                    : 'bg-slate-900/50 border-slate-800 hover:border-slate-700'
                }`}
              >
                <div className="flex items-center gap-2">
                  <span className={`w-2 h-2 rounded-full shrink-0 ${CAT_DOT[ds.category] ?? 'bg-slate-400'}`} />
                  <span className={`text-xs font-bold ${active ? 'text-cyan-200' : 'text-slate-100'}`}>
                    {ds.titleJa}
                  </span>
                  <span className="ml-auto text-[10px] font-mono text-slate-400">{ds.data.length}件</span>
                </div>
                <p className="text-[11px] text-slate-400 mt-1 leading-snug line-clamp-2">{ds.description}</p>
              </button>
            );
          })}
        </div>
      </section>

      {/* ② 地点を選ぶ */}
      <section className="flex flex-col gap-2">
        <div className="flex items-center justify-between">
          <h2 className="flex items-center gap-2 text-xs font-bold text-slate-300">
            <span className="flex w-5 h-5 items-center justify-center rounded-full bg-slate-800 text-cyan-400 text-[11px] font-black">2</span>
            地点を選ぶ
          </h2>
          <button
            id="btn-panel-download-csv"
            onClick={() => {
              downloadDatasetAsCSV(activeDataset);
              onToast(`「${activeDataset.titleJa}」を CSV でダウンロードしました`);
            }}
            className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 border border-slate-700 text-[11px] font-medium transition-colors"
          >
            <Download className="w-3 h-3 text-cyan-400" />
            この{activeDataset.data.length}件をCSV
          </button>
        </div>

        <div className="relative">
          <Search className="w-3.5 h-3.5 text-slate-500 absolute left-2.5 top-1/2 -translate-y-1/2" />
          <input
            id="input-feature-search"
            type="text"
            value={query}
            onChange={e => setQuery(e.target.value)}
            placeholder="名前・ミッション・説明で絞り込み"
            className="w-full bg-slate-900 border border-slate-700/80 rounded-lg pl-8 pr-3 py-1.5 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-cyan-500"
          />
        </div>

        {showList ? (
          <div className="border border-slate-800 rounded-xl divide-y divide-slate-800/70 overflow-hidden max-h-[46vh] overflow-y-auto">
            {list.length === 0 ? (
              <p className="p-4 text-center text-xs text-slate-500">一致する地点がありません</p>
            ) : (
              list.map(f => {
                const selected = selectedFeature?.id === f.id;
                return (
                  <div key={f.id}>
                    <button
                      id={`feature-row-${f.id}`}
                      onClick={() => onSelectFeature(f)}
                      className={`w-full text-left px-3 py-2 flex items-center gap-2 transition-colors ${
                        selected ? 'bg-cyan-950/40' : 'hover:bg-slate-800/40'
                      }`}
                    >
                      <MapPin className={`w-3.5 h-3.5 shrink-0 ${selected ? 'text-cyan-400' : 'text-slate-600'}`} />
                      <div className="min-w-0 flex-1">
                        <div className={`text-xs font-medium truncate ${selected ? 'text-cyan-200' : 'text-slate-100'}`}>
                          {f.nameJa}
                        </div>
                        <div className="text-[10px] font-mono text-slate-500">
                          {fmtCoord(f.latitude)} , {fmtCoord(f.longitude)}
                        </div>
                      </div>
                      <ChevronRight className={`w-3.5 h-3.5 shrink-0 transition-transform ${selected ? 'rotate-90 text-cyan-400' : 'text-slate-600'}`} />
                    </button>

                    {selected && (
                      <div id="feature-detail" className="px-3 pb-3 pt-1 bg-cyan-950/20 border-t border-cyan-500/20">
                        {renderFeatureDetail(f)}
                      </div>
                    )}
                  </div>
                );
              })
            )}
          </div>
        ) : selectedFeature ? (
          // 一覧を畳んでいるときは、選んでいる地点のカードだけをコンパクトに出す
          // （3D月球儀をクリックして地点を選んだときも、ここに詳細が出る）
          <div id="feature-detail" className="border border-cyan-500/30 rounded-xl bg-cyan-950/20 p-3 flex flex-col gap-2">
            <div className="flex items-center gap-2">
              <MapPin className="w-3.5 h-3.5 shrink-0 text-cyan-400" />
              <div className="min-w-0 flex-1">
                <div className="text-xs font-medium truncate text-cyan-200">{selectedFeature.nameJa}</div>
                <div className="text-[10px] font-mono text-slate-500">
                  {fmtCoord(selectedFeature.latitude)} , {fmtCoord(selectedFeature.longitude)}
                </div>
              </div>
            </div>
            {renderFeatureDetail(selectedFeature)}
          </div>
        ) : (
          <p className="p-3 text-center text-[11px] text-slate-500 border border-dashed border-slate-800 rounded-xl">
            3D月球儀をクリックするか、下の一覧から地点を選んでください
          </p>
        )}

        {!query && (
          <button
            id="btn-toggle-feature-list"
            onClick={() => setListExpanded(v => !v)}
            className="flex items-center justify-center gap-1.5 py-1.5 rounded-lg text-[11px] text-slate-400 hover:text-slate-200 hover:bg-slate-900/60 transition-colors"
          >
            <ChevronDown className={`w-3.5 h-3.5 transition-transform ${listExpanded ? 'rotate-180' : ''}`} />
            {listExpanded ? '一覧を閉じる' : `${activeDataset.titleJa}の一覧を見る（${list.length}件）`}
          </button>
        )}

        <p className="flex items-center gap-1.5 text-[10px] text-slate-500">
          <Compass className="w-3 h-3" />
          月面の模様は表示用の近似です。座標は公開情報にもとづく概略値。
        </p>
      </section>

      {/* ③ 分析にすすむ */}
      <section id="analyze" className="flex flex-col gap-2 scroll-mt-4">
        <h2 className="flex items-center gap-2 text-xs font-bold text-slate-300">
          <span className="flex w-5 h-5 items-center justify-center rounded-full bg-slate-800 text-cyan-400 text-[11px] font-black">3</span>
          分析にすすむ
        </h2>
        <p className="text-[11px] text-slate-400 leading-relaxed">
          ノートブックはブラウザの中の Python で動きます（インストール不要・別タブ）。
          初回は読み込みに 1〜2 分。メニューの「▶▶」で最初から最後まで実行できます。
        </p>
        <div className="flex flex-col gap-1.5">
          {NOTEBOOKS.map(nb => (
            <a
              key={nb.path}
              href={`./app/notebooks/index.html?path=${nb.path}`}
              target="_blank"
              rel="noopener noreferrer"
              className={`flex items-start gap-2 rounded-xl border px-3 py-2 transition-colors ${
                nb.primary
                  ? 'bg-cyan-500/15 border-cyan-500/50 hover:border-cyan-400'
                  : 'bg-slate-900/50 border-slate-800 hover:border-slate-700'
              }`}
            >
              <NotebookPen className="w-3.5 h-3.5 text-cyan-400 mt-0.5 shrink-0" />
              <div>
                <div className="text-xs font-bold text-slate-100">{nb.label}</div>
                <div className="text-[10px] text-slate-400 leading-snug">{nb.hint}</div>
              </div>
            </a>
          ))}
          <a
            href="./data.html"
            className="flex items-start gap-2 rounded-xl border border-slate-800 bg-slate-900/50 px-3 py-2 hover:border-slate-700 transition-colors"
          >
            <Table2 className="w-3.5 h-3.5 text-cyan-400 mt-0.5 shrink-0" />
            <div>
              <div className="text-xs font-bold text-slate-100">データ一覧（CSV ダウンロード）</div>
              <div className="text-[10px] text-slate-400 leading-snug">分析に使う月データを行数・出典つきで</div>
            </div>
          </a>
          <a
            href="./shadow_sim.html"
            className="flex items-start gap-2 rounded-xl border border-slate-800 bg-slate-900/50 px-3 py-2 hover:border-slate-700 transition-colors"
          >
            <Compass className="w-3.5 h-3.5 text-cyan-400 mt-0.5 shrink-0" />
            <div>
              <div className="text-xs font-bold text-slate-100">永久影のできかた（シミュレーター）</div>
              <div className="text-[10px] text-slate-400 leading-snug">緯度・クレーターの深さ・太陽高度を動かして、なぜ極だけ永久影になるかを見る</div>
            </div>
          </a>
        </div>
      </section>

      <footer className="pt-3 border-t border-slate-800/80 text-[10px] text-slate-500 leading-relaxed">
        代表地点は USGS Gazetteer・IAU・各ミッション公式・Apollo 地震カタログ等の公開情報にもとづく概略値
        （検証記録 <code>docs/explorer_data_verification.md</code>）。
        月面テクスチャは NASA/LROC カラーモザイク（パブリックドメイン）。
      </footer>
    </div>
  );
};
