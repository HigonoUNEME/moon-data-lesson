import React from 'react';
import { Moon, Table2 } from 'lucide-react';

/**
 * 画面いちばん上の細い帯。タイトルと「データ一覧」への導線（公開最小版 2026-09-18）。
 * ノートブック（Colab/JupyterLite）はまだ用意していないため、そちらへの導線は含めない。
 * 用意でき次第、分析ページへのリンクをここに戻す。
 */
export const Header: React.FC = () => {
  return (
    <header className="shrink-0 border-b border-slate-800/80 bg-slate-950/95 backdrop-blur-md">
      <div className="flex items-center justify-between gap-3 px-4 sm:px-5 py-2.5">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="flex items-center justify-center w-8 h-8 rounded-lg bg-gradient-to-br from-slate-800 to-cyan-950 border border-cyan-500/30 shrink-0">
            <Moon className="w-4 h-4 text-cyan-400" />
          </div>
          <div className="min-w-0">
            <h1 className="text-sm font-bold text-white truncate">
              月データでムーンベースの場所を決めよう
            </h1>
            <p className="text-[11px] text-slate-400 truncate">
              3D月球儀で月のデータを見て、場所ごとのちがいを読みとろう
            </p>
          </div>
        </div>

        <nav className="flex items-center gap-2 shrink-0">
          <a
            href="./data.html"
            // 別タブで開く：同じタブだと月の向き・ピン・層の選択が消えるため
            target="_blank"
            rel="noopener noreferrer"
            title="データ一覧（新しいタブで開きます）"
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-cyan-500 hover:bg-cyan-400 text-slate-950 text-xs font-bold transition-colors"
          >
            <Table2 className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">データ一覧</span>
            <span className="sm:hidden">CSV</span>
          </a>
        </nav>
      </div>
    </header>
  );
};
