
import React, { useState } from 'react';
import { ChevronDown, ChevronUp, AlertTriangle, ExternalLink, FileJson, Settings, Presentation } from 'lucide-react';

export const MappingGuide: React.FC = () => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 mb-6 transition-all duration-200">
      <button
        onClick={() => setIsOpen(!isOpen)}
        className="flex items-center justify-between w-full text-left focus:outline-none group"
      >
        <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2 group-hover:text-blue-600 transition-colors">
          <span className="bg-blue-100 p-1.5 rounded-lg text-blue-600">📚</span>
          ツールの使い方 & 留意事項
        </h2>
        {isOpen ? <ChevronUp className="w-5 h-5 text-slate-400" /> : <ChevronDown className="w-5 h-5 text-slate-400" />}
      </button>

      {isOpen && (
        <div className="mt-4 space-y-6 border-t border-slate-100 pt-5">
          <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
            <div className="bg-blue-50/50 p-4 rounded-xl border border-blue-100">
              <div className="flex items-center gap-2 mb-3">
                <div className="bg-blue-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold">1</div>
                <h3 className="font-bold text-blue-900">データ準備</h3>
              </div>
              <p className="text-sm text-blue-800 mb-3 leading-relaxed">
                SigmaからJSONを抽出します。英語・日本語どちらの項目名でも対応可能です。
              </p>
              <a
                href="https://app.sigmacomputing.com/firework-data/workbook/Channel-Overview-4XT6tSu62d7Q5QLLkqZMNB?:nodeId=AWfljii3Ij&:customView=1c9be963-6cfa-4859-833b-e8f05628b0fe"
                target="_blank"
                rel="noopener noreferrer"
                className="text-xs font-semibold bg-blue-600 text-white px-3 py-2 rounded-lg hover:bg-blue-700 inline-flex items-center gap-1 shadow-sm transition-all"
              >
                Sigmaを開く <ExternalLink className="w-3 h-3" />
              </a>
            </div>

            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
              <div className="flex items-center gap-2 mb-3">
                <div className="bg-slate-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold">2</div>
                <h3 className="font-bold text-slate-900">ファイル選択</h3>
              </div>
              <p className="text-sm text-slate-700 leading-relaxed flex items-start gap-2">
                <FileJson className="w-4 h-4 mt-0.5 text-slate-500 shrink-0" />
                抽出したJSONファイルをアップロードしてください。
              </p>
            </div>

            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
              <div className="flex items-center gap-2 mb-3">
                <div className="bg-slate-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold">3</div>
                <h3 className="font-bold text-slate-900">項目確認</h3>
              </div>
              <p className="text-sm text-slate-700 leading-relaxed flex items-start gap-2">
                <Settings className="w-4 h-4 mt-0.5 text-slate-500 shrink-0" />
                AIが自動マッピングします。必要に応じて手動で調整してください。
              </p>
            </div>

            <div className="bg-orange-50/50 p-4 rounded-xl border border-orange-100">
              <div className="flex items-center gap-2 mb-3">
                <div className="bg-orange-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold">4</div>
                <h3 className="font-bold text-orange-900">資料作成</h3>
              </div>
              <p className="text-sm text-orange-800 mb-3 leading-relaxed flex items-start gap-2">
                <Presentation className="w-4 h-4 mt-0.5 text-orange-600 shrink-0" />
                <span>
                  「プレゼン資料DL」をクリックして、P.4/P.5を含むPPTXを生成します。
                </span>
              </p>
            </div>
          </div>

          <div className="bg-red-50 p-4 rounded-xl border border-red-100">
             <h3 className="text-sm font-bold text-red-700 flex items-center gap-1.5 mb-2">
               <AlertTriangle className="w-4 h-4" /> 留意事項
             </h3>
             <p className="text-xs text-red-800 leading-relaxed">
               自動マッピングされた項目（特に「要確認」表示）は、出力前に必ず正しいソースデータが選択されているか確認してください。
             </p>
          </div>
        </div>
      )}
    </div>
  );
};
