
import React, { useState, useCallback, useMemo } from 'react';
import { 
  Upload, 
  FileCheck, 
  Table as TableIcon, 
  Trash2,
  Layers,
  ChevronDown,
  Presentation,
  AlertCircle,
  FileText,
  Activity,
  Plus,
  // Fix: Added CheckCircle2 to imports
  CheckCircle2
} from 'lucide-react';
import { MappingGuide } from './components/MappingGuide';
import { suggestMappings } from './services/geminiService';
import { generateReportPPTX } from './services/presentationService';
import { DataRow, MappingField, Template, TEMPLATES, DEFAULT_TEMPLATE, Dataset } from './types';

const App: React.FC = () => {
  const [datasets, setDatasets] = useState<Dataset[]>([]);
  const [selectedTemplate, setSelectedTemplate] = useState<Template>(DEFAULT_TEMPLATE);
  const [allMappings, setAllMappings] = useState<{ [key: string]: MappingField[] }>({});
  const [isLoading, setIsLoading] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const hasMinuteData = useMemo(() => datasets.some(ds => ds.type === 'session'), [datasets]);
  const hasSummaryData = useMemo(() => datasets.some(ds => ds.type === 'summary'), [datasets]);

  const mappings = useMemo(() => allMappings[selectedTemplate.id] || [], [allMappings, selectedTemplate]);

  const runAutoMappingForDataset = useCallback(async (currentHeaders: string[]) => {
    const results: { [key: string]: MappingField[] } = {};
    for (const template of TEMPLATES) {
      results[template.id] = await suggestMappings(currentHeaders, template);
    }
    setAllMappings(prev => ({ ...prev, ...results }));
  }, []);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    setIsLoading(true);
    const newDatasets: Dataset[] = [];

    try {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const content = await file.text();
        const json = JSON.parse(content);
        const rows = Array.isArray(json) ? json : [json];
        
        if (rows.length > 0) {
          const headers = Object.keys(rows[0]);
          const isSession = headers.includes('経過時間 (分)') || headers.includes('Minute');
          const isSummary = headers.includes('Live Stream Status') || headers.includes('Unique Viewers Count');
          
          newDatasets.push({
            name: file.name,
            data: rows,
            headers: headers,
            type: isSession ? 'session' : (isSummary ? 'summary' : 'unknown')
          });

          await runAutoMappingForDataset(headers);
        }
      }
      setDatasets(prev => [...prev, ...newDatasets]);
    } catch (err) {
      alert("ファイルの解析に失敗しました。JSON形式を確認してください。");
    } finally {
      setIsLoading(false);
    }
  };

  const handleTemplateChange = (templateId: string) => {
    const template = TEMPLATES.find(t => t.id === templateId) || DEFAULT_TEMPLATE;
    setSelectedTemplate(template);
  };

  const handleReset = () => {
    setDatasets([]);
    setAllMappings({});
    setSelectedTemplate(DEFAULT_TEMPLATE);
  };

  const handleMappingChange = (targetKey: string, sourceKey: string) => {
    setAllMappings(prev => ({
      ...prev,
      [selectedTemplate.id]: prev[selectedTemplate.id].map(m => 
        m.targetKey === targetKey ? { ...m, sourceKey, confidence: 1 } : m
      )
    }));
  };

  const handleDownloadPPTX = async () => {
    if (datasets.length === 0) return;
    setIsLoading(true);
    try {
      await generateReportPPTX(datasets, allMappings);
    } catch (error) {
      console.error("PPTX Generation Error:", error);
      const errorMessage = error instanceof Error ? error.message : String(error);
      alert(`レポートの生成中にエラーが発生しました。\n\n詳細: ${errorMessage}`);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen pb-20">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <div className="flex items-center gap-3">
              <div className="bg-orange-600 text-white p-2 rounded-lg shadow-lg">
                <Presentation className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-slate-900 leading-tight">Firework Slide Automator</h1>
                <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">Smart Multi-Report Generator</p>
              </div>
            </div>
            {datasets.length > 0 && (
              <button 
                onClick={handleReset}
                className="text-slate-400 hover:text-red-600 flex items-center gap-1.5 text-sm font-medium transition-colors p-2 rounded-lg hover:bg-red-50"
              >
                <Trash2 className="w-4 h-4" /> リセット
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <MappingGuide />

        <div className="bg-white border-2 border-dashed border-slate-300 rounded-3xl p-10 text-center transition-all hover:border-orange-400 hover:bg-orange-50/10 group mb-8">
          <div className="max-w-md mx-auto">
            <div className="bg-orange-100 text-orange-600 w-16 h-16 rounded-2xl flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform shadow-inner">
              <Upload className="w-8 h-8" />
            </div>
            <h3 className="text-xl font-bold text-slate-900 mb-2">JSONデータを追加アップロード</h3>
            <p className="text-slate-500 mb-6 text-sm leading-relaxed">
              複数のファイルを一度に、または順次追加できます。<br/>
              <b>SUMMARY</b> と <b>分データ</b> を両方入れると全スライドを生成します。
            </p>
            <label className="inline-flex items-center gap-3 bg-slate-900 text-white px-6 py-3 rounded-xl font-bold hover:bg-slate-800 cursor-pointer shadow-lg transition-all active:scale-95">
              <Plus className="w-5 h-5" />
              ファイルを追加選択
              <input type="file" className="hidden" accept=".json" multiple onChange={handleFileUpload} />
            </label>
          </div>
        </div>

        {datasets.length > 0 && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
              <div className="flex items-center gap-4">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <h2 className="text-2xl font-bold text-slate-900">
                      検出済み: {datasets.length}個のデータセット
                    </h2>
                  </div>
                  <div className="flex gap-2">
                    {hasSummaryData && <span className="bg-blue-100 text-blue-700 text-[10px] px-2 py-1 rounded-md font-bold uppercase">P.4, P.5 対応</span>}
                    {hasMinuteData && <span className="bg-purple-100 text-purple-700 text-[10px] px-2 py-1 rounded-md font-bold uppercase">P.8-12 対応</span>}
                  </div>
                </div>
              </div>
              <button 
                onClick={handleDownloadPPTX}
                className="bg-orange-600 text-white hover:bg-orange-700 flex items-center justify-center gap-3 px-8 py-4 rounded-xl font-bold transition-all shadow-xl shadow-orange-200 hover:scale-105 active:scale-95 text-lg"
              >
                <Presentation className="w-6 h-6" /> PPTXをダウンロード
              </button>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              <div className="lg:col-span-1 space-y-4">
                <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden sticky top-24">
                  <div className="p-4 border-b border-slate-100 bg-slate-50/50">
                    <h3 className="font-bold text-slate-800 flex items-center gap-2 text-sm uppercase tracking-wider">
                      <Layers className="w-4 h-4 text-orange-600" />
                      マッピング確認
                    </h3>
                  </div>
                  
                  <div className="p-2 bg-slate-50 border-b flex flex-wrap gap-1">
                    {TEMPLATES.map(t => (
                      <button
                        key={t.id}
                        onClick={() => handleTemplateChange(t.id)}
                        className={`px-3 py-2 rounded-lg text-[10px] font-bold transition-all ${
                          selectedTemplate.id === t.id ? 'bg-white text-orange-600 shadow-sm ring-1 ring-slate-200' : 'text-slate-500 hover:bg-slate-100'
                        }`}
                      >
                        {t.name}
                      </button>
                    ))}
                  </div>

                  <div className="p-4 space-y-4 max-h-[60vh] overflow-y-auto custom-scrollbar">
                    {mappings.map((m) => (
                      <div key={m.targetKey} className="group">
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex justify-between mb-1">
                          <span>{m.label}</span>
                        </label>
                        <select
                          value={m.sourceKey}
                          onChange={(e) => handleMappingChange(m.targetKey, e.target.value)}
                          className="w-full text-xs rounded-lg border border-slate-200 p-2 bg-white outline-none transition-all focus:border-orange-400"
                        >
                          <option value="">-- 未選択 --</option>
                          {datasets.flatMap(ds => ds.headers).filter((v, i, a) => a.indexOf(v) === i).map(h => (
                            <option key={h} value={h}>{h}</option>
                          ))}
                        </select>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="lg:col-span-2 space-y-6">
                <div className="bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden">
                  <div className="p-6 border-b border-slate-100 bg-slate-50/30">
                    <h3 className="font-bold text-slate-900 flex items-center gap-2">
                      <FileText className="w-5 h-5 text-slate-400" />
                      アップロード済みファイル
                    </h3>
                  </div>
                  <div className="p-4">
                    {datasets.map((ds, idx) => (
                      <div key={idx} className="flex items-center justify-between p-3 border-b last:border-0 hover:bg-slate-50 transition-colors rounded-xl">
                        <div className="flex items-center gap-3">
                          <div className={`p-2 rounded-lg ${ds.type === 'session' ? 'bg-purple-100 text-purple-600' : 'bg-blue-100 text-blue-600'}`}>
                            {ds.type === 'session' ? <Activity className="w-4 h-4" /> : <FileText className="w-4 h-4" />}
                          </div>
                          <div>
                            <p className="text-sm font-bold text-slate-800">{ds.name}</p>
                            <p className="text-[10px] text-slate-500">{ds.data.length}行のデータ / {ds.type === 'session' ? '分単位データ' : 'サマリーデータ'}</p>
                          </div>
                        </div>
                        <CheckCircle2 className="w-5 h-5 text-green-500" />
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {isLoading && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-50 flex items-center justify-center p-6 text-center">
          <div className="bg-white rounded-3xl p-10 max-w-sm w-full shadow-2xl space-y-6">
            <div className="relative w-20 h-20 mx-auto">
              <div className="absolute inset-0 border-4 border-orange-50 rounded-full"></div>
              <div className="absolute inset-0 border-4 border-orange-600 rounded-full border-t-transparent animate-spin"></div>
            </div>
            <div>
              <h3 className="text-xl font-bold text-slate-900 mb-2">処理中...</h3>
              <p className="text-sm text-slate-500 leading-relaxed">AI analysis and report compilation are in progress. As multiple files are being processed, it may take a few moments.</p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
