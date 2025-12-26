
import pptxgen from "pptxgenjs";
import { Dataset, DataRow, MappingField } from "../types";
import { generateInsight } from "./geminiService";

/**
 * 視覚的な定数定義（ご提示のイメージに準拠）
 */
const COLORS = {
  FIREWORK_ORANGE: "E64A19",
  FIREWORK_BLUE: "0066FF",
  FIREWORK_PINK: "FF3399",
  SLATE_900: "0F172A",
  SLATE_700: "334155",
  SLATE_100: "F1F5F9",
  WHITE: "FFFFFF",
  HEADER_GREY: "888888",
  BORDER_GREY: "CCCCCC",
  LIGHT_YELLOW: "FFF9C4",
  LIGHT_PINK: "FFEBEE",
  SOFT_BLUE: "E1F5FE"
};

/**
 * 数値をプレゼンテーション用にフォーマット
 */
const formatValue = (val: any, label: string): string => {
  if (val === undefined || val === null || val === "") return "-";
  if (typeof val === 'number') {
    if (label.includes('率') || label.includes('Rate')) {
      return (val * 100).toFixed(2) + "%";
    }
    if (label.includes('分') || label.includes('時間')) {
      const totalSeconds = Math.round(val * 60);
      const mins = Math.floor(totalSeconds / 60);
      const secs = totalSeconds % 60;
      return `${mins}分${secs.toString().padStart(2, '0')}秒`;
    }
    return val.toLocaleString();
  }
  return String(val);
};

/**
 * チャート上のピーク地点を特定（最大5箇所）
 */
const findSignificantPoints = (values: number[], labels: string[], count: number = 5) => {
  return values
    .map((val, idx) => ({ val, label: labels[idx], idx }))
    .sort((a, b) => b.val - a.val)
    .slice(0, count)
    .sort((a, b) => a.idx - b.idx);
};

/**
 * チャート上に「◯」注釈を追加するヘルパー
 */
const addChartOvalMarker = (
  pptx: pptxgen,
  slide: pptxgen.Slide,
  index: number,
  total: number,
  value: number,
  maxValue: number,
  chartArea: { x: number, y: number, w: number, h: number },
  number?: string
) => {
  const x = chartArea.x + (index / (total - 1)) * chartArea.w;
  const y = chartArea.y + chartArea.h - (value / (maxValue || 1)) * chartArea.h;

  // @ts-ignore - pptx.ShapeType ensures no crash in modern pptxgenjs
  slide.addShape(pptx.ShapeType.OVAL, {
    x: x - 0.15, y: y - 0.15, w: 0.3, h: 0.3,
    line: { color: COLORS.FIREWORK_ORANGE, width: 2 }
  });

  if (number) {
    slide.addText(number, {
      x: x - 0.1, y: y - 0.45, w: 0.2, h: 0.2,
      fontSize: 10, bold: true, align: "center", color: COLORS.FIREWORK_ORANGE
    });
  }
};

export const generateReportPPTX = async (datasets: Dataset[], allMappings: { [templateId: string]: MappingField[] }) => {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";

  const summaryDataset = datasets.find(ds => ds.type === 'summary');
  const sessionDataset = datasets.find(ds => ds.type === 'session');

  // ユニーク視聴者数（分単位平均の計算母数として使用）
  const activeRow = summaryDataset?.data.find(d => String(d['Live Stream Status']).toLowerCase() === 'active');
  const liveUU = Number(activeRow?.['Unique Viewers Count']) || 1;

  // --- 1. DIGITAL SHOWROOM SUMMARY (P.4 - P.5) ---
  if (summaryDataset) {
    const data = summaryDataset.data;
    const groupedData: { [key: string]: { active?: DataRow, replay?: DataRow, date: Date } } = {};
    
    data.forEach(row => {
      const dateStr = row['Event Date Jp'] || row['Started At Jp'] || row['Date'];
      if (!dateStr) return;
      const dateKey = new Date(dateStr).toISOString().split('T')[0];
      const status = String(row['Live Stream Status'] || "").toLowerCase();
      
      if (!groupedData[dateKey]) groupedData[dateKey] = { date: new Date(dateStr) };
      if (status === 'active') groupedData[dateKey].active = row;
      else groupedData[dateKey].replay = row;
    });

    const sortedDateKeys = Object.keys(groupedData).sort((a, b) => 
      new Date(groupedData[a].date).getTime() - new Date(groupedData[b].date).getTime()
    );

    // SLIDE P.4: 最新回配信レポート
    const latestKey = sortedDateKeys[sortedDateKeys.length - 1];
    const latest = groupedData[latestKey];
    if (latest && latest.active) {
      const slide = pptx.addSlide();
      slide.addText("ライブ配信数値 - 最新回詳細報告", { x: 0.5, y: 0.3, w: 10, fontSize: 24, bold: true, color: COLORS.SLATE_900 });
      
      const kpis = [
        { label: "視聴者数", val: latest.active['Unique Viewers Count'] },
        { label: "最大同時接続者数", val: latest.active['Peak Concurrent Viewers Count'] },
        { label: "平均視聴時間", val: latest.active['Average Watched Minutes'], type: 'time' },
        { label: "いいね参加率", val: latest.active['Reaction Rate'], type: 'rate' },
        { label: "商品クリック率", val: latest.active['Product Ctr'], type: 'rate' },
        { label: "チャットユーザー率", val: latest.active['Chat Rate'], type: 'rate' },
      ];

      kpis.forEach((k, i) => {
        const x = 0.5 + i * 2.05;
        // @ts-ignore
        slide.addShape(pptx.ShapeType.RECT, { x, y: 0.8, w: 2, h: 1.8, line: { color: COLORS.BORDER_GREY, width: 0.5 }, fill: { color: COLORS.WHITE } });
        slide.addText(k.label, { x, y: 1.0, w: 2, fontSize: 9, align: "center", bold: true, color: COLORS.HEADER_GREY });
        slide.addText(formatValue(k.val, k.type === 'time' ? '時間' : (k.type === 'rate' ? '率' : '')), { x, y: 1.6, w: 2, fontSize: 20, align: "center", bold: true });
      });

      const insight4 = await generateInsight(kpis, "最新の配信データに基づく全体評価");
      // @ts-ignore
      slide.addShape(pptx.ShapeType.RECT, { x: 0.5, y: 5.5, w: 12.3, h: 1.5, fill: { color: COLORS.SLATE_100 } });
      slide.addText(`【分析考察】\n${insight4}`, { x: 0.7, y: 5.6, w: 11.9, fontSize: 11, color: COLORS.SLATE_700, margin: 10, bold: true });
    }

    // SLIDE P.5: 配信実績まとめ
    const slide5 = pptx.addSlide();
    slide5.addText("配信実績まとめ", { x: 0.5, y: 0.2, w: 10, fontSize: 20, bold: true });
    
    const tableHeaderStyle = { fill: COLORS.HEADER_GREY, color: COLORS.WHITE, bold: true, align: "center", fontSize: 8, border: { type: "solid", color: COLORS.WHITE, pt: 1 } };
    const bodyRows = sortedDateKeys.map(key => {
      const g = groupedData[key];
      const a = g.active || {};
      const r = g.replay || {};
      return [
        { text: key.substring(5).replace('-', '/'), options: { align: "center", fontSize: 9 } },
        { text: formatValue(a['Unique Viewers Count'], ""), options: { align: "center", bold: true, color: COLORS.FIREWORK_PINK } },
        { text: formatValue(r['Unique Viewers Count'], ""), options: { align: "center", bold: true, color: COLORS.FIREWORK_PINK } },
        { text: `${a['Total Chats'] || 0}\nL${a['Visitors With Chats'] || 0}人`, options: { align: "center", fontSize: 8 } },
        { text: `${formatValue(a['Reaction Rate'], "率")}\nL${a['Visitors With Reactions'] || 0}人`, options: { align: "center", fontSize: 8, color: COLORS.FIREWORK_BLUE } },
        { text: formatValue(a['Average Watched Minutes'], "時間"), options: { align: "center", fontSize: 8 } },
        { text: `${formatValue(a['Product Ctr'], "率")}\nL${a['Visitors With Product Clicks'] || 0}人`, options: { align: "center", fontSize: 8 } },
        { text: (Number(a['Total Product Clicks'] || 0) + Number(r['Total Product Clicks'] || 0)).toLocaleString(), options: { align: "center", bold: true, color: COLORS.FIREWORK_PINK } },
        { text: "-", options: { align: "center" } },
        { text: "-", options: { align: "center" } }
      ];
    });

    slide5.addTable([
      [
        { text: "日付", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "視聴者数", options: { ...tableHeaderStyle, colspan: 2 } },
        { text: "チャット数\nLユーザー数", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "いいね参加率\nLユーザー数", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "平均視聴分数", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "商品クリック率\nLユーザー数", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "クリック数\n(累計)", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "注文数", options: { ...tableHeaderStyle, colspan: 2 } }
      ],
      [
        { text: "ライブ", options: tableHeaderStyle }, { text: "アーカイブ", options: tableHeaderStyle },
        { text: "ライブ", options: tableHeaderStyle }, { text: "アーカイブ", options: tableHeaderStyle }
      ],
      ...bodyRows
    ] as any, { x: 0.5, y: 0.6, w: 12.3, border: { type: "solid", color: COLORS.BORDER_GREY, pt: 0.5 } });
  }

  // --- 2. SESSION DETAIL ANALYSIS (P.8 - P.12) ---
  if (sessionDataset) {
    const data = sessionDataset.data;
    const timeLabels = data.map(d => `${d['経過時間 (分)'] || d['Minute']}分`);
    const viewers = data.map(d => Number(d['同時視聴ユーザー数']) || 0);
    const likes = data.map(d => Number(d['いいね数']) || 0);
    const clicks = data.map(d => Number(d['商品クリック数']) || 0);
    const totalUU = liveUU;

    const plotArea = { x: 0.5, y: 1.6, w: 12.3, h: 4.0 };

    /**
     * SLIDE P.8: 視聴分数 - 振り返り
     */
    const slide8 = pptx.addSlide();
    slide8.addText("視聴分数 - 第3回振り返り", { x: 0.5, y: 0.3, w: 8, fontSize: 24, bold: true, color: COLORS.SLATE_900 });
    
    // 正確な平均視聴分数: (同時視聴者数の総和) / (ユニーク視聴者数)
    const sumConcurrent = viewers.reduce((a, b) => a + b, 0);
    const avgWatchTimeMinutes = sumConcurrent / totalUU;

    // KPI Box
    // @ts-ignore
    slide8.addShape(pptx.ShapeType.RECT, { x: 9.2, y: 0.5, w: 2.2, h: 1.0, line: { color: COLORS.SLATE_900, width: 1.5 } });
    slide8.addText("平均視聴時間", { x: 9.2, y: 0.6, w: 2.2, fontSize: 9, align: "center", bold: true });
    slide8.addText(formatValue(avgWatchTimeMinutes, "時間"), { x: 9.2, y: 0.9, w: 2.2, fontSize: 20, align: "center", bold: true, color: COLORS.SLATE_900 });

    slide8.addChart(pptx.ChartType.line, [{ name: "視聴者数", labels: timeLabels, values: viewers }], {
      x: plotArea.x, y: plotArea.y, w: plotArea.w, h: plotArea.h, 
      lineDataSymbol: 'none', chartColors: [COLORS.FIREWORK_BLUE], showLegend: false
    });

    findSignificantPoints(viewers, timeLabels, 5).forEach(p => addChartOvalMarker(pptx, slide8, p.idx, viewers.length, p.val, Math.max(...viewers, 1), plotArea));

    const insight8 = await generateInsight(viewers, "視聴者の維持・離脱要因の分析");
    // @ts-ignore
    slide8.addShape(pptx.ShapeType.RECT, { x: 0.5, y: 5.7, w: 12.3, h: 1.5, fill: { color: COLORS.LIGHT_PINK.replace('#','') } });
    slide8.addText(`【視聴分析】 ${insight8}`, { x: 0.7, y: 5.8, w: 11.9, fontSize: 11, bold: true, color: COLORS.SLATE_700, margin: 10 });

    /**
     * SLIDE P.10: いいね - 振り返り
     */
    const slide10 = pptx.addSlide();
    slide10.addText("いいね - 第3回振り返り", { x: 0.5, y: 0.3, w: 8, fontSize: 24, bold: true });
    
    const likeRate = Number(activeRow?.['Reaction Rate']) || 0;
    const likeUsers = Number(activeRow?.['Visitors With Reactions']) || 0;

    // KPI Boxes
    // @ts-ignore
    slide10.addShape(pptx.ShapeType.RECT, { x: 9.5, y: 0.5, w: 1.8, h: 1.0, line: { color: COLORS.SLATE_900, width: 1.5 } });
    slide10.addText("いいね参加率", { x: 9.5, y: 0.6, w: 1.8, fontSize: 9, align: "center", bold: true });
    slide10.addText(formatValue(likeRate, "率"), { x: 9.5, y: 0.9, w: 1.8, fontSize: 20, align: "center", bold: true, color: COLORS.FIREWORK_PINK });

    // @ts-ignore
    slide10.addShape(pptx.ShapeType.RECT, { x: 11.4, y: 0.5, w: 1.5, h: 1.0, line: { color: COLORS.SLATE_900, width: 1.5 } });
    slide10.addText("いいね参加人数", { x: 11.4, y: 0.6, w: 1.5, fontSize: 9, align: "center", bold: true });
    slide10.addText(`${likeUsers}人`, { x: 11.4, y: 0.9, w: 1.5, fontSize: 18, align: "center", bold: true });

    slide10.addChart(pptx.ChartType.line, [
      { name: "同時視聴数", labels: timeLabels, values: viewers },
      { name: "いいね数", labels: timeLabels, values: likes }
    ], { x: plotArea.x, y: plotArea.y, w: plotArea.w, h: plotArea.h, lineDataSymbol: 'none', chartColors: ["CCCCCC", COLORS.FIREWORK_PINK], showLegend: true });

    findSignificantPoints(likes, timeLabels, 5).forEach(p => addChartOvalMarker(pptx, slide10, p.idx, likes.length, p.val, Math.max(...likes, 1), plotArea));

    /**
     * SLIDE P.11: 商品クリック率 - 振り返り
     */
    const slide11 = pptx.addSlide();
    slide11.addText("商品クリック率 - 第3回振り返り", { x: 0.5, y: 0.3, w: 8, fontSize: 24, bold: true });

    const clickRate = Number(activeRow?.['Product Ctr']) || 0;
    const clickCount = Number(activeRow?.['Total Product Clicks']) || 0;

    // KPI Boxes
    // @ts-ignore
    slide11.addShape(pptx.ShapeType.RECT, { x: 9.5, y: 0.5, w: 1.8, h: 1.0, line: { color: COLORS.SLATE_900, width: 1.5 } });
    slide11.addText("商品クリック率", { x: 9.5, y: 0.6, w: 1.8, fontSize: 9, align: "center", bold: true });
    slide11.addText(formatValue(clickRate, "率"), { x: 9.5, y: 0.9, w: 1.8, fontSize: 20, align: "center", bold: true, color: COLORS.FIREWORK_ORANGE });

    // @ts-ignore
    slide11.addShape(pptx.ShapeType.RECT, { x: 11.4, y: 0.5, w: 1.5, h: 1.0, line: { color: COLORS.SLATE_900, width: 1.5 } });
    slide11.addText("商品クリック数", { x: 11.4, y: 0.6, w: 1.5, fontSize: 9, align: "center", bold: true });
    slide11.addText(`${clickCount}回`, { x: 11.4, y: 0.9, w: 1.5, fontSize: 18, align: "center", bold: true });

    slide11.addChart(pptx.ChartType.line, [
      { name: "同時視聴数", labels: timeLabels, values: viewers },
      { name: "商品クリック数", labels: timeLabels, values: clicks }
    ], { x: plotArea.x, y: plotArea.y, w: plotArea.w, h: plotArea.h, lineDataSymbol: 'none', chartColors: ["CCCCCC", COLORS.FIREWORK_ORANGE], showLegend: true });

    findSignificantPoints(clicks, timeLabels, 5).forEach(p => addChartOvalMarker(pptx, slide11, p.idx, clicks.length, p.val, Math.max(...clicks, 1), plotArea));

    /**
     * SLIDE P.12: 商品クリック率 - まとめ
     */
    const slide12 = pptx.addSlide();
    slide12.addText("商品クリック率 - 第3回まとめ", { x: 0.5, y: 0.3, w: 12, fontSize: 24, bold: true });

    const topClicks = findSignificantPoints(clicks, timeLabels, 5);
    const smallChartArea = { x: 0.5, y: 1.6, w: 8.5, h: 3.5 };

    slide12.addChart(pptx.ChartType.line, [{ name: "クリック数", labels: timeLabels, values: clicks }], {
      x: smallChartArea.x, y: smallChartArea.y, w: smallChartArea.w, h: smallChartArea.h, lineDataSymbol: 'none', chartColors: [COLORS.FIREWORK_ORANGE], showLegend: false
    });

    // サイドバー：クリック発生箇所の詳細
    // @ts-ignore
    slide12.addShape(pptx.ShapeType.RECT, { x: 9.2, y: 1.6, w: 3.8, h: 3.5, fill: { color: COLORS.LIGHT_YELLOW.replace('#','') }, line: { color: COLORS.BORDER_GREY, width: 0.5 } });
    topClicks.forEach((p, i) => {
      const y = 1.8 + i * 0.65;
      const num = String(i + 1);
      slide12.addText(`${num} 経過時間：${p.label} / クリック数：${p.val}`, { x: 9.4, y, w: 3.4, fontSize: 11, bold: true, color: COLORS.SLATE_900 });
      slide12.addText("・商品紹介セグメントでの具体的な興味・関心ピーク。", { x: 9.4, y: y + 0.3, w: 3.4, fontSize: 9, color: COLORS.SLATE_700 });
      
      // チャート上の番号付き◯と連動
      addChartOvalMarker(pptx, slide12, p.idx, clicks.length, p.val, Math.max(...clicks, 1), smallChartArea, num);
    });

    const insight12 = await generateInsight(topClicks, "高クリックが発生した箇所の構成分析と改善案");
    // @ts-ignore
    slide12.addShape(pptx.ShapeType.RECT, { x: 0.5, y: 5.5, w: 12.3, h: 1.8, fill: { color: COLORS.SLATE_100 } });
    slide12.addText(`【総括分析】\n${insight12}`, { x: 0.7, y: 5.6, w: 11.9, fontSize: 11, bold: true, color: COLORS.SLATE_700, margin: 10 });
  }

  const fileName = `Firework_Report_Complete_${new Date().toISOString().split('T')[0]}.pptx`;
  pptx.writeFile({ fileName });
};
