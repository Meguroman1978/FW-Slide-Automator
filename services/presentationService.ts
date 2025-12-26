
import pptxgen from "pptxgenjs";
import { Dataset, DataRow, MappingField } from "../types";
import { generateInsight } from "./geminiService";

/**
 * 視覚的な定数定義
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
  SOFT_BLUE: "E1F5FE",
  ZONE_GREEN: "C8E6C9",
  ZONE_PINK: "FFCDD2",
  ANNOTATION_GREEN: "4CAF50",
  ANNOTATION_BLUE: "2196F3"
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
 * チャート上のピーク地点を特定
 */
const findSignificantPoints = (values: number[], labels: string[], count: number = 5) => {
  return values
    .map((val, idx) => ({ val, label: labels[idx], idx }))
    .sort((a, b) => b.val - a.val)
    .slice(0, count)
    .sort((a, b) => a.idx - b.idx);
};

/**
 * タイムラインマーカー（青い点）を追加
 */
const addTimelineMarkers = (slide: pptxgen.Slide, plotArea: { x: number, y: number, w: number, h: number }, markers: string[]) => {
  const markerY = plotArea.y - 0.3;
  markers.forEach((label, idx) => {
    const x = plotArea.x + (idx / (markers.length - 1)) * plotArea.w;
    
    // 青い点
    slide.addShape("ellipse", {
      x: x - 0.08, y: markerY - 0.08, w: 0.16, h: 0.16,
      fill: { color: COLORS.ANNOTATION_BLUE }
    });
    
    // ラベル
    slide.addText(label, {
      x: x - 0.5, y: markerY - 0.5, w: 1, h: 0.3,
      fontSize: 8, align: "center", color: COLORS.ANNOTATION_BLUE, bold: true
    });
  });
  
  // 青い線
  slide.addShape("line", {
    x: plotArea.x, y: markerY,
    w: plotArea.w, h: 0,
    line: { color: COLORS.ANNOTATION_BLUE, width: 2 }
  });
};

/**
 * 背景ゾーン（グリーン/ピンク）を追加
 */
const addBackgroundZones = (slide: pptxgen.Slide, plotArea: { x: number, y: number, w: number, h: number }, zones: { start: number, end: number, color: string }[], total: number) => {
  zones.forEach(zone => {
    const startX = plotArea.x + (zone.start / total) * plotArea.w;
    const width = ((zone.end - zone.start) / total) * plotArea.w;
    
    slide.addShape("rect", {
      x: startX, y: plotArea.y, w: width, h: plotArea.h,
      fill: { color: zone.color, transparency: 70 },
      line: { type: "none" }
    });
  });
};

/**
 * 注釈ボックスを追加（緑/青枠付き白背景）
 */
const addAnnotationBox = (slide: pptxgen.Slide, x: number, y: number, text: string, color: string = COLORS.ANNOTATION_GREEN) => {
  const boxWidth = 1.8;
  const boxHeight = 0.8;
  
  // 白背景の矩形
  slide.addShape("rect", {
    x, y, w: boxWidth, h: boxHeight,
    fill: { color: COLORS.WHITE },
    line: { color, width: 2 }
  });
  
  // テキスト
  slide.addText(text, {
    x, y: y + 0.1, w: boxWidth, h: boxHeight - 0.2,
    fontSize: 9, align: "center", valign: "middle", bold: true, color: COLORS.SLATE_900
  });
};

/**
 * チャート上に円形マーカーを追加
 */
const addChartCircleMarker = (
  slide: pptxgen.Slide,
  index: number,
  total: number,
  value: number,
  maxValue: number,
  chartArea: { x: number, y: number, w: number, h: number },
  number?: string,
  color: string = "FF0000"
) => {
  const x = chartArea.x + (index / (total - 1)) * chartArea.w;
  const y = chartArea.y + chartArea.h - (value / (maxValue || 1)) * chartArea.h;

  // 円形マーカー
  slide.addShape("ellipse", {
    x: x - 0.2, y: y - 0.2, w: 0.4, h: 0.4,
    line: { color, width: 3 },
    fill: { type: "solid", color: COLORS.WHITE, transparency: 50 }
  });

  // 番号ラベル
  if (number) {
    slide.addText(number, {
      x: x - 0.15, y: y - 0.5, w: 0.3, h: 0.25,
      fontSize: 11, bold: true, align: "center", color
    });
  }
};

export const generateReportPPTX = async (datasets: Dataset[], allMappings: { [templateId: string]: MappingField[] }) => {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";

  const summaryDataset = datasets.find(ds => ds.type === 'summary');
  const sessionDataset = datasets.find(ds => ds.type === 'session');

  // ユニーク視聴者数
  const activeRow = summaryDataset?.data.find(d => String(d['Live Stream Status']).toLowerCase() === 'active');
  const liveUU = Number(activeRow?.['Unique Viewers Count']) || 1;

  // --- 1. DIGITAL SHOWROOM SUMMARY (P.6 - P.7) ---
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

    // SLIDE P.6: 最新回配信レポート（{{YYYY/MM/DD}}配信全体数値形式）
    const latestKey = sortedDateKeys[sortedDateKeys.length - 1];
    const latest = groupedData[latestKey];
    if (latest && latest.active) {
      const slide = pptx.addSlide();
      
      // タイトル（動的に日付を挿入）
      const dateObj = new Date(latest.date);
      const formattedDate = `${dateObj.getFullYear()}/${dateObj.getMonth() + 1}/${dateObj.getDate()}`;
      slide.addText(`${formattedDate}配信全体数値`, { 
        x: 0.5, y: 0.4, w: 6, h: 0.5, 
        fontSize: 18, bold: true, color: COLORS.SLATE_900 
      });

      // タイトル下のサブテキスト
      slide.addText("タイトル", { 
        x: 0.7, y: 1.0, w: 1.2, h: 0.4, 
        fontSize: 11, bold: true, color: COLORS.SLATE_900,
        valign: "middle"
      });
      
      const eventTitle = latest.active['Title'] || latest.active['Event Title'] || "配信イベント";
      slide.addShape("rect", { 
        x: 2.0, y: 1.0, w: 5.0, h: 0.4, 
        fill: { color: COLORS.WHITE }, 
        line: { color: COLORS.BORDER_GREY, width: 0.5 } 
      });
      slide.addText(eventTitle, { 
        x: 2.1, y: 1.0, w: 4.8, h: 0.4, 
        fontSize: 10, color: COLORS.SLATE_900,
        valign: "middle"
      });

      // ライブ配信数値エリア
      slide.addText("ライブ配信数値", { 
        x: 0.5, y: 1.7, w: 7, h: 0.4, 
        fontSize: 13, bold: true, color: COLORS.SLATE_900 
      });
      
      const liveMetrics = [
        { label: "視聴者数", val: latest.active['Unique Viewers Count'] },
        { label: "最大同時接続者数", val: latest.active['Peak Concurrent Viewers Count'] },
        { label: "平均視聴時間", val: latest.active['Average Watched Minutes'], type: 'time' },
        { label: "いいね率", val: latest.active['Reaction Rate'], type: 'rate' },
        { label: "商品クリック率", val: latest.active['Product Ctr'], type: 'rate' },
        { label: "チャットユーザー率", val: latest.active['Chat Rate'], type: 'rate' }
      ];

      // ライブKPIボックス（3列 x 2行）
      liveMetrics.forEach((metric, i) => {
        const col = i % 3;
        const row = Math.floor(i / 3);
        const x = 0.5 + col * 2.3;
        const y = 2.2 + row * 1.1;
        
        slide.addShape("rect", { 
          x, y, w: 2.2, h: 1.0, 
          fill: { color: COLORS.WHITE }, 
          line: { color: COLORS.BORDER_GREY, width: 1 } 
        });
        
        slide.addText(metric.label, { 
          x, y: y + 0.1, w: 2.2, h: 0.35, 
          fontSize: 9, align: "center", bold: true, color: COLORS.HEADER_GREY 
        });
        
        slide.addText(formatValue(metric.val, metric.type === 'time' ? '時間' : (metric.type === 'rate' ? '率' : '')), { 
          x, y: y + 0.5, w: 2.2, h: 0.4, 
          fontSize: 16, align: "center", bold: true, color: COLORS.SLATE_900 
        });
      });

      // アーカイブ数値エリア
      if (latest.replay) {
        slide.addText("アーカイブ数値*{{(MM/DD)}}時点", { 
          x: 0.5, y: 4.5, w: 7, h: 0.4, 
          fontSize: 13, bold: true, color: COLORS.SLATE_900 
        });
        
        const archiveMetrics = [
          { label: "視聴回数", val: latest.replay['Unique Viewers Count'] },
          { label: "アクティブ率", val: latest.replay['Active Rate'] || 0, type: 'rate' },
          { label: "最大同時接続者数", val: latest.replay['Peak Concurrent Viewers Count'] || 0 },
          { label: "平均視聴時間", val: latest.replay['Average Watched Minutes'] || 0, type: 'time' },
          { label: "商品クリック率", val: latest.replay['Product Ctr'] || 0, type: 'rate' },
          { label: "商品クリック数", val: latest.replay['Total Product Clicks'] || 0 }
        ];

        // アーカイブKPIボックス（3列 x 2行）
        archiveMetrics.forEach((metric, i) => {
          const col = i % 3;
          const row = Math.floor(i / 3);
          const x = 0.5 + col * 2.3;
          const y = 5.0 + row * 1.1;
          
          slide.addShape("rect", { 
            x, y, w: 2.2, h: 1.0, 
            fill: { color: COLORS.WHITE }, 
            line: { color: COLORS.BORDER_GREY, width: 1 } 
          });
          
          slide.addText(metric.label, { 
            x, y: y + 0.1, w: 2.2, h: 0.35, 
            fontSize: 9, align: "center", bold: true, color: COLORS.HEADER_GREY 
          });
          
          slide.addText(formatValue(metric.val, metric.type === 'time' ? '時間' : (metric.type === 'rate' ? '率' : '')), { 
            x, y: y + 0.5, w: 2.2, h: 0.4, 
            fontSize: 16, align: "center", bold: true, color: COLORS.SLATE_900 
          });
        });
      }

      // 右側の画像エリア（ブラックフライデー風デザイン）
      slide.addShape("rect", { 
        x: 7.5, y: 0.5, w: 5.5, h: 6.5, 
        fill: { color: "2C3E50" }
      });
      
      slide.addText("イベント配信", { 
        x: 7.7, y: 1.0, w: 5.1, h: 0.6, 
        fontSize: 20, bold: true, color: "FFA500", align: "center" 
      });
      
      slide.addText(`${formattedDate}\n19:30〜20:30`, { 
        x: 7.7, y: 2.0, w: 5.1, h: 1.0, 
        fontSize: 14, bold: true, color: COLORS.WHITE, align: "center", valign: "middle" 
      });
      
      // Fireworkロゴエリア
      slide.addText("∞ Firework", { 
        x: 7.7, y: 6.5, w: 5.1, h: 0.4, 
        fontSize: 11, color: COLORS.WHITE, align: "right" 
      });
    }

    // SLIDE P.7: 過去配信との比較
    const slide7 = pptx.addSlide();
    
    // タイトル
    slide7.addText("過去配信との比較", { 
      x: 0.5, y: 0.3, w: 6, h: 0.5, 
      fontSize: 18, bold: true, color: COLORS.SLATE_900 
    });
    
    // Insightを動的に生成
    const comparisonData = sortedDateKeys.map(key => {
      const g = groupedData[key];
      return {
        date: key,
        liveViewers: g.active?.['Unique Viewers Count'] || 0,
        archiveViewers: g.replay?.['Unique Viewers Count'] || 0,
        avgTime: g.active?.['Average Watched Minutes'] || 0,
        clickRate: g.active?.['Product Ctr'] || 0
      };
    });
    const insight7 = await generateInsight(comparisonData, "過去配信データの比較分析と傾向把握");
    
    slide7.addText(insight7, { 
      x: 0.5, y: 1.0, w: 12, h: 0.8, 
      fontSize: 11, bold: false, color: COLORS.SLATE_700 
    });
    
    // テーブル作成
    const tableHeaderStyle = { 
      fill: "808080", 
      color: COLORS.WHITE, 
      bold: true, 
      align: "center", 
      fontSize: 9,
      valign: "middle"
    };
    
    const bodyRows = sortedDateKeys.map((key, index) => {
      const g = groupedData[key];
      const a = g.active || {};
      const r = g.replay || {};
      const dateObj = new Date(g.date);
      const formattedDate = `${dateObj.getFullYear()}\n/${dateObj.getMonth() + 1}\n/${dateObj.getDate()}\n${index + 1}`;
      
      return [
        { text: formattedDate, options: { align: "center", fontSize: 8, bold: true, valign: "middle" } },
        { text: (a['Unique Viewers Count'] || 0).toLocaleString() + "人", options: { align: "center", fontSize: 11, bold: true, color: index === sortedDateKeys.length - 1 ? "FF1744" : "000000" } },
        { text: (r['Unique Viewers Count'] || 0).toLocaleString() + "人", options: { align: "center", fontSize: 11, bold: false, color: index === sortedDateKeys.length - 1 ? "FF1744" : "000000" } },
        { text: `${a['Total Chats'] || 0}\nL${a['Visitors With Chats'] || 0}人`, options: { align: "center", fontSize: 9, color: index === sortedDateKeys.length - 1 ? "FF1744" : "000000" } },
        { text: `${formatValue(a['Reaction Rate'], "率")}\nL${a['Visitors With Reactions'] || 0}人`, options: { align: "center", fontSize: 9, color: index === sortedDateKeys.length - 1 ? "FF1744" : "0066FF" } },
        { text: formatValue(a['Average Watched Minutes'], "時間"), options: { align: "center", fontSize: 10, color: index === sortedDateKeys.length - 1 ? "2196F3" : "000000" } },
        { text: `${formatValue(a['Product Ctr'], "率")}\nL${a['Visitors With Product Clicks'] || 0}人`, options: { align: "center", fontSize: 9, color: index === sortedDateKeys.length - 1 ? "FF1744" : "000000" } },
        { text: `${(Number(a['Total Product Clicks'] || 0) + Number(r['Total Product Clicks'] || 0)).toLocaleString()}\n(ライブは${a['Total Product Clicks'] || 0})`, options: { align: "center", fontSize: 9, bold: true, color: index === sortedDateKeys.length - 1 ? "FF1744" : "000000" } },
        { text: (a['Checkout Count'] || 0).toString(), options: { align: "center", fontSize: 11, bold: true, color: index === sortedDateKeys.length - 1 ? "E91E63" : "000000" } },
        { text: (r['Checkout Count'] || 0).toString(), options: { align: "center", fontSize: 11, bold: false } }
      ];
    });

    slide7.addTable([
      [
        { text: "", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "視聴者数", options: { ...tableHeaderStyle, colspan: 2 } },
        { text: "チャット数\nLユーザー数", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "いいね参加率\nLユーザー数", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "平均視聴分数", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "商品クリック率\nLユーザー数\n(ライブ)", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "商品クリック数\n(ライブ+\nアーカイブ)", options: { ...tableHeaderStyle, rowspan: 2 } },
        { text: "チェックアウト数", options: { ...tableHeaderStyle, colspan: 2 } }
      ],
      [
        { text: "ライブ", options: tableHeaderStyle }, 
        { text: "アーカイブ", options: tableHeaderStyle },
        { text: "ライブ", options: tableHeaderStyle }, 
        { text: "アーカイブ", options: tableHeaderStyle }
      ],
      ...bodyRows
    ] as any, { 
      x: 0.5, y: 2.0, w: 12.5, 
      border: { type: "solid", color: COLORS.BORDER_GREY, pt: 0.5 },
      rowH: [0.5, 0.5, ...Array(bodyRows.length).fill(0.7)]
    });
    
    // テーブル下の注釈
    slide7.addText("※12/01時点の数値（7/16,9/12のアーカイブは現在時点の数値)", { 
      x: 6.0, y: 2.0 + 0.5 + 0.5 + bodyRows.length * 0.7 + 0.1, w: 6.5, h: 0.3, 
      fontSize: 8, color: COLORS.SLATE_700, align: "right" 
    });
  }

  // --- 2. SESSION DETAIL ANALYSIS (P.8 - P.12) ---
  if (sessionDataset) {
    const data = sessionDataset.data;
    const timeLabels = data.map(d => `${d['経過時間 (分)'] || d['Minute']}`);
    const viewers = data.map(d => Number(d['同時視聴ユーザー数']) || 0);
    const likes = data.map(d => Number(d['いいね数']) || 0);
    const chats = data.map(d => Number(d['チャット数']) || 0);
    const clicks = data.map(d => Number(d['商品クリック数']) || 0);
    const totalUU = liveUU;

    const plotArea = { x: 1.5, y: 2.3, w: 11, h: 3.8 };
    
    // タイムラインマーカー
    const timeMarkers = ["オープニング", "1500 New", "1000 New", "2000 New / 3000 New", "240 New / 500 New", "エンディング"];

    /**
     * SLIDE P.8: 視聴分数 - 振り返り
     */
    const slide8 = pptx.addSlide();
    slide8.addText("視聴分数 - 第3回振り返り", { x: 0.5, y: 0.4, w: 8, fontSize: 20, bold: true, color: COLORS.SLATE_900 });
    
    // 正確な平均視聴分数
    const sumConcurrent = viewers.reduce((a, b) => a + b, 0);
    const avgWatchTimeMinutes = sumConcurrent / totalUU;
    const peakViewers = Math.max(...viewers);

    // KPI Boxes
    slide8.addShape("rect", { x: 10.5, y: 0.5, w: 1.3, h: 0.8, fill: { color: COLORS.WHITE }, line: { color: COLORS.SLATE_900, width: 1 } });
    slide8.addText("平均視聴時間", { x: 10.5, y: 0.55, w: 1.3, h: 0.25, fontSize: 8, align: "center", bold: true });
    slide8.addText(formatValue(avgWatchTimeMinutes, "時間"), { x: 10.5, y: 0.85, w: 1.3, h: 0.35, fontSize: 16, align: "center", bold: true });

    slide8.addShape("rect", { x: 11.9, y: 0.5, w: 1.3, h: 0.8, fill: { color: COLORS.WHITE }, line: { color: COLORS.SLATE_900, width: 1 } });
    slide8.addText("最大同時接続者数", { x: 11.9, y: 0.55, w: 1.3, h: 0.25, fontSize: 8, align: "center", bold: true });
    slide8.addText(`${peakViewers}人`, { x: 11.9, y: 0.85, w: 1.3, h: 0.35, fontSize: 16, align: "center", bold: true });

    // タイムラインマーカー
    addTimelineMarkers(slide8, plotArea, timeMarkers);

    // 背景ゾーン
    const zones8 = [
      { start: 0, end: 10, color: COLORS.ZONE_PINK },
      { start: 20, end: 35, color: COLORS.ZONE_GREEN },
      { start: 45, end: 55, color: COLORS.ZONE_GREEN }
    ];
    addBackgroundZones(slide8, plotArea, zones8, data.length);

    // グラフ
    slide8.addChart(pptx.ChartType.line, [
      { name: "同時視聴ユーザー数", labels: timeLabels, values: viewers }
    ], {
      x: plotArea.x, y: plotArea.y, w: plotArea.w, h: plotArea.h,
      lineDataSymbol: 'none', chartColors: [COLORS.FIREWORK_BLUE], showLegend: false,
      showTitle: false,
      catAxisLabelFontSize: 8,
      valAxisLabelFontSize: 8
    });

    // ピークポイントをマーク
    const topViewerPoints = findSignificantPoints(viewers, timeLabels, 3);
    topViewerPoints.forEach((p, i) => {
      addChartCircleMarker(slide8, p.idx, viewers.length, p.val, Math.max(...viewers), plotArea, undefined, COLORS.FIREWORK_ORANGE);
    });

    // 注釈ボックス
    addAnnotationBox(slide8, 3.5, 1.5, "最大脱退ポイント\n(87人)", COLORS.ANNOTATION_BLUE);
    addAnnotationBox(slide8, 6.5, 3.0, "配信コンタイム", "FFA726");

    // 分析コメント
    const insight8 = await generateInsight({ viewers, avgTime: avgWatchTimeMinutes }, "視聴者の維持・離脱要因の分析");
    slide8.addShape("rect", { x: 0.5, y: 6.3, w: 12.5, h: 0.9, fill: { color: "FFE5E5" }, line: { type: "none" } });
    slide8.addText(`【視聴分析】${insight8}`, { x: 0.7, y: 6.4, w: 12.1, h: 0.7, fontSize: 9, color: COLORS.SLATE_900, valign: "middle" });

    /**
     * SLIDE P.10: いいね - 振り返り
     */
    const slide10 = pptx.addSlide();
    slide10.addText("いいね - 第3回振り返り", { x: 0.5, y: 0.4, w: 8, fontSize: 20, bold: true });
    
    const likeRate = Number(activeRow?.['Reaction Rate']) || 0;
    const likeUsers = Number(activeRow?.['Visitors With Reactions']) || 0;

    // KPI Boxes
    slide10.addShape("rect", { x: 10.5, y: 0.5, w: 1.3, h: 0.8, fill: { color: COLORS.WHITE }, line: { color: COLORS.SLATE_900, width: 1 } });
    slide10.addText("いいね参加率", { x: 10.5, y: 0.55, w: 1.3, h: 0.25, fontSize: 8, align: "center", bold: true });
    slide10.addText(formatValue(likeRate, "率"), { x: 10.5, y: 0.85, w: 1.3, h: 0.35, fontSize: 16, align: "center", bold: true, color: COLORS.FIREWORK_PINK });

    slide10.addShape("rect", { x: 11.9, y: 0.5, w: 1.3, h: 0.8, fill: { color: COLORS.WHITE }, line: { color: COLORS.SLATE_900, width: 1 } });
    slide10.addText("いいね参加人数", { x: 11.9, y: 0.55, w: 1.3, h: 0.25, fontSize: 8, align: "center", bold: true });
    slide10.addText(`${likeUsers}人`, { x: 11.9, y: 0.85, w: 1.3, h: 0.35, fontSize: 16, align: "center", bold: true });

    // タイムラインマーカー
    addTimelineMarkers(slide10, plotArea, timeMarkers);

    // 背景ゾーン
    addBackgroundZones(slide10, plotArea, zones8, data.length);

    // グラフ（2線）
    slide10.addChart(pptx.ChartType.line, [
      { name: "同時視聴ユーザー数", labels: timeLabels, values: viewers },
      { name: "いいね数", labels: timeLabels, values: likes }
    ], {
      x: plotArea.x, y: plotArea.y, w: plotArea.w, h: plotArea.h,
      lineDataSymbol: 'none',
      chartColors: ["CCCCCC", COLORS.FIREWORK_PINK],
      showLegend: true,
      legendPos: "b",
      showTitle: false,
      catAxisLabelFontSize: 8,
      valAxisLabelFontSize: 8
    });

    // ピークポイントをマーク
    const topLikePoints = findSignificantPoints(likes, timeLabels, 3);
    topLikePoints.forEach(p => {
      addChartCircleMarker(slide10, p.idx, likes.length, p.val, Math.max(...likes), plotArea, undefined, "FFD700");
    });

    // いいね分析コメント
    const insight10 = await generateInsight({ likes, likeRate, likeUsers, topPoints: topLikePoints }, "いいね数の推移と視聴者エンゲージメント分析");
    slide10.addShape("rect", { x: 0.5, y: 6.3, w: 12.5, h: 0.9, fill: { color: "E8F5E9" }, line: { type: "none" } });
    slide10.addText(`【いいね分析】${insight10}`, { x: 0.7, y: 6.4, w: 12.1, h: 0.7, fontSize: 9, color: COLORS.SLATE_900, valign: "middle" });

    /**
     * SLIDE P.11: 商品クリック率 - 振り返り
     */
    const slide11 = pptx.addSlide();
    slide11.addText("商品クリック率 - 第3回振り返り", { x: 0.5, y: 0.4, w: 8, fontSize: 20, bold: true });

    const clickRate = Number(activeRow?.['Product Ctr']) || 0;
    const clickCount = Number(activeRow?.['Total Product Clicks']) || 0;

    // KPI Boxes
    slide11.addShape("rect", { x: 10.5, y: 0.5, w: 1.3, h: 0.8, fill: { color: COLORS.WHITE }, line: { color: COLORS.SLATE_900, width: 1 } });
    slide11.addText("商品クリック率", { x: 10.5, y: 0.55, w: 1.3, h: 0.25, fontSize: 8, align: "center", bold: true });
    slide11.addText(formatValue(clickRate, "率"), { x: 10.5, y: 0.85, w: 1.3, h: 0.35, fontSize: 16, align: "center", bold: true, color: COLORS.FIREWORK_ORANGE });

    slide11.addShape("rect", { x: 11.9, y: 0.5, w: 1.3, h: 0.8, fill: { color: COLORS.WHITE }, line: { color: COLORS.SLATE_900, width: 1 } });
    slide11.addText("商品クリック数", { x: 11.9, y: 0.55, w: 1.3, h: 0.25, fontSize: 8, align: "center", bold: true });
    slide11.addText(`${clickCount}回`, { x: 11.9, y: 0.85, w: 1.3, h: 0.35, fontSize: 16, align: "center", bold: true });

    // タイムラインマーカー
    addTimelineMarkers(slide11, plotArea, timeMarkers);

    // 背景ゾーン
    addBackgroundZones(slide11, plotArea, zones8, data.length);

    // グラフ（2線）
    slide11.addChart(pptx.ChartType.line, [
      { name: "同時視聴ユーザー数", labels: timeLabels, values: viewers },
      { name: "商品クリック数", labels: timeLabels, values: clicks }
    ], {
      x: plotArea.x, y: plotArea.y, w: plotArea.w, h: plotArea.h,
      lineDataSymbol: 'none',
      chartColors: ["CCCCCC", COLORS.FIREWORK_ORANGE],
      showLegend: true,
      legendPos: "b",
      showTitle: false,
      catAxisLabelFontSize: 8,
      valAxisLabelFontSize: 8
    });

    // ピークポイントをマーク
    const topClickPoints = findSignificantPoints(clicks, timeLabels, 3);
    topClickPoints.forEach(p => {
      addChartCircleMarker(slide11, p.idx, clicks.length, p.val, Math.max(...clicks), plotArea, undefined, COLORS.FIREWORK_ORANGE);
    });

    // 商品クリック分析コメント
    const insight11 = await generateInsight({ clicks, clickRate, clickCount, topPoints: topClickPoints }, "商品クリック数の推移と購買意欲の分析");
    slide11.addShape("rect", { x: 0.5, y: 6.3, w: 12.5, h: 0.9, fill: { color: "FFF3E0" }, line: { type: "none" } });
    slide11.addText(`【クリック分析】${insight11}`, { x: 0.7, y: 6.4, w: 12.1, h: 0.7, fontSize: 9, color: COLORS.SLATE_900, valign: "middle" });

    /**
     * SLIDE P.12: 商品クリック率 - まとめ
     */
    const slide12 = pptx.addSlide();
    slide12.addText("商品クリック率 - 第3回まとめ", { x: 0.5, y: 0.4, w: 12, fontSize: 20, bold: true });

    const topClicks = findSignificantPoints(clicks, timeLabels, 5);
    const smallChartArea = { x: 1.5, y: 2.3, w: 8, h: 3.5 };

    // タイムラインマーカー
    addTimelineMarkers(slide12, smallChartArea, timeMarkers);

    // 背景ゾーン
    addBackgroundZones(slide12, smallChartArea, zones8, data.length);

    // グラフ
    slide12.addChart(pptx.ChartType.line, [
      { name: "商品クリック数", labels: timeLabels, values: clicks }
    ], {
      x: smallChartArea.x, y: smallChartArea.y, w: smallChartArea.w, h: smallChartArea.h,
      lineDataSymbol: 'none',
      chartColors: [COLORS.FIREWORK_ORANGE],
      showLegend: false,
      showTitle: false,
      catAxisLabelFontSize: 8,
      valAxisLabelFontSize: 8
    });

    // サイドバー：クリック発生箇所の詳細
    slide12.addShape("rect", { x: 9.7, y: 2.3, w: 3.3, h: 3.5, fill: { color: COLORS.LIGHT_YELLOW }, line: { color: COLORS.BORDER_GREY, width: 0.5 } });
    
    topClicks.forEach((p, i) => {
      const y = 2.5 + i * 0.65;
      const num = String(i + 1);
      
      // 番号付きマーカー
      addChartCircleMarker(slide12, p.idx, clicks.length, p.val, Math.max(...clicks), smallChartArea, num, COLORS.FIREWORK_ORANGE);
      
      // サイドバーの説明
      slide12.addText(`${num}  経過時間：${p.label}分 / クリック数：${p.val}`, {
        x: 9.8, y, w: 3.1, h: 0.25,
        fontSize: 9, bold: true, color: COLORS.SLATE_900
      });
      slide12.addText("・商品紹介セグメントでの具体的な興味・関心ピーク。", {
        x: 9.8, y: y + 0.27, w: 3.1, h: 0.3,
        fontSize: 8, color: COLORS.SLATE_700
      });
    });

    // 総括分析
    const insight12 = await generateInsight(topClicks, "高クリックが発生した箇所の構成分析と改善案");
    slide12.addShape("rect", { x: 0.5, y: 6.1, w: 12.5, h: 1.1, fill: { color: COLORS.SLATE_100 }, line: { type: "none" } });
    slide12.addText(`【総括分析】\n${insight12}`, { x: 0.7, y: 6.2, w: 12.1, h: 0.9, fontSize: 9, bold: true, color: COLORS.SLATE_700 });
  }

  const fileName = `Firework_Report_Complete_${new Date().toISOString().split('T')[0]}.pptx`;
  pptx.writeFile({ fileName });
};
