
export interface DataRow {
  [key: string]: any;
}

export interface MappingField {
  targetKey: string;
  label: string;
  sourceKey: string;
  isCustom?: boolean;
  confidence?: number; // 0 to 1
}

export interface Template {
  id: string;
  name: string;
  description: string;
  fields: { targetKey: string; label: string }[];
  mode: 'latest' | 'all' | 'session';
}

export interface Dataset {
  name: string;
  data: DataRow[];
  headers: string[];
  type: 'summary' | 'session' | 'unknown';
}

export const TEMPLATES: Template[] = [
  {
    id: 'p4-latest',
    name: '【P.4】最新回 詳細報告',
    description: 'ライブ配信数値とアーカイブ数値を抽出します。',
    mode: 'latest',
    fields: [
      { targetKey: 'date', label: '配信日' },
      { targetKey: 'title', label: '配信タイトル' },
      { targetKey: 'live_uu', label: '視聴者数 (ライブ)' },
      { targetKey: 'live_peak', label: '最大同時接続者数 (ライブ)' },
      { targetKey: 'live_avg_time', label: '平均視聴時間 (ライブ)' },
      { targetKey: 'live_like_rate', label: 'いいね率 (ライブ)' },
      { targetKey: 'live_click_rate', label: '商品クリック率 (ライブ)' },
      { targetKey: 'live_chat_rate', label: 'チャットユーザー率 (ライブ)' },
      { targetKey: 'arc_uu', label: '視聴総数 (アーカイブ)' },
      { targetKey: 'arc_active_rate', label: 'アクティブ率 (アーカイブ)' },
      { targetKey: 'arc_avg_time', label: '平均視聴時間 (アーカイブ)' },
      { targetKey: 'arc_clicks', label: '商品クリック数 (アーカイブ)' }
    ]
  },
  {
    id: 'p5-summary',
    name: '【P.5】配信実績まとめ',
    description: '全日程のデータをテーブル形式で出力します。',
    mode: 'all',
    fields: [
      { targetKey: 'date', label: '日程' },
      { targetKey: 'live_uu', label: '視聴者数 (ライブ)' },
      { targetKey: 'arc_uu', label: '視聴者数 (アーカイブ)' },
      { targetKey: 'chats_users', label: 'チャット数 / Lユーザー数' },
      { targetKey: 'likes_users', label: 'いいね参加率 / Lユーザー数' },
      { targetKey: 'avg_min', label: '平均視聴分数' },
      { targetKey: 'click_rate_live', label: '商品クリック率 (ライブ)' },
      { targetKey: 'click_total', label: '商品クリック数 (ライブ+アーカイブ)' }
    ]
  },
  {
    id: 'p8-session',
    name: '【P.8-12】セッション分析',
    description: '分単位の推移データを分析し、時系列グラフを出力します。',
    mode: 'session',
    fields: [
      { targetKey: 'minute', label: '経過時間 (分)' },
      { targetKey: 'concurrent', label: '同時視聴ユーザー数' },
      { targetKey: 'likes', label: 'いいね数' },
      { targetKey: 'chats', label: 'チャット数' },
      { targetKey: 'clicks', label: '商品クリック数' },
      { targetKey: 'cart', label: 'カート追加クリック数' }
    ]
  }
];

export const DEFAULT_TEMPLATE = TEMPLATES[0];
