/**
 * ============================================================
 *  Config.js - SFA 全体設定
 * ============================================================
 *  全モジュールが参照する一元的な設定オブジェクト。
 *  デプロイ前にスクリプトプロパティ「Gemini_API」と ENV セクションを設定してください。
 * ============================================================
 */

const SFA_CONFIG = {

  // ────────────────────────────
  //  環境変数 (★ ここを自分の値に書き換え)
  // ────────────────────────────
  ENV: {
    FOLDER_ID: '1mDptHj7xmXw89pwBQI71EXU2UnKBRL5l',   // Google Drive 保存先フォルダID
    API_KEY:   PropertiesService.getScriptProperties().getProperty('Gemini_API') || '',   // スクリプトプロパティ「Gemini_API」から取得

    // Google Custom Search API (機能2で使用 / 任意)
    // ※ 未設定の場合は Gemini の知識のみで回答します
    CUSTOM_SEARCH_API_KEY: '',     // Google Cloud の API キー
    CUSTOM_SEARCH_CX:     '',      // Custom Search Engine の CX ID

    // Slack 通知 (機能3で使用 / 任意)
    SLACK_WEBHOOK_URL: '',          // Slack Incoming Webhook URL

    // Chatwork 通知 (機能3で使用 / 任意)
    CHATWORK_API_TOKEN: '',
    CHATWORK_ROOM_ID:   '',

    // Teams 通知 (機能3で使用 / 任意)
    TEAMS_WEBHOOK_URL: '',          // Teams Incoming Webhook URL
  },

  // ────────────────────────────
  //  Microsoft 365 連携設定
  // ────────────────────────────
  MICROSOFT: {
    CLIENT_ID:     PropertiesService.getScriptProperties().getProperty('MS_CLIENT_ID') || '',
    CLIENT_SECRET: PropertiesService.getScriptProperties().getProperty('MS_CLIENT_SECRET') || '',
    TENANT_ID:     PropertiesService.getScriptProperties().getProperty('MS_TENANT_ID') || 'common',
    SCOPES: [
      'https://graph.microsoft.com/Mail.Read',
      'https://graph.microsoft.com/Mail.ReadWrite',
      'https://graph.microsoft.com/User.Read',
    ],
    EMAIL_PROVIDER: 'gmail',   // 'gmail' | 'outlook' | 'both'
    SYNC_BATCH_SIZE: 50,
    SYNC_DELAY_MS: 500,
  },

  // ────────────────────────────
  //  Gemini モデル設定
  // ────────────────────────────
  GEMINI: {
    MODEL:        'gemini-3-flash-preview',
    API_BASE:     'https://generativelanguage.googleapis.com/v1beta/models/',
    TEMPERATURE:  0.3,     // SFA分析はやや創造的に
    MAX_RETRIES:  3,
    RETRY_DELAY:  1500,    // ms
  },

  // ────────────────────────────
  //  スプレッドシート シート名
  // ────────────────────────────
  SHEETS: {
    MAIN:               'シート1',                // メイン顧客リスト
    SIMILAR_COMPANIES:  '類似企業リスト',          // 機能5 出力先
    DORMANT_DRAFTS:     '休眠顧客メール下書き',    // 機能4 出力先
    NOTIFICATION_LOG:   '通知ログ',                // 機能3 アラートログ
  },

  // ────────────────────────────
  //  各シートのヘッダー定義
  // ────────────────────────────
  HEADERS: {
    MAIN: [
      '登録日',           // A
      '会社名',           // B
      '氏名',             // C
      '役職',             // D
      'Email',            // E
      '電話番号',          // F
      '住所',             // G
      'WebサイトURL',     // H
      '最終接触日',        // I
      '担当者名',          // J
      '名刺画像',          // K
      'X検索URL',         // L
      'Facebook検索URL',  // M
      'Instagram検索URL', // N
      'YouTube検索URL',   // O
      'TikTok検索URL',    // P
      '企業サイトURL',     // Q
      '推定業種',          // R
      '業界トレンド・ニュース', // S
      '想定課題',          // T
      '類似企業リスト',    // U
      '重複アラート',      // V
      '備考',              // W
    ],
    SIMILAR_COMPANIES: [
      '基準企業名',
      '類似企業名',
      '業種',
      '類似理由',
      'ターゲット優先度',
      '推定URL',
      '生成日',
    ],
    DORMANT_DRAFTS: [
      '生成日',
      '会社名',
      '氏名',
      'Email',
      '最終接触日',
      '経過日数',
      '業界最新ニュース',
      'メール件名',
      'メール本文',
      'ステータス',
    ],
    NOTIFICATION_LOG: [
      '日時',
      '種別',
      '対象会社名',
      '対象氏名',
      'メッセージ',
      '通知先',
    ],
  },

  // ────────────────────────────
  //  メイン顧客シートのカラムインデックス (0-based)
  // ────────────────────────────
  COL: {
    REGISTERED_DATE:  0,   // A: 登録日
    COMPANY_NAME:     1,   // B: 会社名
    FULL_NAME:        2,   // C: 氏名
    TITLE:            3,   // D: 役職
    EMAIL:            4,   // E: Email
    PHONE:            5,   // F: 電話番号
    ADDRESS:          6,   // G: 住所
    WEBSITE:          7,   // H: WebサイトURL
    LAST_CONTACT:     8,   // I: 最終接触日
    STAFF_NAME:       9,   // J: 担当者名
    IMAGE_URL:       10,   // K: 名刺画像
    X_URL:           11,   // L
    FACEBOOK_URL:    12,   // M
    INSTAGRAM_URL:   13,   // N
    YOUTUBE_URL:     14,   // O
    TIKTOK_URL:      15,   // P
    COMPANY_SITE:    16,   // Q
    INDUSTRY:        17,   // R
    TRENDS:          18,   // S
    CHALLENGES:      19,   // T
    SIMILAR:         20,   // U
    DUP_ALERT:       21,   // V
    NOTES:           22,   // W: 備考
  },

  // ────────────────────────────
  //  業務パラメータ
  // ────────────────────────────
  PARAMS: {
    DORMANT_THRESHOLD_DAYS: 180,   // 休眠顧客の閾値（日数）
    SIMILAR_COMPANY_COUNT:  5,     // 類似企業の提案数
    MAX_DORMANT_PROCESS:    20,    // 1回のトリガーで処理する最大件数
  },
};
