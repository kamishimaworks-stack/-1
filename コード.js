/**
 * ============================================================
 *  名刺SFA (Sales Force Automation) ツール v3
 *  メインエントリーポイント & 既存OCR機能
 * ============================================================
 *  使用モデル: Gemini (設定で切替可能)
 *  依存ファイル:
 *    - Config.js          … 全設定値
 *    - GeminiService.js   … Gemini API ラッパー
 *    - SnsResearch.js     … 機能1: SNS・Web情報リサーチ
 *    - IndustryAnalysis.js… 機能2: 業界ニュース・企業分析
 *    - DuplicateCheck.js  … 機能3: 重複チェック・接触履歴照合
 *    - DormantRevival.js  … 機能4: 休眠顧客掘り起こし
 *    - SimilarCompany.js  … 機能5: 類似企業・ターゲット拡張
 *    - SheetHelper.js     … スプレッドシート操作ユーティリティ
 * ============================================================
 */

// ==========================================
// Webアプリ エントリーポイント
// ==========================================

/**
 * Webアプリのエントリーポイント
 */
function doGet(e) {
  const page = e && e.parameter && e.parameter.page;

  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('dashboard')
      .evaluate()
      .setTitle('SFA ダッシュボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('名刺AI解析スキャナー + SFA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLテンプレートからインクルード読込
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ==========================================
// メイン処理: Web画面からのアップロード解析
// ==========================================

/**
 * Web画面から呼び出される名刺解析処理 (SFA機能統合版)
 * @param {Object} formData - { mode: 'merge'|'multi', images: [...], staffName: '担当者名' }
 * @return {Object} 解析結果
 */
function analyzeFromWeb(formData) {
  const sheet = SheetHelper.getMainSheet();
  const folder = DriveApp.getFolderById(SFA_CONFIG.ENV.FOLDER_ID);

  const mode      = formData.mode;
  const images    = formData.images;
  const staffName = formData.staffName || '';
  const timestamp = new Date();

  // ── 1. 画像をドライブに保存 ──
  const savedFileUrls = images.map((img, index) => {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(img.data),
      img.mimeType,
      `WEB_UPLOAD_${mode}_${Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyyMMdd_HHmmss')}_${index}.jpg`
    );
    return folder.createFile(blob).getUrl();
  });

  // ── 2. Gemini OCR 実行 ──
  const parts = [];
  parts.push({ text: (mode === 'merge') ? getMergePrompt() : getMultiPrompt() });
  images.forEach(img => {
    parts.push({ inline_data: { mime_type: img.mimeType, data: img.data } });
  });

  const result = GeminiService.callWithParts(parts);
  let resultsArray = Array.isArray(result) ? result : [result];

  // ── 3. 各名刺データについてSFA機能を実行 ──
  const sfaResults = [];

  resultsArray.forEach(cardData => {
    const fullName    = cardData.fullName || cardData.name || cardData.personName || '';
    const companyName = cardData.companyName || '';
    const title       = cardData.title || '';
    const email       = cardData.email || '';
    const phone       = cardData.phone || '';
    const address     = cardData.address || '';
    const website     = cardData.website || '';

    // --- 機能3: 重複チェック (書き込み前に実行) ---
    let duplicateAlert = '';
    try {
      const dupResult = DuplicateChecker.check(companyName, fullName);
      if (dupResult.found) {
        duplicateAlert = dupResult.message;
        DuplicateChecker.notify(dupResult.message);
      }
    } catch (e) {
      console.warn('[DuplicateCheck] ' + e.message);
    }

    // --- 機能1: SNS・Webリサーチ ---
    let snsInfo = {};
    try {
      snsInfo = SnsResearch.generateSearchUrls(fullName, companyName, website);
    } catch (e) {
      console.warn('[SnsResearch] ' + e.message);
    }

    // --- 機能2: 業界ニュース・企業分析 ---
    let industryInfo = {};
    try {
      industryInfo = IndustryAnalysis.analyze(companyName, title);
    } catch (e) {
      console.warn('[IndustryAnalysis] ' + e.message);
    }

    // --- 機能5: 類似企業・ターゲット拡張 ---
    let similarCompanies = '';
    try {
      const scResult = SimilarCompanyFinder.find(companyName, industryInfo.industry || '');
      similarCompanies = scResult.summary || '';
      // 別シートにも出力
      SimilarCompanyFinder.writeToSheet(companyName, scResult.companies || []);
    } catch (e) {
      console.warn('[SimilarCompany] ' + e.message);
    }

    // ── 4. スプレッドシートへ書き込み (拡張カラム) ──
    //  A: 登録日, B: 会社名, C: 氏名, D: 役職, E: Email,
    //  F: 電話番号, G: 住所, H: WebサイトURL, I: 最終接触日,
    //  J: 担当者名, K: 名刺画像,
    //  L: X検索URL, M: Facebook検索URL, N: Instagram検索URL,
    //  O: YouTube検索URL, P: TikTok検索URL, Q: 企業サイトURL,
    //  R: 推定業種, S: 業界トレンド・ニュース, T: 想定課題,
    //  U: 類似企業リスト, V: 重複アラート
    sheet.appendRow([
      timestamp,                                          // A: 登録日
      companyName,                                        // B: 会社名
      fullName,                                           // C: 氏名
      title,                                              // D: 役職
      email,                                              // E: Email
      phone,                                              // F: 電話番号
      address,                                            // G: 住所
      website,                                            // H: WebサイトURL
      timestamp,                                          // I: 最終接触日 (初回=登録日)
      staffName,                                          // J: 担当者名
      savedFileUrls[0] || '',                             // K: 名刺画像
      snsInfo.xUrl         || '',                         // L: X検索URL
      snsInfo.facebookUrl  || '',                         // M: Facebook検索URL
      snsInfo.instagramUrl || '',                         // N: Instagram検索URL
      snsInfo.youtubeUrl   || '',                         // O: YouTube検索URL
      snsInfo.tiktokUrl    || '',                         // P: TikTok検索URL
      snsInfo.companySiteUrl || website || '',             // Q: 企業サイト
      industryInfo.industry       || '',                  // R: 推定業種
      industryInfo.trends         || '',                  // S: 業界トレンド
      industryInfo.challenges     || '',                  // T: 想定課題
      similarCompanies,                                   // U: 類似企業
      duplicateAlert                                      // V: 重複アラート
    ]);

    sfaResults.push({
      companyName,
      fullName,
      title,
      email,
      duplicateAlert,
      industry: industryInfo.industry || '',
      trends: industryInfo.trends || '',
      challenges: industryInfo.challenges || '',
      similarCompanies
    });
  });

  return { success: true, count: sfaResults.length, data: sfaResults };
}


// ==========================================
// OCR用プロンプト (既存)
// ==========================================

function getMergePrompt() {
  return `
あなたは名刺OCRの専門AIです。
提供された画像は、同一人物の1枚の名刺の「表面」と「裏面」です。
両面の情報を統合し、最も正確な1つのJSONデータを作成してください。

【重要：氏名の抽出ルール】
1. 名刺の中で「最も大きく記載されている人物名」を "fullName" としてください。
2. 会社名や役職と混同しないように注意してください。
3. 漢字とローマ字がある場合は、漢字を優先してください。
4. 氏名が見つからない場合は空文字にしてください。

出力フォーマット（JSON）:
{
  "companyName": "会社名",
  "fullName": "氏名（姓 名）",
  "title": "役職",
  "email": "メールアドレス",
  "phone": "電話番号（ハイフン付き）",
  "address": "住所",
  "website": "URL（https://を含む）"
}`;
}

function getMultiPrompt() {
  return `
あなたは名刺OCRの専門AIです。
画像内の【全ての名刺】を検出し、それぞれの情報を抽出してください。

【重要：氏名の抽出ルール】
1. 各名刺の中で「最も大きく記載されている人物名」を必ず抽出してください。
2. "fullName" というキーを必ず使用してください。

出力フォーマット（JSON配列）:
[
  {
    "companyName": "会社名",
    "fullName": "氏名（姓 名）",
    "title": "役職",
    "email": "メールアドレス",
    "phone": "電話番号（ハイフン付き）",
    "address": "住所",
    "website": "URL（https://を含む）"
  }
]`;
}


// ==========================================
// セットアップ用ユーティリティ
// ==========================================

/**
 * 初回セットアップ: ヘッダー行の作成 & 必要シートの生成
 * GASエディタから手動で1回実行してください。
 */
function setupSFA() {
  console.log('=== SFA セットアップ開始 ===');

  // メインシートのヘッダー
  SheetHelper.ensureHeaders(SFA_CONFIG.SHEETS.MAIN, SFA_CONFIG.HEADERS.MAIN);

  // 類似企業シート
  SheetHelper.ensureHeaders(SFA_CONFIG.SHEETS.SIMILAR_COMPANIES, SFA_CONFIG.HEADERS.SIMILAR_COMPANIES);

  // 休眠顧客メール下書きシート
  SheetHelper.ensureHeaders(SFA_CONFIG.SHEETS.DORMANT_DRAFTS, SFA_CONFIG.HEADERS.DORMANT_DRAFTS);

  // 通知ログシート
  SheetHelper.ensureHeaders(SFA_CONFIG.SHEETS.NOTIFICATION_LOG, SFA_CONFIG.HEADERS.NOTIFICATION_LOG);

  console.log('=== SFA セットアップ完了 ===');
  console.log('次のステップ:');
  console.log('1. スクリプトプロパティ「Gemini_API」に API キーを設定し、Config.js の ENV セクションにフォルダIDを設定');
  console.log('2. 日次トリガーを設定: setupDailyTrigger() を実行');
}

/**
 * 日次トリガーの設定 (休眠顧客チェック用)
 */
function setupDailyTrigger() {
  // 既存の同名トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'dailyDormantCustomerCheck') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎日 AM9:00 に実行
  ScriptApp.newTrigger('dailyDormantCustomerCheck')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  console.log('日次トリガーを設定しました (毎日 9:00 実行)');
}

/**
 * 週次トリガーの設定 (休眠顧客チェック用 - 週1回で十分な場合)
 */
function setupWeeklyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'dailyDormantCustomerCheck') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('dailyDormantCustomerCheck')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  console.log('週次トリガーを設定しました (毎週月曜 9:00 実行)');
}
