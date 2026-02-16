/**
 * 名刺解析ツール v2 - Web UI & Drive Trigger
 * 使用モデル: gemini-3-flash-preview
 */

// 設定項目
const CONFIG = {
  SHEET_NAME: 'シート1',
  GEMINI_MODEL: 'gemini-3-flash-preview',
  API_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent',
};

// ▼▼▼ ここに直接入力してください ▼▼▼
const ENV = {
  FOLDER_ID: '1mDptHj7xmXw89pwBQI71EXU2UnKBRL5l', 
  API_KEY:   'AIzaSyC1BgtDMwt5WfK5iGm9nYSVYBvLzam_SK4'
};
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

// ==========================================
// 1. Web UI用関数 (メイン)
// ==========================================

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('名刺AI解析スキャナー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Web画面から呼び出される処理
 */
function analyzeFromWeb(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const folder = DriveApp.getFolderById(ENV.FOLDER_ID);
  
  const mode = formData.mode; 
  const images = formData.images;
  const timestamp = new Date();
  
  // 1. 画像をドライブに保存
  const savedFileUrls = images.map((img, index) => {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(img.data), 
      img.mimeType, 
      `WEB_UPLOAD_${mode}_${timestamp.getTime()}_${index}.jpg`
    );
    const file = folder.createFile(blob);
    return file.getUrl();
  });

  // 2. Geminiへのリクエスト作成
  const parts = [];
  
  // プロンプト取得
  const promptText = (mode === 'merge') ? getMergePrompt() : getMultiPrompt();
  parts.push({ text: promptText });

  images.forEach(img => {
    parts.push({
      inline_data: { mime_type: img.mimeType, data: img.data }
    });
  });

  // 3. API実行
  const result = callGeminiAPI(parts);
  
  // 結果が単一オブジェクトの場合も配列化して扱う
  let resultsArray = Array.isArray(result) ? result : [result];

  // 4. スプレッドシートへ書き込み
  // IDやモードは含めず、純粋に必要な情報のみを書き込みます
  resultsArray.forEach(cardData => {
    const finalName = cardData.fullName || cardData.name || cardData.personName || '氏名不明';

    // 出力順序: 
    // 1.タイムスタンプ, 2.会社名, 3.氏名, 4.役職, 5.メール, 6.電話番号, 7.住所, 8.ウェブサイト, 9.画像
    sheet.appendRow([
      timestamp,
      cardData.companyName || '',
      finalName,
      cardData.title || '',
      cardData.email || '',
      cardData.phone || '',
      cardData.address || '',
      cardData.website || '',
      savedFileUrls[0] // 画像リンク（複数枚アップしても1枚目のリンクを代表として記録）
    ]);
  });

  return { success: true, count: resultsArray.length, data: resultsArray };
}


// ==========================================
// 2. プロンプト定義
// ==========================================

function getMergePrompt() {
  return `
    提供された画像は、同一人物の1枚の名刺の「表面」と「裏面」です（または複数枚）。
    両面の情報を統合し、最も正確な1つのJSONデータを作成してください。
    
    【重要：氏名の抽出ルール】
    1. 名刺の中で「最も大きく記載されている人物名」を "fullName" としてください。
    2. 会社名や役職と混同しないように注意してください。
    3. 漢字とローマ字がある場合は、漢字を優先してください。
    4. 氏名が見つからない場合は null ではなく空文字にしてください。

    出力フォーマット（JSON）:
    {
      "companyName": "会社名",
      "fullName": "氏名",
      "title": "役職",
      "email": "メールアドレス",
      "phone": "電話番号",
      "address": "住所",
      "website": "URL"
    }
  `;
}

function getMultiPrompt() {
  return `
    画像内の【全ての名刺】を検出し、それぞれの情報を抽出してください。
    
    【重要：氏名の抽出ルール】
    1. 各名刺の中で「最も大きく記載されている人物名」を必ず抽出してください。
    2. "fullName" というキーを必ず使用してください。

    出力フォーマット（JSON配列）:
    [
      {
        "companyName": "会社名",
        "fullName": "氏名",
        "title": "役職",
        "email": "メール",
        "phone": "電話",
        "address": "住所",
        "website": "URL"
      }
    ]
  `;
}


// ==========================================
// 3. Gemini API 呼び出し & ユーティリティ
// ==========================================

function callGeminiAPI(parts) {
  const payload = {
    contents: [{ parts: parts }],
    generationConfig: {
      response_mime_type: "application/json",
      temperature: 0.1
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const url = `${CONFIG.API_ENDPOINT}?key=${ENV.API_KEY}`;
  
  for (let i = 0; i < 3; i++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        let text = json.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (!text) throw new Error("Geminiからの応答が空です");

        // クリーニング処理
        text = text.replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '');
        
        return JSON.parse(text);
      }
      Utilities.sleep(1000 * (i + 1));
    } catch (e) {
      console.warn(`Retry ${i}: ${e.message}`);
      if (i === 2) throw e;
    }
  }
}