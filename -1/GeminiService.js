/**
 * ============================================================
 *  GeminiService.js - Gemini API 統合ラッパー
 * ============================================================
 *  全モジュールが共通で使う Gemini API 呼び出しロジック。
 *  リトライ、レート制限対応、JSON パース、テキスト生成を提供。
 * ============================================================
 */

const GeminiService = (() => {

  /**
   * Gemini API のエンドポイントURLを生成
   * @param {string} [model] - モデル名（省略時は設定値）
   * @return {string} URL
   */
  function _getEndpoint(model) {
    const m = model || SFA_CONFIG.GEMINI.MODEL;
    return `${SFA_CONFIG.GEMINI.API_BASE}${m}:generateContent?key=${SFA_CONFIG.ENV.API_KEY}`;
  }

  /**
   * 共通リクエスト送信 (リトライ付き)
   * @param {Object} payload - Gemini API リクエストボディ
   * @return {string} レスポンステキスト
   */
  function _sendRequest(payload) {
    const url = _getEndpoint();
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    const maxRetries = SFA_CONFIG.GEMINI.MAX_RETRIES;

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        const response = UrlFetchApp.fetch(url, options);
        const code = response.getResponseCode();

        if (code === 200) {
          const json = JSON.parse(response.getContentText());
          const text = json.candidates?.[0]?.content?.parts?.[0]?.text;
          if (!text) throw new Error('Gemini APIからの応答テキストが空です');
          return text;
        }

        // 429 (Rate Limit) or 503 (Service Unavailable) → リトライ
        if (code === 429 || code === 503) {
          console.warn(`[GeminiService] HTTP ${code} - リトライ ${attempt + 1}/${maxRetries}`);
          Utilities.sleep(SFA_CONFIG.GEMINI.RETRY_DELAY * (attempt + 1));
          continue;
        }

        // その他のエラー
        throw new Error(`Gemini API エラー: HTTP ${code} - ${response.getContentText().substring(0, 300)}`);

      } catch (e) {
        console.warn(`[GeminiService] 試行 ${attempt + 1} 失敗: ${e.message}`);
        if (attempt === maxRetries - 1) throw e;
        Utilities.sleep(SFA_CONFIG.GEMINI.RETRY_DELAY * (attempt + 1));
      }
    }

    throw new Error('[GeminiService] 最大リトライ回数に達しました');
  }

  /**
   * パーツ配列でのAPI呼び出し (画像OCR等)
   * @param {Array} parts - Gemini API の parts 配列
   * @return {Object|Array} パース済みJSONレスポンス
   */
  function callWithParts(parts) {
    const payload = {
      contents: [{ parts: parts }],
      generationConfig: {
        response_mime_type: 'application/json',
        temperature: 0.1,
      },
    };

    const text = _sendRequest(payload);
    return _parseJson(text);
  }

  /**
   * テキストプロンプトでの呼び出し → JSON レスポンス
   * @param {string} systemPrompt - システムプロンプト
   * @param {string} userPrompt   - ユーザープロンプト
   * @return {Object} パース済みJSON
   */
  function generateJson(systemPrompt, userPrompt) {
    const payload = {
      system_instruction: {
        parts: [{ text: systemPrompt }],
      },
      contents: [{
        parts: [{ text: userPrompt }],
      }],
      generationConfig: {
        response_mime_type: 'application/json',
        temperature: SFA_CONFIG.GEMINI.TEMPERATURE,
      },
    };

    const text = _sendRequest(payload);
    return _parseJson(text);
  }

  /**
   * テキストプロンプトでの呼び出し → プレーンテキスト
   * @param {string} systemPrompt - システムプロンプト
   * @param {string} userPrompt   - ユーザープロンプト
   * @return {string} レスポンステキスト
   */
  function generateText(systemPrompt, userPrompt) {
    const payload = {
      system_instruction: {
        parts: [{ text: systemPrompt }],
      },
      contents: [{
        parts: [{ text: userPrompt }],
      }],
      generationConfig: {
        temperature: SFA_CONFIG.GEMINI.TEMPERATURE,
      },
    };

    return _sendRequest(payload);
  }

  /**
   * Google Custom Search API を呼び出す (設定済みの場合)
   * @param {string} query - 検索クエリ
   * @param {number} [num=3] - 取得件数
   * @return {Array} 検索結果の配列 [{title, link, snippet}]
   */
  function customSearch(query, num) {
    const apiKey = SFA_CONFIG.ENV.CUSTOM_SEARCH_API_KEY;
    const cx     = SFA_CONFIG.ENV.CUSTOM_SEARCH_CX;

    if (!apiKey || !cx) {
      console.log('[GeminiService] Custom Search API未設定 - スキップ');
      return [];
    }

    num = num || 3;

    try {
      const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${cx}&q=${encodeURIComponent(query)}&num=${num}`;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

      if (response.getResponseCode() !== 200) {
        console.warn(`[CustomSearch] HTTP ${response.getResponseCode()}`);
        return [];
      }

      const data = JSON.parse(response.getContentText());
      return (data.items || []).map(item => ({
        title:   item.title || '',
        link:    item.link || '',
        snippet: item.snippet || '',
      }));

    } catch (e) {
      console.warn(`[CustomSearch] エラー: ${e.message}`);
      return [];
    }
  }

  /**
   * JSONテキストのクリーニング & パース
   * @param {string} text
   * @return {Object|Array}
   */
  function _parseJson(text) {
    // Markdown コードブロックの除去
    let cleaned = text
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/\s*```$/i, '')
      .trim();

    try {
      return JSON.parse(cleaned);
    } catch (e) {
      // フォールバック: JSON部分のみ抽出
      const match = cleaned.match(/[\[{][\s\S]*[\]}]/);
      if (match) {
        return JSON.parse(match[0]);
      }
      throw new Error(`JSON パースエラー: ${e.message}\n原文: ${cleaned.substring(0, 200)}`);
    }
  }

  // ── Public API ──
  return {
    callWithParts,
    generateJson,
    generateText,
    customSearch,
  };

})();
