/**
 * ============================================================
 *  OutlookService.js - Microsoft Graph API ラッパー
 * ============================================================
 *  Outlook メール履歴の検索、下書き作成を提供。
 *  GeminiService.js と同パターン (IIFE, retry, muteHttpExceptions)。
 *  依存: MicrosoftAuth.js
 * ============================================================
 */

const OutlookService = (() => {

  const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
  const MAX_RETRIES = 3;
  const RETRY_DELAY = 1500; // ms

  /**
   * Graph API 共通フェッチ (リトライ付き)
   * @param {string} endpoint - Graph API エンドポイント (/me/... 等)
   * @param {Object} [options] - fetch オプション (method, payload 等)
   * @return {Object} パース済み JSON レスポンス
   */
  function _graphFetch(endpoint, options) {
    const url = GRAPH_BASE + endpoint;
    const token = MicrosoftAuth.getAccessToken();

    const fetchOptions = {
      method: (options && options.method) || 'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
      },
      muteHttpExceptions: true,
    };

    if (options && options.payload) {
      fetchOptions.payload = JSON.stringify(options.payload);
    }

    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      try {
        const response = UrlFetchApp.fetch(url, fetchOptions);
        const code = response.getResponseCode();

        if (code >= 200 && code < 300) {
          const text = response.getContentText();
          return text ? JSON.parse(text) : {};
        }

        // 429 (Rate Limit) or 503 (Service Unavailable) → リトライ
        if (code === 429 || code === 503) {
          console.warn(`[OutlookService] HTTP ${code} - リトライ ${attempt + 1}/${MAX_RETRIES}`);
          Utilities.sleep(RETRY_DELAY * (attempt + 1));
          continue;
        }

        throw new Error(`Graph API エラー: HTTP ${code} - ${response.getContentText().substring(0, 300)}`);

      } catch (e) {
        console.warn(`[OutlookService] 試行 ${attempt + 1} 失敗: ${e.message}`);
        if (attempt === MAX_RETRIES - 1) throw e;
        Utilities.sleep(RETRY_DELAY * (attempt + 1));
      }
    }

    throw new Error('[OutlookService] 最大リトライ回数に達しました');
  }

  /**
   * 指定メールアドレスとの最終メール日付を取得
   * 送信メール・受信メール両方を検索し、最新の日付を返す
   * @param {string} emailAddress - 検索対象のメールアドレス
   * @return {Date|null} 最終接触日 (なければ null)
   */
  function getLastEmailDate(emailAddress) {
    if (!emailAddress) return null;

    let latestDate = null;

    // 送信メール検索
    try {
      const sentResult = _graphFetch(
        `/me/mailFolders/SentItems/messages?$search="to:${emailAddress}"&$top=1&$orderby=sentDateTime desc&$select=sentDateTime`
      );
      if (sentResult.value && sentResult.value.length > 0) {
        const sentDate = new Date(sentResult.value[0].sentDateTime);
        if (!latestDate || sentDate > latestDate) {
          latestDate = sentDate;
        }
      }
    } catch (e) {
      console.warn(`[OutlookService] 送信メール検索失敗 (${emailAddress}): ${e.message}`);
    }

    // 受信メール検索
    try {
      const receivedResult = _graphFetch(
        `/me/messages?$filter=from/emailAddress/address eq '${emailAddress}'&$top=1&$orderby=receivedDateTime desc&$select=receivedDateTime`
      );
      if (receivedResult.value && receivedResult.value.length > 0) {
        const receivedDate = new Date(receivedResult.value[0].receivedDateTime);
        if (!latestDate || receivedDate > latestDate) {
          latestDate = receivedDate;
        }
      }
    } catch (e) {
      console.warn(`[OutlookService] 受信メール検索失敗 (${emailAddress}): ${e.message}`);
    }

    return latestDate;
  }

  /**
   * 全顧客のメール履歴をスキャンし、最終接触日 (Column I) を自動更新
   * @return {Object} { updated: number, errors: number, total: number }
   */
  function syncAllContactDates() {
    console.log('=== Outlook メール同期開始 ===');

    const sheet = SheetHelper.getMainSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('同期対象のデータがありません');
      return { updated: 0, errors: 0, total: 0 };
    }

    const COL = SFA_CONFIG.COL;
    const batchSize = SFA_CONFIG.MICROSOFT.SYNC_BATCH_SIZE;
    const delayMs = SFA_CONFIG.MICROSOFT.SYNC_DELAY_MS;

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    let updatedCount = 0;
    let errorCount = 0;

    for (let i = 0; i < data.length; i++) {
      const email = data[i][COL.EMAIL];
      if (!email) continue;

      try {
        const lastDate = getLastEmailDate(email);
        if (lastDate) {
          const currentLastContact = data[i][COL.LAST_CONTACT];
          // 既存の最終接触日より新しい場合のみ更新
          if (!currentLastContact || new Date(lastDate) > new Date(currentLastContact)) {
            const rowNum = i + 2; // 1-based, ヘッダー分+1
            sheet.getRange(rowNum, COL.LAST_CONTACT + 1).setValue(lastDate);
            updatedCount++;
            console.log(`[OutlookService] ${email} の最終接触日を更新: ${lastDate}`);
          }
        }
      } catch (e) {
        errorCount++;
        console.warn(`[OutlookService] ${email} の同期失敗: ${e.message}`);
      }

      // バッチ間隔
      if ((i + 1) % batchSize === 0) {
        console.log(`[OutlookService] ${i + 1}/${data.length} 件処理済み、待機中...`);
        Utilities.sleep(delayMs);
      }
    }

    console.log(`=== Outlook メール同期完了: 更新 ${updatedCount} 件, エラー ${errorCount} 件 ===`);
    return { updated: updatedCount, errors: errorCount, total: data.length };
  }

  /**
   * Outlook に下書きメッセージを作成
   * @param {string} to      - 宛先メールアドレス
   * @param {string} subject - 件名
   * @param {string} body    - 本文
   * @return {Object} 作成されたメッセージ
   */
  function createDraft(to, subject, body) {
    const payload = {
      subject: subject,
      body: {
        contentType: 'Text',
        content: body,
      },
      toRecipients: [{
        emailAddress: {
          address: to,
        },
      }],
      isDraft: true,
    };

    const result = _graphFetch('/me/messages', {
      method: 'post',
      payload: payload,
    });

    console.log(`[OutlookService] Outlook下書き作成: → ${to}`);
    return result;
  }

  // ── Public API ──
  return {
    getLastEmailDate,
    syncAllContactDates,
    createDraft,
  };

})();
