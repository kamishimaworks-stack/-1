/**
 * ============================================================
 *  TeamsNotification.js - Teams Webhook 通知
 * ============================================================
 *  Adaptive Card 形式で Teams チャネルに通知を送信する。
 *  DuplicateCheck.js 内 _sendSlack() と同じ Webhook パターン。
 *  Workflows 対応のペイロード形式を使用。
 * ============================================================
 */

const TeamsNotification = (() => {

  /**
   * Teams Webhook にペイロードを送信 (共通)
   * @param {Object} cardContent - Adaptive Card の content オブジェクト
   */
  function _sendToTeams(cardContent) {
    const webhookUrl = SFA_CONFIG.ENV.TEAMS_WEBHOOK_URL;
    if (!webhookUrl) {
      console.log('[TeamsNotification] Webhook URL 未設定 - スキップ');
      return;
    }

    const payload = {
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: cardContent,
      }],
    };

    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      console.log('[TeamsNotification] Teams通知を送信しました');
    } catch (e) {
      console.warn('[TeamsNotification] Teams通知失敗: ' + e.message);
    }
  }

  /**
   * 休眠顧客アラート (Adaptive Card)
   * @param {Object} params - { companyName, fullName, email, dormantDays, lastContact, staffName }
   */
  function sendDormantAlert(params) {
    const card = {
      '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          size: 'Medium',
          weight: 'Bolder',
          text: '休眠顧客アラート',
          color: 'Warning',
        },
        {
          type: 'FactSet',
          facts: [
            { title: '会社名', value: params.companyName || '-' },
            { title: '氏名', value: params.fullName || '-' },
            { title: 'Email', value: params.email || '-' },
            { title: '休眠日数', value: (params.dormantDays || '-') + ' 日' },
            { title: '最終接触日', value: params.lastContact || '-' },
            { title: '担当者', value: params.staffName || '-' },
          ],
        },
      ],
    };

    _sendToTeams(card);
  }

  /**
   * 重複検知アラート (Adaptive Card)
   * @param {string} message - 重複検知メッセージ
   * @param {Array} matches - マッチした顧客データ配列
   */
  function sendDuplicateAlert(message, matches) {
    const bodyItems = [
      {
        type: 'TextBlock',
        size: 'Medium',
        weight: 'Bolder',
        text: '重複顧客検知',
        color: 'Attention',
      },
      {
        type: 'TextBlock',
        text: message,
        wrap: true,
      },
    ];

    if (matches && matches.length > 0) {
      const facts = matches.map(m => ({
        title: m.companyName + ' / ' + m.fullName,
        value: '担当: ' + (m.staffName || '不明') + ' / 接触: ' + (m.lastContact || '不明'),
      }));
      bodyItems.push({
        type: 'FactSet',
        facts: facts,
      });
    }

    const card = {
      '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      body: bodyItems,
    };

    _sendToTeams(card);
  }

  /**
   * バッチ処理完了通知
   * @param {Object} params - { taskName, processed, total, errors }
   */
  function sendBatchResult(params) {
    const card = {
      '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          size: 'Medium',
          weight: 'Bolder',
          text: 'バッチ処理完了: ' + (params.taskName || ''),
        },
        {
          type: 'FactSet',
          facts: [
            { title: '処理件数', value: String(params.processed || 0) + ' / ' + String(params.total || 0) },
            { title: 'エラー', value: String(params.errors || 0) + ' 件' },
          ],
        },
      ],
    };

    _sendToTeams(card);
  }

  // ── Public API ──
  return {
    sendDormantAlert,
    sendDuplicateAlert,
    sendBatchResult,
  };

})();
