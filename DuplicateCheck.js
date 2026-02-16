/**
 * ============================================================
 *  DuplicateCheck.js - 機能3: 社内接触履歴の照合（重複チェック）
 * ============================================================
 *  新規登録された名刺の「会社名」「氏名」が既存データに存在するか
 *  チェックし、ヒット時にアラートを生成。
 *  Slack / Chatwork への通知テンプレートも含む。
 * ============================================================
 */

const DuplicateChecker = (() => {

  /**
   * 重複チェックを実行
   * @param {string} companyName - 新規の会社名
   * @param {string} fullName    - 新規の氏名
   * @return {Object} { found: boolean, message: string, matches: Object[] }
   */
  function check(companyName, fullName) {
    if (!companyName && !fullName) {
      return { found: false, message: '', matches: [] };
    }

    const matches = SheetHelper.searchCustomers(companyName, fullName);

    if (matches.length === 0) {
      return { found: false, message: '', matches: [] };
    }

    // 重複メッセージを構築
    const messages = matches.map(m => {
      const staff = m.staffName || '不明';
      const date  = m.lastContact
        ? Utilities.formatDate(new Date(m.lastContact), 'Asia/Tokyo', 'yyyy/MM/dd')
        : '日付不明';
      return `${m.companyName} の ${m.fullName} さんは、${staff} が ${date} に接触済みです`;
    });

    const message = `【重複検知】\n` + messages.join('\n');

    // 通知ログに記録
    _logNotification('重複検知', companyName, fullName, message);

    return {
      found: true,
      message: message,
      matches: matches,
    };
  }

  /**
   * 通知を送信 (Slack / Chatwork / ログ)
   * @param {string} message - 通知メッセージ
   */
  function notify(message) {
    console.log('[DuplicateChecker] ' + message);

    // --- Slack 通知 ---
    if (SFA_CONFIG.ENV.SLACK_WEBHOOK_URL) {
      _sendSlack(message);
    }

    // --- Chatwork 通知 ---
    if (SFA_CONFIG.ENV.CHATWORK_API_TOKEN && SFA_CONFIG.ENV.CHATWORK_ROOM_ID) {
      _sendChatwork(message);
    }
  }

  /**
   * Slack Incoming Webhook で通知
   * @param {string} message
   */
  function _sendSlack(message) {
    try {
      const payload = {
        text: message,
        username: '名刺SFA Bot',
        icon_emoji: ':card_index:',
      };

      UrlFetchApp.fetch(SFA_CONFIG.ENV.SLACK_WEBHOOK_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      console.log('[DuplicateChecker] Slack通知を送信しました');
    } catch (e) {
      console.warn('[DuplicateChecker] Slack通知失敗: ' + e.message);
    }
  }

  /**
   * Chatwork API で通知
   * @param {string} message
   */
  function _sendChatwork(message) {
    try {
      const url = `https://api.chatwork.com/v2/rooms/${SFA_CONFIG.ENV.CHATWORK_ROOM_ID}/messages`;

      UrlFetchApp.fetch(url, {
        method: 'post',
        headers: {
          'X-ChatWorkToken': SFA_CONFIG.ENV.CHATWORK_API_TOKEN,
        },
        payload: {
          body: `[info][title]名刺SFA 重複検知[/title]${message}[/info]`,
        },
        muteHttpExceptions: true,
      });

      console.log('[DuplicateChecker] Chatwork通知を送信しました');
    } catch (e) {
      console.warn('[DuplicateChecker] Chatwork通知失敗: ' + e.message);
    }
  }

  /**
   * Google Chat (Webhook) で通知する場合のテンプレート
   * @param {string} webhookUrl - Google Chat Webhook URL
   * @param {string} message
   */
  function sendGoogleChat(webhookUrl, message) {
    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: message }),
        muteHttpExceptions: true,
      });
    } catch (e) {
      console.warn('[DuplicateChecker] Google Chat通知失敗: ' + e.message);
    }
  }

  /**
   * メール通知のテンプレート
   * @param {string} recipientEmail - 通知先メールアドレス
   * @param {string} message
   */
  function sendEmailAlert(recipientEmail, message) {
    try {
      MailApp.sendEmail({
        to: recipientEmail,
        subject: '【名刺SFA】重複顧客が検知されました',
        body: message,
      });
    } catch (e) {
      console.warn('[DuplicateChecker] メール通知失敗: ' + e.message);
    }
  }

  /**
   * 通知ログシートに記録
   * @param {string} type
   * @param {string} company
   * @param {string} name
   * @param {string} message
   */
  function _logNotification(type, company, name, message) {
    try {
      const notifyTargets = [];
      if (SFA_CONFIG.ENV.SLACK_WEBHOOK_URL) notifyTargets.push('Slack');
      if (SFA_CONFIG.ENV.CHATWORK_API_TOKEN) notifyTargets.push('Chatwork');
      if (notifyTargets.length === 0) notifyTargets.push('ログのみ');

      SheetHelper.appendToSheet(SFA_CONFIG.SHEETS.NOTIFICATION_LOG, [
        new Date(),
        type,
        company,
        name,
        message,
        notifyTargets.join(', '),
      ]);
    } catch (e) {
      console.warn('[DuplicateChecker] ログ記録失敗: ' + e.message);
    }
  }

  // ── Public API ──
  return {
    check,
    notify,
    sendGoogleChat,
    sendEmailAlert,
  };

})();
