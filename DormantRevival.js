/**
 * ============================================================
 *  DormantRevival.js - 機能4: 休眠顧客の掘り起こし
 * ============================================================
 *  トリガー実行を想定:
 *   - 最終接触日から180日以上経過した顧客を抽出
 *   - 業界ニュースを再検索
 *   - ニュースをフックにしたメール文面を Gemini で生成
 *   - Gmail下書き保存 or シートに出力
 * ============================================================
 */

const DormantRevival = (() => {

  /**
   * メール生成用システムプロンプト
   */
  const EMAIL_SYSTEM_PROMPT = `あなたは、日本のBtoB営業のプロフェッショナルです。
休眠顧客（しばらく連絡していなかった取引先）に対して、
「ご無沙汰しております」から始まる自然な再接触メールを作成してください。

【メール作成のルール】
1. 冒頭は「ご無沙汰しております。〇〇（担当者名）でございます。」から始める
2. 業界の最新ニュースやトレンドを自然なフックとして織り込む
3. 押し付けがましくない、あくまで情報提供や近況確認のトーン
4. 「もしよろしければ、近況をお聞かせいただけますと幸いです」程度の軟らかいCTA
5. 署名は含めない（送信時に自動付与を想定）
6. 全体で200〜400文字程度
7. ビジネスメールとして適切な敬語を使用する`;

  /**
   * 日次/週次トリガーから呼ばれるメイン関数
   * 休眠顧客を検出し、メールドラフトを生成する
   */
  function processAll() {
    console.log('=== 休眠顧客チェック開始 ===');

    const dormants = SheetHelper.getDormantCustomers();

    if (dormants.length === 0) {
      console.log('休眠顧客はありませんでした');
      return { processed: 0 };
    }

    console.log(`休眠顧客 ${dormants.length} 件を検出`);

    // 処理上限
    const maxProcess = SFA_CONFIG.PARAMS.MAX_DORMANT_PROCESS;
    const targets = dormants.slice(0, maxProcess);
    let processedCount = 0;

    targets.forEach(customer => {
      try {
        // 1. 業界ニュースを再検索
        const latestNews = IndustryAnalysis.refreshNews(
          customer.companyName,
          customer.industry
        );

        // 2. メールドラフトを生成
        const draft = _generateEmailDraft(customer, latestNews);

        // 3. 結果を「休眠顧客メール下書き」シートに出力
        SheetHelper.appendToSheet(SFA_CONFIG.SHEETS.DORMANT_DRAFTS, [
          new Date(),                               // 生成日
          customer.companyName,                      // 会社名
          customer.fullName,                         // 氏名
          customer.email,                            // Email
          customer.lastContact,                      // 最終接触日
          customer.dormantDays,                      // 経過日数
          latestNews,                                // 業界最新ニュース
          draft.subject,                             // メール件名
          draft.body,                                // メール本文
          '下書き',                                   // ステータス
        ]);

        // 4. Gmail下書きにも保存 (Emailがある場合)
        if (customer.email) {
          _saveGmailDraft(customer.email, draft.subject, draft.body);
        }

        processedCount++;
        console.log(`[${processedCount}/${targets.length}] ${customer.companyName} ${customer.fullName} - 処理完了`);

        // API レート制限対策: 1件ごとに少し待機
        Utilities.sleep(2000);

      } catch (e) {
        console.error(`[DormantRevival] ${customer.companyName} の処理失敗: ${e.message}`);
      }
    });

    console.log(`=== 休眠顧客チェック完了: ${processedCount}/${targets.length} 件処理 ===`);
    return { processed: processedCount, total: dormants.length };
  }

  /**
   * メール文面を Gemini で生成
   * @param {Object} customer  - 顧客データ
   * @param {string} newsText  - 最新ニュース
   * @return {Object} { subject, body }
   */
  function _generateEmailDraft(customer, newsText) {
    const userPrompt = `以下の情報を元に、休眠顧客への再接触メールを作成してください。

【顧客情報】
- 会社名: ${customer.companyName}
- 氏名: ${customer.fullName} 様
- 役職: ${customer.title || '不明'}
- 最終接触日: ${_formatDate(customer.lastContact)}
- 経過日数: ${customer.dormantDays}日
- 担当者名: ${customer.staffName || '（担当者名）'}

【業界の最新ニュース/トレンド】
${newsText || '特になし'}

以下のJSON形式で出力してください:
{
  "subject": "メール件名（30文字以内）",
  "body": "メール本文（200〜400文字）"
}`;

    try {
      const result = GeminiService.generateJson(EMAIL_SYSTEM_PROMPT, userPrompt);
      return {
        subject: result.subject || `${customer.companyName} ${customer.fullName}様 ご無沙汰しております`,
        body:    result.body || _fallbackEmailBody(customer),
      };
    } catch (e) {
      console.warn('[DormantRevival] メール生成失敗、フォールバック使用: ' + e.message);
      return {
        subject: `${customer.companyName} ${customer.fullName}様 ご無沙汰しております`,
        body:    _fallbackEmailBody(customer),
      };
    }
  }

  /**
   * Gemini失敗時のフォールバックメール文面
   * @param {Object} customer
   * @return {string}
   */
  function _fallbackEmailBody(customer) {
    const staffName = customer.staffName || '（担当者名）';
    return `${customer.fullName} 様

ご無沙汰しております。${staffName}でございます。

以前はお忙しい中お時間をいただき、誠にありがとうございました。
その後、御社のご状況はいかがでしょうか。

もしよろしければ、改めてお話をお伺いする機会をいただけますと幸いです。
ご都合の良いタイミングがございましたら、お気軽にご連絡くださいませ。

何卒よろしくお願いいたします。`;
  }

  /**
   * Gmail の下書きとして保存
   * @param {string} to      - 宛先
   * @param {string} subject - 件名
   * @param {string} body    - 本文
   */
  function _saveGmailDraft(to, subject, body) {
    try {
      GmailApp.createDraft(to, subject, body);
      console.log(`[DormantRevival] Gmail下書き作成: → ${to}`);
    } catch (e) {
      // GmailApp が使えない環境(実行権限不足等)の場合はスキップ
      console.warn('[DormantRevival] Gmail下書き作成スキップ: ' + e.message);
    }
  }

  /**
   * 日付フォーマットヘルパー
   * @param {Date|string} date
   * @return {string}
   */
  function _formatDate(date) {
    if (!date) return '不明';
    try {
      return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy年MM月dd日');
    } catch (e) {
      return String(date);
    }
  }

  // ── Public API ──
  return {
    processAll,
  };

})();


// ==========================================
// トリガーから呼ばれるグローバル関数
// ==========================================

/**
 * 日次/週次トリガー実行用のエントリーポイント
 * GASのトリガーはグローバル関数のみ呼べるため、
 * ここからモジュールの processAll() を呼び出す。
 */
function dailyDormantCustomerCheck() {
  try {
    const result = DormantRevival.processAll();
    console.log(`休眠顧客処理結果: ${JSON.stringify(result)}`);
  } catch (e) {
    console.error('休眠顧客チェックでエラーが発生しました: ' + e.message);
  }
}
