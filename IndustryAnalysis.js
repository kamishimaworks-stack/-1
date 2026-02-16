/**
 * ============================================================
 *  IndustryAnalysis.js - 機能2: 業界最新ニュースと企業分析
 * ============================================================
 *  名刺の「会社名」と「業種(推測)」を元に、Gemini で:
 *   - 業界の直近トレンド・ニュースを取得
 *   - 企業が抱えていそうな課題を仮説生成
 * ============================================================
 */

const IndustryAnalysis = (() => {

  /**
   * 業界分析のシステムプロンプト
   */
  const SYSTEM_PROMPT = `あなたは日本市場に精通したビジネスアナリスト兼コンサルタントです。
営業担当者が初回商談の準備をする際に役立つ情報を提供してください。

【あなたの役割】
- 企業名と役職から、その企業の業種を正確に推定する
- その業種の最新トレンド・ニュース（直近6ヶ月以内）を3〜5件リストアップ
- その企業が直面していそうなビジネス課題の仮説を3〜5つ提示
- 営業アプローチに活用できる具体的な洞察を含める

【注意事項】
- 推測の場合は「推定」と明記する
- 具体的な数値やソース名がある場合は含める
- 日本語で回答する`;

  /**
   * 企業分析を実行
   * @param {string} companyName - 会社名
   * @param {string} [jobTitle]  - 名刺の役職 (業種推定の補助に使用)
   * @return {Object} { industry, trends, challenges, searchResults }
   */
  function analyze(companyName, jobTitle) {
    if (!companyName) {
      return { industry: '', trends: '', challenges: '' };
    }

    // Step 1: Custom Search API でリアルタイムニュースを取得 (設定済みの場合)
    let searchContext = '';
    const searchResults = GeminiService.customSearch(`${companyName} ニュース 最新`);
    if (searchResults.length > 0) {
      searchContext = '\n\n【参考: Web検索結果】\n' + searchResults.map((r, i) =>
        `${i + 1}. ${r.title}\n   ${r.snippet}\n   ${r.link}`
      ).join('\n');
    }

    // Step 2: Gemini で総合分析
    const userPrompt = `以下の企業について分析してください。

企業名: ${companyName}
名刺上の役職: ${jobTitle || '不明'}
${searchContext}

以下のJSON形式で回答してください:
{
  "industry": "推定される業種（例：IT・通信、製造業、不動産、コンサルティングなど）",
  "industryTrends": [
    "トレンド1: 具体的な説明",
    "トレンド2: 具体的な説明",
    "トレンド3: 具体的な説明"
  ],
  "estimatedChallenges": [
    "課題1: 具体的な仮説",
    "課題2: 具体的な仮説",
    "課題3: 具体的な仮説"
  ],
  "salesTip": "この企業への営業アプローチで活用できる一言アドバイス"
}`;

    try {
      const result = GeminiService.generateJson(SYSTEM_PROMPT, userPrompt);

      return {
        industry:    result.industry || '不明',
        trends:      _formatList(result.industryTrends, 'トレンド'),
        challenges:  _formatList(result.estimatedChallenges, '課題'),
        salesTip:    result.salesTip || '',
      };

    } catch (e) {
      console.error('[IndustryAnalysis] 分析失敗: ' + e.message);
      return {
        industry:   '分析失敗',
        trends:     'APIエラーにより取得できませんでした',
        challenges: 'APIエラーにより取得できませんでした',
      };
    }
  }

  /**
   * 業界ニュースのみを再検索 (休眠顧客の掘り起こし用)
   * @param {string} companyName
   * @param {string} industry
   * @return {string} ニュース要約テキスト
   */
  function refreshNews(companyName, industry) {
    if (!companyName) return '';

    // Custom Search で最新情報取得
    const query = industry
      ? `${industry} 最新ニュース トレンド 2025 2026`
      : `${companyName} 業界 最新ニュース`;

    let searchContext = '';
    const searchResults = GeminiService.customSearch(query, 5);
    if (searchResults.length > 0) {
      searchContext = '\n\n【最新Web検索結果】\n' + searchResults.map((r, i) =>
        `${i + 1}. ${r.title}: ${r.snippet}`
      ).join('\n');
    }

    const userPrompt = `以下の企業/業界の最新ニュースやトレンドを3つ簡潔にまとめてください。
営業メールのフックとして使えるような切り口でお願いします。

企業名: ${companyName}
業種: ${industry || '不明'}
${searchContext}

箇条書きで3つ、各50文字以内でまとめてください。`;

    try {
      return GeminiService.generateText(
        'あなたはビジネスニュースのキュレーターです。営業活動に役立つ簡潔な情報を提供してください。',
        userPrompt
      );
    } catch (e) {
      console.warn('[IndustryAnalysis] ニュース再検索失敗: ' + e.message);
      return '最新ニュースの取得に失敗しました';
    }
  }

  /**
   * 配列をセル用テキストにフォーマット
   * @param {string[]} items
   * @param {string} label
   * @return {string}
   */
  function _formatList(items, label) {
    if (!Array.isArray(items) || items.length === 0) {
      return `${label}情報なし`;
    }
    return items.map((item, i) => `${i + 1}. ${item}`).join('\n');
  }

  // ── Public API ──
  return {
    analyze,
    refreshNews,
  };

})();
