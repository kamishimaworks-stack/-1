/**
 * ============================================================
 *  SnsResearch.js - 機能1: SNS・Web情報の自動リサーチ
 * ============================================================
 *  名刺の「氏名」と「会社名」を元に、各SNSの
 *  プロフィール検索用URLを生成し、企業サイトURLを推論。
 * ============================================================
 */

const SnsResearch = (() => {

  /**
   * SNSプロフィール検索URLとWebサイトURLを生成
   * @param {string} fullName    - 氏名
   * @param {string} companyName - 会社名
   * @param {string} [website]   - OCRで取得済みのWebサイトURL
   * @return {Object} SNS検索URL群
   */
  function generateSearchUrls(fullName, companyName, website) {
    if (!fullName && !companyName) {
      return _emptyResult();
    }

    // 検索クエリを構築
    const personQuery = [fullName, companyName].filter(Boolean).join(' ');
    const companyQuery = companyName || '';

    // 各SNSの検索URL生成
    const result = {
      xUrl:          _buildXSearchUrl(personQuery),
      facebookUrl:   _buildFacebookSearchUrl(personQuery),
      instagramUrl:  _buildInstagramSearchUrl(personQuery),
      youtubeUrl:    _buildYouTubeSearchUrl(personQuery),
      tiktokUrl:     _buildTikTokSearchUrl(personQuery),
      companySiteUrl: website || '',
    };

    // 企業サイトURLがない場合、Gemini で推論を試みる
    if (!result.companySiteUrl && companyName) {
      try {
        result.companySiteUrl = _inferCompanySite(companyName);
      } catch (e) {
        console.warn('[SnsResearch] 企業サイト推論失敗: ' + e.message);
        // Google検索URLをフォールバック
        result.companySiteUrl = `https://www.google.com/search?q=${encodeURIComponent(companyName + ' 公式サイト')}`;
      }
    }

    return result;
  }

  // ── 各SNSの検索URL生成 ──

  /**
   * X (旧Twitter) の検索URL
   */
  function _buildXSearchUrl(query) {
    return `https://x.com/search?q=${encodeURIComponent(query)}&src=typed_query&f=user`;
  }

  /**
   * Facebook の人物検索URL
   */
  function _buildFacebookSearchUrl(query) {
    return `https://www.facebook.com/search/people/?q=${encodeURIComponent(query)}`;
  }

  /**
   * Instagram の検索URL (Web版)
   */
  function _buildInstagramSearchUrl(query) {
    // Instagram はWeb検索が制限的なのでGoogle検索経由
    return `https://www.google.com/search?q=site:instagram.com+${encodeURIComponent(query)}`;
  }

  /**
   * YouTube の検索URL
   */
  function _buildYouTubeSearchUrl(query) {
    return `https://www.youtube.com/results?search_query=${encodeURIComponent(query)}&sp=EgIQAg%253D%253D`;
  }

  /**
   * TikTok の検索URL
   */
  function _buildTikTokSearchUrl(query) {
    return `https://www.tiktok.com/search/user?q=${encodeURIComponent(query)}`;
  }

  /**
   * Gemini を使って企業の公式サイトURLを推論
   * @param {string} companyName
   * @return {string} 推定URL
   */
  function _inferCompanySite(companyName) {
    const systemPrompt = `あなたは日本企業のデータベースに精通したアシスタントです。
企業名から公式Webサイトの URL を推定してください。
確信が持てない場合は、Google検索URLを返してください。`;

    const userPrompt = `以下の企業の公式WebサイトURLをJSON形式で回答してください。

企業名: ${companyName}

出力形式:
{
  "url": "https://example.co.jp",
  "confidence": "high" | "medium" | "low"
}`;

    const result = GeminiService.generateJson(systemPrompt, userPrompt);
    if (result && result.url) {
      return result.url;
    }
    return `https://www.google.com/search?q=${encodeURIComponent(companyName + ' 公式サイト')}`;
  }

  /**
   * 空の結果オブジェクトを返す
   */
  function _emptyResult() {
    return {
      xUrl: '',
      facebookUrl: '',
      instagramUrl: '',
      youtubeUrl: '',
      tiktokUrl: '',
      companySiteUrl: '',
    };
  }

  // ── Public API ──
  return {
    generateSearchUrls,
  };

})();
