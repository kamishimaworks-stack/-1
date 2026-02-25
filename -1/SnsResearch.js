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
   * @param {string} [mainBusiness] - メイン事業キーワード（省略時はGeminiで推定）
   * @return {Object} SNS検索URL群
   */
  function generateSearchUrls(fullName, companyName, website, mainBusiness) {
    if (!fullName && !companyName) {
      return _emptyResult();
    }

    // 検索クエリを構築
    // 法人名がある場合は企業名で検索（個人名を含めるとヒットしにくい）
    // 法人名がなく個人名のみの場合は個人名で検索
    const snsQuery = companyName || fullName;

    // Instagram用: メイン事業キーワードを取得（同名企業との区別用）
    let businessKeyword = mainBusiness || '';
    if (!businessKeyword && companyName) {
      try {
        businessKeyword = _inferMainBusiness(companyName);
      } catch (e) {
        console.warn('[SnsResearch] メイン事業推定失敗: ' + e.message);
      }
    }

    // 各SNSの検索URL生成
    const result = {
      xUrl:          _buildXSearchUrl(snsQuery),
      facebookUrl:   _buildFacebookSearchUrl(snsQuery),
      instagramUrl:  _buildInstagramSearchUrl(snsQuery, businessKeyword),
      youtubeUrl:    _buildYouTubeSearchUrl(snsQuery),
      tiktokUrl:     _buildTikTokSearchUrl(snsQuery),
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
   * @param {string} query - 企業名
   * @param {string} [businessKeyword] - メイン事業キーワード（同名企業区別用）
   */
  function _buildInstagramSearchUrl(query, businessKeyword) {
    // 「Instagram 企業名 メイン事業」でGoogle検索（同名企業を除外しやすくする）
    const parts = ['Instagram', query];
    if (businessKeyword) parts.push(businessKeyword);
    return `https://www.google.com/search?q=${encodeURIComponent(parts.join(' '))}`;
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
   * Gemini を使って企業のメイン事業を短いキーワードで推定
   * @param {string} companyName
   * @return {string} メイン事業キーワード（例: 映像制作、Web開発、不動産）
   */
  function _inferMainBusiness(companyName) {
    const systemPrompt = 'あなたは日本企業に精通したアシスタントです。企業名からメイン事業を推定し、短いキーワード1つで回答してください。';
    const userPrompt = `以下の企業のメイン事業を、検索に使える短いキーワード1つ（2〜6文字程度）で回答してください。
余計な説明は不要です。キーワードのみ返してください。

企業名: ${companyName}

例:
- 株式会社○○映像 → 映像制作
- △△建設株式会社 → 建設
- □□テクノロジーズ → IT`;

    const result = GeminiService.generateText(systemPrompt, userPrompt);
    // 前後の空白・改行・記号を除去して返す
    return (result || '').trim().replace(/^[「『"']+|[」』"']+$/g, '');
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
