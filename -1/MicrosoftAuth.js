/**
 * ============================================================
 *  MicrosoftAuth.js - Microsoft 365 OAuth2 認証管理
 * ============================================================
 *  Apps Script OAuth2 ライブラリを使用して Azure AD 認証を管理する。
 *  依存: OAuth2 ライブラリ (1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBmN0d0ZSrq)
 * ============================================================
 */

const MicrosoftAuth = (() => {

  const SERVICE_NAME = 'microsoft';

  /**
   * OAuth2 サービスを構築
   * @return {OAuth2.Service}
   */
  function getService() {
    const ms = SFA_CONFIG.MICROSOFT;
    const tenantId = ms.TENANT_ID || 'common';

    return OAuth2.createService(SERVICE_NAME)
      .setAuthorizationBaseUrl(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`)
      .setTokenUrl(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`)
      .setClientId(ms.CLIENT_ID)
      .setClientSecret(ms.CLIENT_SECRET)
      .setCallbackFunction('microsoftAuthCallback')
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope(ms.SCOPES.join(' '))
      .setParam('response_type', 'code')
      .setParam('prompt', 'consent');
  }

  /**
   * 認証 URL を返す
   * @return {string} 認証URL
   */
  function getAuthorizationUrl() {
    return getService().getAuthorizationUrl();
  }

  /**
   * トークンの有効性を確認
   * @return {boolean}
   */
  function isAuthorized() {
    return getService().hasAccess();
  }

  /**
   * Graph API 呼び出し用アクセストークンを取得（自動リフレッシュ）
   * @return {string} アクセストークン
   */
  function getAccessToken() {
    const service = getService();
    if (!service.hasAccess()) {
      throw new Error('Microsoft 認証が必要です。設定ページから認証を行ってください。');
    }
    return service.getAccessToken();
  }

  /**
   * トークンをリセット（ログアウト）
   */
  function logout() {
    getService().reset();
    console.log('[MicrosoftAuth] トークンをリセットしました');
  }

  // ── Public API ──
  return {
    getService,
    getAuthorizationUrl,
    isAuthorized,
    getAccessToken,
    logout,
  };

})();


// ==========================================
// OAuth2 コールバック (グローバル関数)
// ==========================================

/**
 * OAuth2 認証コールバック
 * @param {Object} request - コールバックリクエスト
 * @return {HtmlOutput}
 */
function microsoftAuthCallback(request) {
  const service = MicrosoftAuth.getService();
  const authorized = service.handleCallback(request);

  if (authorized) {
    return HtmlService.createHtmlOutput(
      '<h3>認証成功</h3><p>Microsoft 365 との連携が完了しました。このウィンドウを閉じてください。</p>' +
      '<script>setTimeout(function(){ window.close(); }, 3000);</script>'
    );
  } else {
    return HtmlService.createHtmlOutput(
      '<h3>認証失敗</h3><p>認証に失敗しました。再度お試しください。</p>'
    );
  }
}
