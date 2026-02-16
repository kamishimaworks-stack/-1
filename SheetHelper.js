/**
 * ============================================================
 *  SheetHelper.js - スプレッドシート操作ユーティリティ
 * ============================================================
 */

const SheetHelper = (() => {

  /**
   * アクティブなスプレッドシートを取得
   * @return {Spreadsheet}
   */
  function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  /**
   * メイン顧客シートを取得
   * @return {Sheet}
   */
  function getMainSheet() {
    return _getOrCreateSheet(SFA_CONFIG.SHEETS.MAIN);
  }

  /**
   * 指定名のシートを取得 (無ければ作成)
   * @param {string} sheetName
   * @return {Sheet}
   */
  function _getOrCreateSheet(sheetName) {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      console.log(`[SheetHelper] シート「${sheetName}」を新規作成しました`);
    }
    return sheet;
  }

  /**
   * ヘッダー行を確認・設定
   * @param {string} sheetName - シート名
   * @param {string[]} headers  - ヘッダー配列
   */
  function ensureHeaders(sheetName, headers) {
    const sheet = _getOrCreateSheet(sheetName);
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

    // ヘッダーが未設定 or 異なる場合のみ上書き
    const needsUpdate = headers.some((h, i) => firstRow[i] !== h);
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // ヘッダー行を太字 + 背景色
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');

      // 行を固定
      sheet.setFrozenRows(1);

      console.log(`[SheetHelper] シート「${sheetName}」のヘッダーを設定しました`);
    }
  }

  /**
   * メインシートから全データを取得 (ヘッダー除く)
   * @return {Object[]} 各行を連想配列化した配列
   */
  function getAllCustomers() {
    const sheet = getMainSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const COL = SFA_CONFIG.COL;

    return data.map((row, index) => ({
      rowIndex:      index + 2,  // シート上の行番号 (1-based, ヘッダー除く)
      registeredDate: row[COL.REGISTERED_DATE] || null,
      companyName:   row[COL.COMPANY_NAME] || '',
      fullName:      row[COL.FULL_NAME] || '',
      title:         row[COL.TITLE] || '',
      email:         row[COL.EMAIL] || '',
      phone:         row[COL.PHONE] || '',
      address:       row[COL.ADDRESS] || '',
      website:       row[COL.WEBSITE] || '',
      lastContact:   row[COL.LAST_CONTACT] || null,
      staffName:     row[COL.STAFF_NAME] || '',
      industry:      row[COL.INDUSTRY] || '',
      trends:        row[COL.TRENDS] || '',
      challenges:    row[COL.CHALLENGES] || '',
    }));
  }

  /**
   * 指定シートに1行追加
   * @param {string} sheetName
   * @param {Array} rowData
   */
  function appendToSheet(sheetName, rowData) {
    const sheet = _getOrCreateSheet(sheetName);
    sheet.appendRow(rowData);
  }

  /**
   * メインシートの特定セルを更新
   * @param {number} row - 行番号 (1-based)
   * @param {number} col - 列番号 (1-based)
   * @param {*} value
   */
  function updateCell(row, col, value) {
    const sheet = getMainSheet();
    sheet.getRange(row, col + 1).setValue(value); // COLは0-basedなので+1
  }

  /**
   * 会社名または氏名で検索
   * @param {string} companyName
   * @param {string} fullName
   * @return {Object[]} マッチした顧客データ配列
   */
  function searchCustomers(companyName, fullName) {
    const all = getAllCustomers();
    const results = [];

    const companyLower = (companyName || '').toLowerCase().trim();
    const nameLower    = (fullName || '').toLowerCase().trim();

    all.forEach(customer => {
      const cName = customer.companyName.toLowerCase().trim();
      const fName = customer.fullName.toLowerCase().trim();

      let matched = false;

      // 会社名の部分一致
      if (companyLower && cName && (cName.includes(companyLower) || companyLower.includes(cName))) {
        matched = true;
      }
      // 氏名の完全一致
      if (nameLower && fName && fName === nameLower) {
        matched = true;
      }

      if (matched) {
        results.push(customer);
      }
    });

    return results;
  }

  /**
   * 休眠顧客を抽出 (最終接触日からN日以上経過)
   * @param {number} [thresholdDays] - 閾値日数 (デフォルト: CONFIG値)
   * @return {Object[]} 休眠顧客配列
   */
  function getDormantCustomers(thresholdDays) {
    const threshold = thresholdDays || SFA_CONFIG.PARAMS.DORMANT_THRESHOLD_DAYS;
    const now = new Date();
    const all = getAllCustomers();

    return all.filter(customer => {
      if (!customer.lastContact) return false;

      const lastDate = new Date(customer.lastContact);
      if (isNaN(lastDate.getTime())) return false;

      const diffDays = Math.floor((now - lastDate) / (1000 * 60 * 60 * 24));
      return diffDays >= threshold;
    }).map(customer => {
      const lastDate = new Date(customer.lastContact);
      const diffDays = Math.floor((now - lastDate) / (1000 * 60 * 60 * 24));
      return { ...customer, dormantDays: diffDays };
    });
  }

  // ── Public API ──
  return {
    getSpreadsheet,
    getMainSheet,
    ensureHeaders,
    getAllCustomers,
    appendToSheet,
    updateCell,
    searchCustomers,
    getDormantCustomers,
  };

})();
