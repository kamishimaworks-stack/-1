/**
 * ============================================================
 *  SheetHelper.js - スプレッドシート操作ユーティリティ
 * ============================================================
 *  全モジュール + フロントエンドから利用される
 *  CRUD / 検索 / 統計 / バルク操作を提供。
 * ============================================================
 */

const SheetHelper = (() => {

  // ──────────── 基本操作 ────────────

  function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getMainSheet() {
    return _getOrCreateSheet(SFA_CONFIG.SHEETS.MAIN);
  }

  function _getOrCreateSheet(sheetName) {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      console.log('[SheetHelper] シート「' + sheetName + '」を新規作成しました');
    }
    return sheet;
  }

  function ensureHeaders(sheetName, headers) {
    const sheet = _getOrCreateSheet(sheetName);
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const needsUpdate = headers.some((h, i) => firstRow[i] !== h);
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }

  // ──────────── READ: 全顧客取得 ────────────

  function getAllCustomers() {
    const sheet = getMainSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const COL = SFA_CONFIG.COL;

    return data.map((row, index) => ({
      rowIndex:       index + 2,
      registeredDate: row[COL.REGISTERED_DATE] ? new Date(row[COL.REGISTERED_DATE]).toISOString() : '',
      companyName:    row[COL.COMPANY_NAME] || '',
      fullName:       row[COL.FULL_NAME] || '',
      title:          row[COL.TITLE] || '',
      email:          row[COL.EMAIL] || '',
      phone:          row[COL.PHONE] || '',
      address:        row[COL.ADDRESS] || '',
      website:        row[COL.WEBSITE] || '',
      lastContact:    row[COL.LAST_CONTACT] ? new Date(row[COL.LAST_CONTACT]).toISOString() : '',
      staffName:      row[COL.STAFF_NAME] || '',
      imageUrl:       row[COL.IMAGE_URL] || '',
      xUrl:           row[COL.X_URL] || '',
      facebookUrl:    row[COL.FACEBOOK_URL] || '',
      instagramUrl:   row[COL.INSTAGRAM_URL] || '',
      youtubeUrl:     row[COL.YOUTUBE_URL] || '',
      tiktokUrl:      row[COL.TIKTOK_URL] || '',
      companySite:    row[COL.COMPANY_SITE] || '',
      industry:       row[COL.INDUSTRY] || '',
      trends:         row[COL.TRENDS] || '',
      challenges:     row[COL.CHALLENGES] || '',
      similar:        row[COL.SIMILAR] || '',
      dupAlert:       row[COL.DUP_ALERT] || '',
    }));
  }

  // ──────────── READ: 休眠顧客メール下書き取得 ────────────

  function getDormantDrafts() {
    const sheet = _getOrCreateSheet(SFA_CONFIG.SHEETS.DORMANT_DRAFTS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    return data.map((row, i) => ({
      rowIndex:     i + 2,
      generatedDate: row[0] ? new Date(row[0]).toISOString() : '',
      companyName:  row[1] || '',
      fullName:     row[2] || '',
      email:        row[3] || '',
      lastContact:  row[4] ? new Date(row[4]).toISOString() : '',
      dormantDays:  row[5] || 0,
      news:         row[6] || '',
      subject:      row[7] || '',
      body:         row[8] || '',
      status:       row[9] || '',
    }));
  }

  // ──────────── READ: 類似企業取得 ────────────

  function getSimilarCompanies() {
    const sheet = _getOrCreateSheet(SFA_CONFIG.SHEETS.SIMILAR_COMPANIES);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    return data.map((row, i) => ({
      rowIndex:      i + 2,
      baseCompany:   row[0] || '',
      similarName:   row[1] || '',
      industry:      row[2] || '',
      reason:        row[3] || '',
      priority:      row[4] || '',
      estimatedUrl:  row[5] || '',
      generatedDate: row[6] ? new Date(row[6]).toISOString() : '',
    }));
  }

  // ──────────── READ: 通知ログ取得 ────────────

  function getNotificationLogs() {
    const sheet = _getOrCreateSheet(SFA_CONFIG.SHEETS.NOTIFICATION_LOG);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    return data.map((row, i) => ({
      rowIndex:  i + 2,
      datetime:  row[0] ? new Date(row[0]).toISOString() : '',
      type:      row[1] || '',
      company:   row[2] || '',
      name:      row[3] || '',
      message:   row[4] || '',
      target:    row[5] || '',
    })).reverse(); // 最新順
  }

  // ──────────── UPDATE: 顧客データ更新 ────────────

  function updateCustomer(rowIndex, updates) {
    const sheet = getMainSheet();
    const COL = SFA_CONFIG.COL;
    const fieldMap = {
      companyName: COL.COMPANY_NAME,
      fullName:    COL.FULL_NAME,
      title:       COL.TITLE,
      email:       COL.EMAIL,
      phone:       COL.PHONE,
      address:     COL.ADDRESS,
      website:     COL.WEBSITE,
      lastContact: COL.LAST_CONTACT,
      staffName:   COL.STAFF_NAME,
      industry:    COL.INDUSTRY,
    };

    Object.keys(updates).forEach(key => {
      if (fieldMap[key] !== undefined) {
        const col = fieldMap[key] + 1; // 1-based
        let val = updates[key];
        if (key === 'lastContact' && val) val = new Date(val);
        sheet.getRange(rowIndex, col).setValue(val);
      }
    });

    return { success: true };
  }

  // ──────────── DELETE: 顧客削除 ────────────

  function deleteCustomer(rowIndex) {
    const sheet = getMainSheet();
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error('無効な行番号: ' + rowIndex);
    }
    sheet.deleteRow(rowIndex);
    return { success: true };
  }

  // ──────────── STATS: ダッシュボード統計 ────────────

  function getDashboardStats() {
    const all = getAllCustomers();
    const now = new Date();
    const threshold = SFA_CONFIG.PARAMS.DORMANT_THRESHOLD_DAYS;

    // 基本KPI
    const totalCustomers = all.length;
    let dormantCount = 0;
    let activeCount = 0;
    let noContactCount = 0;
    const monthlyMap = {};   // 月別登録数
    const industryMap = {};  // 業種別件数
    const staffMap = {};     // 担当者別件数

    all.forEach(c => {
      // 休眠判定
      if (c.lastContact) {
        const last = new Date(c.lastContact);
        const diff = Math.floor((now - last) / (1000 * 60 * 60 * 24));
        if (diff >= threshold) { dormantCount++; } else { activeCount++; }
      } else {
        noContactCount++;
      }

      // 月別集計
      if (c.registeredDate) {
        const d = new Date(c.registeredDate);
        const ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
        monthlyMap[ym] = (monthlyMap[ym] || 0) + 1;
      }

      // 業種別
      if (c.industry) {
        industryMap[c.industry] = (industryMap[c.industry] || 0) + 1;
      }

      // 担当者別
      if (c.staffName) {
        staffMap[c.staffName] = (staffMap[c.staffName] || 0) + 1;
      }
    });

    // 月別を配列化（直近12ヶ月）
    const monthly = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      monthly.push({ month: ym, count: monthlyMap[ym] || 0 });
    }

    // 業種を件数降順上位10
    const industries = Object.entries(industryMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([name, count]) => ({ name, count }));

    // 担当者を件数降順
    const staffRanking = Object.entries(staffMap)
      .sort((a, b) => b[1] - a[1])
      .map(([name, count]) => ({ name, count }));

    return {
      totalCustomers,
      dormantCount,
      activeCount,
      noContactCount,
      monthly,
      industries,
      staffRanking,
    };
  }

  // ──────────── 既存ユーティリティ (互換) ────────────

  function appendToSheet(sheetName, rowData) {
    const sheet = _getOrCreateSheet(sheetName);
    sheet.appendRow(rowData);
  }

  function updateCell(row, col, value) {
    const sheet = getMainSheet();
    sheet.getRange(row, col + 1).setValue(value);
  }

  function searchCustomers(companyName, fullName) {
    const all = getAllCustomers();
    const companyLower = (companyName || '').toLowerCase().trim();
    const nameLower    = (fullName || '').toLowerCase().trim();
    return all.filter(c => {
      const cn = c.companyName.toLowerCase().trim();
      const fn = c.fullName.toLowerCase().trim();
      let matched = false;
      if (companyLower && cn && (cn.includes(companyLower) || companyLower.includes(cn))) matched = true;
      if (nameLower && fn && fn === nameLower) matched = true;
      return matched;
    });
  }

  function getDormantCustomers(thresholdDays) {
    const threshold = thresholdDays || SFA_CONFIG.PARAMS.DORMANT_THRESHOLD_DAYS;
    const now = new Date();
    return getAllCustomers().filter(c => {
      if (!c.lastContact) return false;
      const last = new Date(c.lastContact);
      if (isNaN(last.getTime())) return false;
      return Math.floor((now - last) / (1000 * 60 * 60 * 24)) >= threshold;
    }).map(c => {
      const diff = Math.floor((now - new Date(c.lastContact)) / (1000 * 60 * 60 * 24));
      return { ...c, dormantDays: diff };
    });
  }

  // ── Public API ──
  return {
    getSpreadsheet,
    getMainSheet,
    ensureHeaders,
    getAllCustomers,
    getDormantDrafts,
    getSimilarCompanies,
    getNotificationLogs,
    getDashboardStats,
    updateCustomer,
    deleteCustomer,
    appendToSheet,
    updateCell,
    searchCustomers,
    getDormantCustomers,
  };

})();
