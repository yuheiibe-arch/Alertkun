/**
 * ============================================================================
 * 【0-A_共通_キャッシュ・汎用】
 * ・同一実行内のスプレッドシートOpen＆データ読込オンメモリキャッシュ
 * ・通信エラー時の自動リトライ（safeExecute / safeOpenById / safeOpenByUrl）
 * ・アラートリスト／アーカイブ初期化・安全な行追記ヘルパー
 * ・日付・時間・ヘッダー列検索の汎用ユーティリティ
 * ============================================================================
 */

function getAlertkunRuntimeCache_() {
  if (!globalThis._ALERTKUN_RUNTIME_CACHE_) {
    globalThis._ALERTKUN_RUNTIME_CACHE_ = {
      ssById: new Map(),
      displayValues: new Map(),
      rawValues: new Map(),
      locMasterBySsId: new Map(),
      twoDocByKey: new Map()
    };
  }
  return globalThis._ALERTKUN_RUNTIME_CACHE_;
}

function clearCheckerRuntimeCache() {
  const cache = getAlertkunRuntimeCache_();
  cache.ssById.clear();
  cache.displayValues.clear();
  cache.rawValues.clear();
  cache.locMasterBySsId.clear();
  cache.twoDocByKey.clear();
}

function getCachedSs(id) {
  const cache = getAlertkunRuntimeCache_();
  if (cache.ssById.has(id)) {
    return cache.ssById.get(id);
  }
  const ss = safeOpenById(id);
  cache.ssById.set(id, ss);
  return ss;
}

function getCachedDisplayValues(sheet) {
  if (!sheet) return [];
  const cache = getAlertkunRuntimeCache_();
  const key = `${sheet.getParent().getId()}_${sheet.getSheetId()}`;
  if (cache.displayValues.has(key)) {
    return cache.displayValues.get(key);
  }
  const data = safeExecute(() => sheet.getDataRange().getDisplayValues(), 3, `表示値取得(${sheet.getName()})`);
  cache.displayValues.set(key, data);
  return data;
}

function getCachedValues(sheet) {
  if (!sheet) return [];
  const cache = getAlertkunRuntimeCache_();
  const key = `${sheet.getParent().getId()}_${sheet.getSheetId()}`;
  if (cache.rawValues.has(key)) {
    return cache.rawValues.get(key);
  }
  const data = safeExecute(() => sheet.getDataRange().getValues(), 3, `値取得(${sheet.getName()})`);
  cache.rawValues.set(key, data);
  return data;
}

function safeOpenById(id) {
  return safeExecute(() => SpreadsheetApp.openById(id), 3, `IDからのシート取得(${id})`);
}

function safeOpenByUrl(url) {
  return safeExecute(() => SpreadsheetApp.openByUrl(url), 3, `URLの取得(${url})`);
}

function safeExecute(action, maxRetry = 3, actionName = "スプレッドシート操作") {
  for (let i = 0; i < maxRetry; i++) {
    try {
      return action();
    } catch (e) {
      if (i === maxRetry - 1) {
        Logger.log(`[致命的エラー] ${actionName}に${maxRetry}回失敗しました: ${e.message}`);
        throw e;
      }
      Logger.log(`[通信エラー] ${actionName}に失敗。${i + 1}回目のリトライを行います...`);
      Utilities.sleep(2000 + (i * 2000));
    }
  }
}

function setupSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  let isNew = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setFrozenRows(1);
    isNew = true;
  }
  const currentFirstHeader = String(sheet.getRange(1, 1).getValue() || "").trim();
  if (isNew || currentFirstHeader !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold")
      .setHorizontalAlignment("left");
  }
  return sheet;
}

function appendAlertRowsSafe(alertSheet, headersLength, rowsValues, rowsBackgrounds, richTextsCol10 = null) {
  if (!rowsValues || rowsValues.length === 0) return;
  const startRow = alertSheet.getLastRow() + 1;
  const requiredRows = startRow + rowsValues.length - 1;
  if (alertSheet.getMaxRows() < requiredRows) {
    alertSheet.insertRowsAfter(alertSheet.getMaxRows(), requiredRows - alertSheet.getMaxRows());
  }

  alertSheet.getRange(startRow, 1, rowsValues.length, headersLength)
    .setValues(rowsValues)
    .setBackgrounds(rowsBackgrounds)
    .setHorizontalAlignment("left");
  alertSheet.getRange(startRow, 1, rowsValues.length, 1).insertCheckboxes();

  if (richTextsCol10 && richTextsCol10.length === rowsValues.length) {
    alertSheet.getRange(startRow, 10, richTextsCol10.length, 1).setRichTextValues(richTextsCol10);
  }
}

function toYYYYMMDD(d) {
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

function getColumnIndices(headers, targetNames) {
  const indices = {};
  const cleanHeaders = headers.map(h => String(h).replace(/\r?\n/g, '').trim());
  targetNames.forEach(name => {
    indices[name] = cleanHeaders.findIndex(h => h.includes(name));
  });
  return indices;
}

function parseTimeToMinutes(timeInput) {
  if (!timeInput) return NaN;
  const parts = String(timeInput).trim().split(':');
  return parts.length >= 2 ? parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10) : NaN;
}

function formatMinutesToHHMM(mins) {
  if (isNaN(mins) || mins < 0) return "不明";
  return `${String(Math.floor(mins / 60)).padStart(2, '0')}:${String(mins % 60).padStart(2, '0')}`;
}

function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) {
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }
  const cleanStr = String(dateInput)
    .replace(/\s*[（(].*?[）)]/, '')
    .replace(/[年月]/g, '/')
    .replace(/日/g, '')
    .replace(/-/g, '/')
    .trim();
  const parts = cleanStr.split('/');
  return parts.length === 3
    ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10))
    : null;
}