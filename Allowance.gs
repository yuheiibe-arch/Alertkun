/**
 * ==========================================
 * シフトチェッカー 依頼手当監査（独立ファイル化）
 * ==========================================
 */
function runShiftCheckerAllowance() {
  Logger.log("=== 依頼手当 監査 実行開始 ===");
  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet();
  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA'; 
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co'; 
  const ALLOWANCE_DOCS_ID = '1Cw2iHRgIJdUf4aeCaE04DhGVBMPpyQdhpiC8U_OdVtI'; 
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  
  let pasteSs, shiftSs, allowSs, locSs;
  try {
    pasteSs = SpreadsheetApp.openById(PASTE_MASTER_ID);
    shiftSs = SpreadsheetApp.openById(SHIFT_MASTER_ID);
    allowSs = SpreadsheetApp.openById(ALLOWANCE_DOCS_ID);
    locSs = SpreadsheetApp.openById(LOC_MASTER_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message); return;
  }

  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  
  const existingIds = new Set();
  const archiveData = setupSheet(ACTIVE_SS, "アーカイブ", ALERT_HEADERS).getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) if (archiveData[i][11]) existingIds.add(String(archiveData[i][11]));
  
  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  const alertData = alertSheet.getDataRange().getValues();
  for (let i = 1; i < alertData.length; i++) if (alertData[i][11]) existingIds.add(String(alertData[i][11]));

  const errorValues = [], errorBackgrounds = [];
  const addError = (type, displayDate, clinic, doctor, workTime, errorDetail, uniqueId) => {
    if (!existingIds.has(uniqueId)) {
      const actionMsg = `依頼手当の金額が規定と一致しません（または未入力です）。確認してください。`;
      errorValues.push([false, type, displayDate, clinic, "複数", doctor, "常勤/非常勤", workTime, errorDetail, actionMsg, "", uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, null, null, null, null, null, null, null]);
      existingIds.add(uniqueId);
    }
  };

  const locMap = new Map();
  const locSheet = locSs.getSheetByName("拠点名");
  if (locSheet) {
    const lData = locSheet.getDataRange().getDisplayValues();
    const lHeaders = lData[0].map(h => String(h).replace(/[\s ]+/g, ''));
    const idxFormal = lHeaders.indexOf("正規記載");
    const idxNo = lHeaders.indexOf("クリニックNo");
    const aliasIndices = lHeaders.map((h, i) => h.includes("表記揺れ") ? i : -1).filter(i => i !== -1);
    
    if (idxFormal > -1) {
      for (let i = 1; i < lData.length; i++) {
        const formalName = String(lData[i][idxFormal]).replace(/[\s ]+/g, '');
        const clinicNo = idxNo > -1 ? String(lData[i][idxNo]).trim() : "";
        if (!formalName && !clinicNo) continue; 
        if (formalName.toUpperCase().includes("MQC")) continue; 
        if (formalName) {
          locMap.set(formalName, formalName);
          aliasIndices.forEach(idx => {
            const alias = String(lData[i][idx]).replace(/[\s ]+/g, '');
            if (alias) locMap.set(alias, formalName);
          });
        }
      }
    }
  }

  const normalizeClinic = (rawName) => {
    const clean = String(rawName).replace(/[\s 【】]+/g, "");
    return locMap.has(clean) ? locMap.get(clean) : null;
  };

  const targetDoctors = new Set();
  const allowSheet = allowSs.getSheetByName("依頼手当医師");
  if (allowSheet) {
    const data = allowSheet.getDataRange().getDisplayValues();
    let nameCol = 0, idCol = 1;
    if(data.length > 0) {
       const headers = data[0].map(h => String(h).replace(/[\s ]+/g, ''));
       nameCol = Math.max(0, headers.findIndex(h => h.includes("氏名") || h.includes("名前")));
       idCol = Math.max(1, headers.findIndex(h => h.includes("医籍番号")));
    }
    for (let i = 1; i < data.length; i++) {
      const docName = String(data[i][nameCol]).replace(/[\s ]+/g, ''); 
      const docId = String(data[i][idCol]).trim();                  
      if (docId) targetDoctors.add(`ID_${docId}`);
      if (docName) targetDoctors.add(`NAME_${docName}`);
    }
  }

  const parseTime = (timeStr) => {
    const parts = String(timeStr).trim().split(':');
    return parts.length >= 2 ? parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10) : NaN;
  };

  const getExpectedAllowance = (startStr, endStr, clinicName) => {
    const sMin = parseTime(startStr);
    const eMin = parseTime(endStr);
    if (isNaN(sMin) || isNaN(eMin)) return 0;
    if (clinicName.includes("北葛西") && startStr === "17:00" && endStr === "20:00") return 4500;
    if (sMin <= 9*60 && eMin >= 21*60) return 15000;
    let total = 0;
    if (sMin < 13*60 && eMin > 9*60) total += 6000;
    if (sMin < 18*60 && eMin > 15*60) total += 4500;
    if (sMin < 21*60 && eMin > 18*60) total += 4500;
    return total;
  };

  const checkStartDate = new Date("2026/10/01");
  const dailySummary = new Map();

  const sourcesToScan = [
    { sheet: pasteSs.getSheetByName("貼付用"), sourceName: "実績" },
    { sheet: shiftSs.getSheetByName("確定シフト"), sourceName: "確定" },
    { sheet: shiftSs.getSheetByName("応募シフト"), sourceName: "応募" }
  ];

  const today = new Date();
  const currentMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);

  sourcesToScan.forEach(src => {
    const sheet = src.sheet;
    const sourceName = src.sourceName;
    if (!sheet) return;
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h).replace(/\r?\n/g, '').trim());
    
    const cDate = headers.indexOf("勤務日");
    const cName = headers.indexOf("名前");
    const cId = headers.indexOf("医籍番号");
    const cClinic = headers.indexOf("クリニック名");
    const cType = headers.indexOf("勤務種別");
    const cStart = headers.indexOf("勤務開始時間");
    const cEnd = headers.indexOf("勤務終了時間");
    
    const addItems = [1, 2, 3, 4, 5].map(n => headers.indexOf(`追加支給項目${n}`));
    const addValues = [1, 2, 3, 4, 5].map(n => headers.indexOf(`追加支給額${n}`));
    const comCols = [1, 2, 3, 4, 5].map(n => headers.indexOf(`スタッフコメント${n}`)).filter(idx => idx !== -1);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dStr = row[cDate];
      if (!dStr) continue;
      
      const dObj = new Date(dStr.replace(/[\s（(].*$/, '').replace(/[年月]/g, '/').replace(/日/g, ''));
      if (isNaN(dObj.getTime()) || dObj < checkStartDate) continue;

      if (sourceName === "実績") {
        if (dObj < currentMonthStart || dObj >= thresholdDate) continue;
      } else if (sourceName === "確定") {
        if (dObj >= currentMonthStart && dObj < thresholdDate) continue;
      }

      const rawClinic = String(row[cClinic]);
      const type = cType !== -1 ? String(row[cType]) : "";
      
      if (rawClinic.includes("欠勤") || type.includes("欠勤") || rawClinic.includes("有給") || rawClinic.includes("有休") || type.includes("有給") || type.includes("有休")) continue;

      const normClinic = normalizeClinic(rawClinic);
      if (!normClinic) continue;

      const docClean = String(row[cName]).replace(/[\s ]+/g, '');
      const docId = String(row[cId]).trim();
      
      const isTarget = (docId && targetDoctors.has(`ID_${docId}`)) || (docClean && targetDoctors.has(`NAME_${docClean}`));
      if (!isTarget) continue;

      const comments = comCols.map(idx => String(row[idx])).join(" ");
      const isAdditional = comments.includes("所定休出") || comments.includes("所定外休出") || comments.includes("追加");
      if (!isAdditional) continue;

      let actualAllowance = 0;
      for (let j = 0; j < 5; j++) {
        if (addItems[j] !== -1 && addValues[j] !== -1) {
          const itemName = String(row[addItems[j]]).trim();
          if (itemName.includes("依頼手当") && !itemName.includes("事務局")) {
            actualAllowance += Number(String(row[addValues[j]]).replace(/[^\d]/g, '')) || 0;
          }
        }
      }

      const sStr = String(row[cStart]);
      const eStr = String(row[cEnd]);
      const expectedAllowance = getExpectedAllowance(sStr, eStr, normClinic);

      const dailyKey = `${Utilities.formatDate(dObj, "JST", "yyyy/MM/dd")}_${docClean}`;
      if (!dailySummary.has(dailyKey)) {
        dailySummary.set(dailyKey, {
          dateObj: dObj, dateStr: Utilities.formatDate(dObj, "JST", "yyyy/MM/dd"),
          docClean: docClean, clinics: new Set(), timeRanges: [],
          expectedTotal: 0, actualTotal: 0
        });
      }
      const record = dailySummary.get(dailyKey);
      record.clinics.add(normClinic);
      record.timeRanges.push(`${sStr}-${eStr}`);
      record.expectedTotal += expectedAllowance;
      record.actualTotal += actualAllowance;
    }
  });

  for (const [key, record] of dailySummary.entries()) {
    if (record.expectedTotal > 0 && record.actualTotal !== record.expectedTotal) {
      const clinicStr = Array.from(record.clinics).join(', ');
      const timeStr = record.timeRanges.join(' / ');
      const displayDate = `${record.dateStr}(${jpDays[record.dateObj.getDay()]})`;
      const errorDetail = `正：¥${record.expectedTotal.toLocaleString()}\n実：¥${record.actualTotal.toLocaleString()}`;
      
      addError("依頼手当エラー", displayDate, clinicStr, record.docClean, timeStr, errorDetail, `依頼手当_${record.dateStr}_${record.docClean}`);
    }
  }

  if (errorValues.length > 0) {
    const startRow = alertSheet.getLastRow() + 1;
    const requiredRows = startRow + errorValues.length - 1;
    if (alertSheet.getMaxRows() < requiredRows) {
      alertSheet.insertRowsAfter(alertSheet.getMaxRows(), requiredRows - alertSheet.getMaxRows());
    }
    alertSheet.getRange(startRow, 1, errorValues.length, ALERT_HEADERS.length).setValues(errorValues).setBackgrounds(errorBackgrounds).setHorizontalAlignment("left");
    alertSheet.getRange(startRow, 1, errorValues.length, 1).insertCheckboxes(); 
    Logger.log(`✅ ${errorValues.length}件の依頼手当エラーを末尾に追記しました。`);
  } else {
    Logger.log("✅ 依頼手当の新規エラーはありませんでした。");
  }
}