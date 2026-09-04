/**
 * ==========================================
 * シフトチェッカー Phase 2: 定期・常勤漏れ＆有給漏れ検知
 * (対応指示・メモ欄追加 ＆ 過去1ヶ月拡張 ＆ Jinjerリンク生成対応版)
 * ★UPDATE 1: 医籍番号による表記揺れ（高橋/髙橋等）完全突破対応
 * ★UPDATE 2: 【年度(4月〜翌3月)】自動判定＆最長9ヶ月先までの動的走査・スキップログ対応
 * ==========================================
 */
function runShiftCheckerPhase2() {
  Logger.log("=== Phase 2: 定期・常勤・有給漏れ検知 実行開始 ===");

  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet(); 
  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA'; 
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co'; 
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs'; 
  const ATTENDANCE_SS_ID = '1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA'; 
  const RECRUIT_MASTER_ID = '1NfSmzERA-ee-8ToDYSwL_PYkxvMSxRv5Y_CehfeAQcQ'; 
  
  let pasteSs, shiftSs, locSs, attSs, recruitSs;
  try {
    pasteSs = SpreadsheetApp.openById(PASTE_MASTER_ID);
    shiftSs = SpreadsheetApp.openById(SHIFT_MASTER_ID);
    locSs = SpreadsheetApp.openById(LOC_MASTER_ID);
    attSs = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
    recruitSs = SpreadsheetApp.openById(RECRUIT_MASTER_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message); return;
  }

  const scriptTimeZone = Session.getScriptTimeZone();
  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const errorValues = [], errorBackgrounds = [];
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  
  const archivedIds = new Set();
  const archiveData = setupSheet(ACTIVE_SS, "アーカイブ", ALERT_HEADERS).getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) if (archiveData[i][11]) archivedIds.add(String(archiveData[i][11])); 

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId) => {
    if (!archivedIds.has(uniqueId)) {
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, "", uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, doctor===""?"#eeeeee":null, empType===""?"#eeeeee":null, workTime===""?"#eeeeee":null, null, null, null, null]);
    }
  };

  const locMaster = getCheckerLocationMaster(locSs);
  const closedDataMap = getCheckerClosedDays(ACTIVE_SS, scanStartDate, locMaster.normalize, scriptTimeZone);
  
  const dailyShiftsMap = new Map(); 

  const loadShifts = (sheet) => {
    if (!sheet) return;
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;

    const cols = getColumnIndices(data[0], ['名前', '医籍番号', 'クリニック名', '勤務日', '勤務開始時間', '勤務終了時間']);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateObj = parseDateToSafeDateObj(row[cols['勤務日']]);
      if (!dateObj || dateObj < scanStartDate) continue; 

      const dateStr = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy/MM/dd");
      const rawClinic = cols['クリニック名'] !== -1 ? String(row[cols['クリニック名']]) : "";
      const doctor = cols['名前'] !== -1 ? String(row[cols['名前']]).replace(/\s+/g, '') : "";
      const docId = cols['医籍番号'] !== -1 ? String(row[cols['医籍番号']]).trim() : "";
      
      if (!rawClinic || !doctor) continue;

      if (!dailyShiftsMap.has(dateStr)) dailyShiftsMap.set(dateStr, []);
      
      dailyShiftsMap.get(dateStr).push({
        rawClinic: rawClinic,
        normClinic: locMaster.normalize(rawClinic),
        startMin: parseTimeToMinutes(row[cols['勤務開始時間']]),
        endMin: parseTimeToMinutes(row[cols['勤務終了時間']]),
        startStr: String(row[cols['勤務開始時間']]).trim(),
        endStr: String(row[cols['勤務終了時間']]).trim(),
        docClean: doctor,
        docId: docId
      });
    }
  };

  loadShifts(pasteSs.getSheetByName("貼付用"));
  loadShifts(shiftSs.getSheetByName("確定シフト"));

  const paidLeaveSheet = recruitSs.getSheetByName("有給申請");
  if (paidLeaveSheet) {
    const sheetId = paidLeaveSheet.getSheetId();
    const pData = paidLeaveSheet.getDataRange().getDisplayValues();
    const pCols = getColumnIndices(pData[0], ['対象日', '医師名', '対象開始時間', '対象終了時間', '対応有無']);
    for (let i = 1; i < pData.length; i++) {
      if (!String(pData[i][pCols['対応有無']]).includes("済")) continue;
      const dateObj = parseDateToSafeDateObj(pData[i][pCols['対象日']]);
      if (!dateObj || dateObj < scanStartDate) continue; 

      const dateStr = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy/MM/dd");
      const displayDate = `${dateStr}(${jpDays[dateObj.getDay()]})`;
      const rawDocName = String(pData[i][pCols['医師名']]).trim();
      const docClean = rawDocName.replace(/\s+/g, '');
      const sTime = String(pData[i][pCols['対象開始時間']]).trim();
      const eTime = String(pData[i][pCols['対象終了時間']]).trim();
      if (!docClean) continue;
      
      const allDailyActuals = dailyShiftsMap.get(dateStr) || [];
      const actuals = allDailyActuals.filter(a => a.docClean === docClean);

      if (!actuals.some(a => a.rawClinic.includes("有給") || a.rawClinic.includes("有休") || a.rawClinic.includes("欠勤"))) {
        const targetUrl = `https://docs.google.com/spreadsheets/d/${RECRUIT_MASTER_ID}/edit#gid=${sheetId}&range=${i+1}:${i+1}`;
        const actionMsg = `有給申請に済となっていますが、確定シフトに有給シフトがありません（または通常勤務になっています）。確認してください。\n詳細はこちら:\n${targetUrl}`;
        const errorDetail = `正：${sTime}-${eTime}\n実：無し`;

        addError("Jinjer有給未反映", displayDate, "有給申請", "", rawDocName, "", `${sTime}-${eTime}`, errorDetail, actionMsg, `有給漏れ_${dateStr}_${docClean}`);
      }
    }
  }

  const getNendo = (d) => {
    return d.getMonth() < 3 ? d.getFullYear() - 1 : d.getFullYear();
  };

  const maxForecastDate = new Date(today.getTime());
  maxForecastDate.setMonth(maxForecastDate.getMonth() + 9);
  
  const startNendo = getNendo(scanStartDate);
  const endNendo = getNendo(maxForecastDate);
  
  const targetSheetNames = [];
  for (let y = startNendo; y <= endNendo; y++) {
    targetSheetNames.push(`常勤勤怠${y}`);
    targetSheetNames.push(`定期非常勤勤怠${y}`);
  }

  targetSheetNames.forEach(sheetName => {
    const attSheet = attSs.getSheetByName(sheetName);
    
    if (!attSheet) {
      const yearMatch = sheetName.match(/\d{4}/);
      const y = yearMatch ? yearMatch[0] : "";
      Logger.log(`⚠️ シート [${sheetName}] がないため、${y}年度（${y}年4月～翌3月）の該当雇用区分のシフト確認はスキップしています。`);
      return; 
    }

    const empType = sheetName.includes("定期非常勤") ? "定期非常勤" : "常勤";
    const aData = attSheet.getDataRange().getDisplayValues();
    if (aData.length < 3) return;
    
    for (let r = 2; r < aData.length; r++) {
      const dateObj = parseDateToSafeDateObj(aData[r][0]);
      if (!dateObj || dateObj < scanStartDate) continue; 

      const dateStr = Utilities.formatDate(dateObj, scriptTimeZone, "yyyy/MM/dd");
      const displayDate = `${dateStr}(${jpDays[dateObj.getDay()]})`;

      for (let c = 6; c < aData[0].length; c++) {
        const rawDocName = String(aData[0][c]).trim();
        if (!rawDocName) continue;
        const docClean = rawDocName.replace(/\s+/g, '');
        const docId = String(aData[1][c]).trim(); 
        
        let shiftText = String(aData[r][c]).trim();
        if (!shiftText || shiftText === "休" || shiftText === "-") continue;

        const hasKekkinMark = shiftText.includes("欠勤");

        if (shiftText.includes('契約：')) {
          const match = shiftText.match(/契約：([\s\S]*?)(?:\n確定：|$)/);
          if (match) shiftText = match[1].trim(); 
        }

        const cleanLines = shiftText.split('\n').filter(line => 
          !line.startsWith('→') && 
          !line.includes('※振替') && 
          !line.includes('半日有給') && 
          !line.includes('有給') && 
          !line.includes('欠勤') && 
          !line.includes('移動依頼') && 
          !line.includes('確定：')
        );

        cleanLines.forEach(line => {
          const locMatch = line.match(/【(.*?)】/);
          if (!locMatch) return;
          
          const normExpClinic = locMaster.normalize(locMatch[1]);
          const isClosed = closedDataMap.has(`${dateStr}_${normExpClinic}`);
          const dept = normExpClinic.includes("内科") ? "内科" : "小児科";

          let expDur = 0, expectedTimeStr = "";
          const timeMatch = line.match(/(\d{1,2})[:：]?(\d{0,2})\s*[-~～]\s*(\d{1,2})[:：]?(\d{0,2})/);
          if (timeMatch) {
            const sh = timeMatch[1].padStart(2, '0'), sm = (timeMatch[2] || "00").padStart(2, '0');
            const eh = timeMatch[3].padStart(2, '0'), em = (timeMatch[4] || "00").padStart(2, '0');
            expectedTimeStr = `${sh}:${sm}-${eh}:${em}`;
            const sMin = parseInt(sh) * 60 + parseInt(sm), eMin = parseInt(eh) * 60 + parseInt(em);
            expDur = eMin >= sMin ? eMin - sMin : (eMin + 24 * 60) - sMin;
          } else {
            const rawText = line.replace(/契約：|確定：|【.*?】/g, '').trim();
            expectedTimeStr = rawText !== "" ? rawText : "(時間記載なし)";
          }
          
          if (locMaster.locSystemStatus.get(normExpClinic) === true) {
            
            const allDailyActuals = dailyShiftsMap.get(dateStr) || [];
            const actuals = allDailyActuals.filter(a => {
              if (docId && a.docId === docId) return true;
              return a.docClean === docClean; 
            });

            let isExcused = false, isReflected = false, actualDurStr = "無し";
            
            if (actuals.length > 0) {
              if (actuals.some(a => a.rawClinic.includes("有給") || a.rawClinic.includes("有休") || a.rawClinic.includes("欠勤"))) {
                isExcused = true;
              } else if (actuals.find(a => Math.abs(expDur - (!isNaN(a.startMin) && !isNaN(a.endMin) ? (a.endMin >= a.startMin ? a.endMin - a.startMin : (a.endMin + 24 * 60) - a.startMin) : 0)) <= 60)) {
                isExcused = true;
              }
              const matched = actuals.find(a => a.normClinic === normExpClinic);
              if (matched) {
                isReflected = true; actualDurStr = `${matched.startStr}-${matched.endStr}`;
              }
            }
            
            if (!isReflected && !isExcused) {
              let itemType = hasKekkinMark ? "欠勤作成忘れ" : "シフト(欠勤)作成忘れ";
              if (!hasKekkinMark && isClosed) itemType = "休館日(振替/欠勤漏れ)";

              const actionMsg = `契約上の勤務日ですが、シフトが登録されていません。\n勤務なら勤務作成\n欠勤なら欠勤シフトを作成してください。\n\n通常勤務：${expectedTimeStr}\n現在の登録：${actualDurStr}`;
              const errorDetail = `正：${expectedTimeStr}\n実：${actualDurStr}`;

              addError(itemType, displayDate, normExpClinic, dept, rawDocName, empType, expectedTimeStr, errorDetail, actionMsg, `漏れ_${dateStr}_${normExpClinic}_${docClean}`);
            }
          }
        });
      }
    }
  });

  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);

  if (errorValues.length > 0) {
    const combined = errorValues.map((val, i) => ({ val, bg: errorBackgrounds[i] }));
    combined.sort((a, b) => a.val[11].localeCompare(b.val[11]));
    
    const sortedValues = combined.map(item => item.val);
    const sortedBackgrounds = combined.map(item => item.bg);

    const startRow = alertSheet.getLastRow() + 1;
    const requiredRows = startRow + sortedValues.length - 1;
    if (alertSheet.getMaxRows() < requiredRows) {
      alertSheet.insertRowsAfter(alertSheet.getMaxRows(), requiredRows - alertSheet.getMaxRows());
    }

    alertSheet.getRange(startRow, 1, sortedValues.length, ALERT_HEADERS.length).setValues(sortedValues).setBackgrounds(sortedBackgrounds).setHorizontalAlignment("left");
    alertSheet.getRange(startRow, 1, sortedValues.length, 1).insertCheckboxes(); 
    Logger.log(`✅ ${sortedValues.length}件のエラーを末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規エラーはありませんでした。");
  }
}