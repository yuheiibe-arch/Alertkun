/**
 * ==========================================
 * シフトチェッカー 共通ヘルパー関数群 
 * (過去1ヶ月 ＆ Jinjer有給データ連携対応版)
 * ==========================================
 */

const toYYYYMMDD = (d) => `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;

function getCheckerLocationMaster(locSs) {
  const locMap = new Map();
  const locSystemStatus = new Map();
  const locSheet = locSs.getSheetByName("拠点名");
  
  if (locSheet) {
    const lData = locSheet.getDataRange().getDisplayValues();
    const lHeaders = lData[0].map(h => String(h).replace(/\s+/g, ''));
    const idxFormal = lHeaders.indexOf("正規記載");
    const idxStatus = lHeaders.indexOf("システム反映");
    const aliasIndices = lHeaders.map((h, i) => (h.includes("表記揺れ") || h.includes("クリニックNo")) ? i : -1).filter(i => i !== -1);
    
    if (idxFormal > -1) {
      for (let i = 1; i < lData.length; i++) {
        const formalName = String(lData[i][idxFormal]).replace(/\s+/g, '');
        if (!formalName) continue;
        if (idxStatus > -1) locSystemStatus.set(formalName, String(lData[i][idxStatus]).trim() === "済");
        
        locMap.set(formalName, formalName);
        aliasIndices.forEach(idx => {
          const alias = String(lData[i][idx]).replace(/\s+/g, '');
          if (alias) locMap.set(alias, formalName);
        });
      }
    }
  }

  return {
    locSystemStatus,
    normalize: (rawName) => {
      const clean = String(rawName).replace(/[\s 【】]+/g, "");
      return locMap.has(clean) ? locMap.get(clean) : clean;
    }
  };
}

function getCheckerClosedDays(activeSs, scanStartDate, normalizeFunc) {
  const closedDataMap = new Set();
  const closedSheet = activeSs.getSheetByName("休館日");
  if (closedSheet) {
    const cData = closedSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      const dateObj = parseDateToSafeDateObj(cData[i][0]);
      if (!dateObj || dateObj < scanStartDate) continue; 
      const rawClinicName = String(cData[i][2]).trim();
      if (rawClinicName) {
        closedDataMap.add(`${toYYYYMMDD(dateObj)}_${normalizeFunc(rawClinicName)}`);
      }
    }
  }
  return closedDataMap;
}

function getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, normalizeFunc) {
  const actualShiftsMap = new Map();
  
  const load = (sheet, sourceName) => {
    if (!sheet) return;
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;
    
    const cols = getColumnIndices(data[0], [
      '医籍番号', '名前', 'クリニックNo', 'クリニック名', '診療科', '勤務種別', '勤務日', 
      '勤務開始時間', '勤務終了時間', '時給1', '時給2', '時給3', '時給4', '時給合計', '合計日給'
    ]);
    
    const totalCol = cols['時給合計'] !== -1 ? cols['時給合計'] : cols['合計日給'];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateObj = parseDateToSafeDateObj(row[cols['勤務日']]);
      // ★フィルターを過去1ヶ月（scanStartDate）に変更
      if (!dateObj || dateObj < scanStartDate) continue;

      const dateStr = toYYYYMMDD(dateObj);
      const rawClinic = cols['クリニック名'] !== -1 ? String(row[cols['クリニック名']]) : "";
      let doctor = cols['名前'] !== -1 ? String(row[cols['名前']]).replace(/\s+/g, '') : "";
      
      if (sourceName === "募集" && !doctor) doctor = "募集";
      if (!rawClinic || !doctor) continue;

      const key = sourceName === "募集" ? `${dateStr}_${doctor}_${i}` : `${dateStr}_${doctor}`;
      if (!actualShiftsMap.has(key)) {
        actualShiftsMap.set(key, { hasLeave: false, shifts: [] });
      }
      
      const record = actualShiftsMap.get(key);
      const type = cols['勤務種別'] !== -1 ? String(row[cols['勤務種別']]) : "";
      
      if (rawClinic.includes("有給") || rawClinic.includes("有休") || rawClinic.includes("欠勤") || type.includes("有給") || type.includes("欠勤")) {
        record.hasLeave = true;
      }
      
      record.shifts.push({
        sourceSheet: sourceName,
        doctorId: cols['医籍番号'] !== -1 ? String(row[cols['医籍番号']]).trim() : "",
        clinicId: cols['クリニックNo'] !== -1 ? String(row[cols['クリニックNo']]).trim() : "",
        rawClinic: rawClinic,
        normClinic: normalizeFunc(rawClinic),
        dept: cols['診療科'] !== -1 ? String(row[cols['診療科']]).trim() : "",
        type: type,
        startStr: String(row[cols['勤務開始時間']]).trim(),
        endStr: String(row[cols['勤務終了時間']]).trim(),
        startMin: parseTimeToMinutes(row[cols['勤務開始時間']]),
        endMin: parseTimeToMinutes(row[cols['勤務終了時間']]),
        wages: [
          Number(row[cols['時給1']]) || 0, Number(row[cols['時給2']]) || 0,
          Number(row[cols['時給3']]) || 0, Number(row[cols['時給4']]) || 0
        ],
        wageTotal: totalCol !== -1 ? Number(row[totalCol]) || 0 : 0
      });
    }
  };

  load(pasteSs.getSheetByName("貼付用"), "実績");
  load(shiftSs.getSheetByName("確定シフト"), "確定");
  load(shiftSs.getSheetByName("応募シフト"), "応募"); 
  load(shiftSs.getSheetByName("募集シフト"), "募集"); 
  
  return actualShiftsMap;
}

function setupSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setFrozenRows(1);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setHorizontalAlignment("left");
  return sheet;
}

function getColumnIndices(headers, targetNames) {
  const indices = {};
  const cleanHeaders = headers.map(h => String(h).replace(/\r?\n/g, '').trim());
  targetNames.forEach(name => indices[name] = cleanHeaders.findIndex(h => h.includes(name)));
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

// ★「2027年02月22日」形式にも対応するよう拡張
function parseDateToSafeDateObj(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  const cleanStr = String(dateInput).replace(/\s*（.*?）/, '').replace(/[年月]/g, '/').replace(/日/g, '').replace(/-/g, '/').trim();
  const parts = cleanStr.split('/');
  return parts.length === 3 ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)) : null;
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

/**
 * ==========================================
 * ★ Jinjer 有給データ取得エンジン
 * ==========================================
 */
function getJinjerPaidLeaveData() {
  const JINJER_SS_ID = '1NfSmzERA-ee-8ToDYSwL_PYkxvMSxRv5Y_CehfeAQcQ';
  const leavesMap = new Map();
  
  const ss = safeExecute(() => SpreadsheetApp.openById(JINJER_SS_ID), 3, "Jinjer有給シート取得");
  if (!ss) return leavesMap;
  
  const sheet = ss.getSheets()[0]; 
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return leavesMap;
  
  // 動的にヘッダー行を特定（行1〜5をスキャン）
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const rowStr = data[i].join('');
    if (rowStr.includes('対象日') && rowStr.includes('医師名')) {
      headerRowIdx = i;
      break;
    }
  }
  
  const headers = data[headerRowIdx].map(h => String(h).replace(/\s+/g, ''));
  const colDate = headers.findIndex(h => h.includes('対象日'));
  const colName = headers.findIndex(h => h.includes('医師名'));
  const colStart = headers.findIndex(h => h.includes('開始'));
  const colEnd = headers.findIndex(h => h.includes('終了'));
  const colStatus = headers.findIndex(h => h.includes('対応有無'));
  
  if (colDate === -1 || colName === -1 || colStatus === -1) return leavesMap;
  
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const status = String(data[i][colStatus]).trim();
    if (status !== "済") continue; // 「済」のみ突合対象
    
    const dObj = parseDateToSafeDateObj(data[i][colDate]);
    if (!dObj) continue;
    
    const dateStr = toYYYYMMDD(dObj);
    const doctor = String(data[i][colName]).replace(/\s+/g, '');
    if (!doctor) continue;
    
    const key = `${dateStr}_${doctor}`;
    leavesMap.set(key, {
      start: colStart !== -1 ? data[i][colStart] : "",
      end: colEnd !== -1 ? data[i][colEnd] : "",
      found: false // 監査ループ内で有休シフトと紐づいたか判定用
    });
  }
  
  return leavesMap;
}

const NewWageEngine = (function() {
  const LOC_WAGE_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  let _isInitialized = false;
  let _wageDB = {};
  let _holidayMap = {};

  function _init() {
    if (_isInitialized) return;
    try {
      const ss = SpreadsheetApp.openById(LOC_WAGE_MASTER_ID);
      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        const sName = sheet.getName();
        if (!sName.includes("時給") || sName.includes("特別")) return;
        
        _wageDB[sName] = {};
        const data = sheet.getDataRange().getDisplayValues();
        for (let i = 1; i < data.length; i++) {
          const cId = String(data[i][0]).trim();
          const rawLoc = String(data[i][1]).trim();
          const rawDept = String(data[i][2]).trim();
          if (!cId && !rawLoc) continue;

          const getNum = (val) => Number(String(val).replace(/[^\d]/g, '')) || 0;
          
          const rateData = {
            wd_am: getNum(data[i][3]), wd_pm: getNum(data[i][4]), wd_nt: getNum(data[i][5]),
            sat_am: getNum(data[i][6]), sat_pm: getNum(data[i][7]), sat_nt: getNum(data[i][8]),
            hol_am: getNum(data[i][9]), hol_pm: getNum(data[i][10]), hol_nt: getNum(data[i][11])
          };

          if (cId) _wageDB[sName][`${cId}_${rawDept}`] = rateData;
          const nLoc = rawLoc.replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").replace(/\/.*/, "").replace(/[\s ]+/g, "").trim();
          _wageDB[sName][`${nLoc}_${rawDept}`] = rateData;
        }
      });
      
      const holSs = SpreadsheetApp.openById('1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw');
      holSs.getSheets().forEach(sheet => {
        const hData = sheet.getDataRange().getValues();
        for (let i = 2; i < hData.length; i++) {
          let dObj = hData[i][0];
          if (dObj instanceof Date) {
            let dStr = Utilities.formatDate(dObj, "JST", "yyyy/MM/dd");
            let isHol = String(hData[i][3]).trim() !== "" || String(hData[i][5]).trim() !== "";
            _holidayMap[dStr] = isHol;
          }
        }
      });

      _isInitialized = true;
    } catch(e) {
      throw new Error("時給マスタ初期化失敗: " + e.message);
    }
  }

  function calculateDailyTotal(clinicId, locName, deptName, dateStr, startStr, endStr, contractText) {
    _init();

    const dObj = new Date(dateStr);
    const m = dObj.getMonth() + 1;
    const y = dObj.getFullYear();
    const fy = (m >= 4) ? y : y - 1;
    const half = (m >= 4 && m <= 9) ? "上期" : "下期";

    const targetSheetNames = [`${fy}${half}時給`, `${fy}時給`];
    let targetDB = null;
    for (let name of targetSheetNames) {
      if (_wageDB[name]) {
        targetDB = _wageDB[name];
        break;
      }
    }
    
    if (contractText && (contractText.includes("2025年度") || contractText.includes("2025年"))) {
      if (_wageDB['2025時給']) targetDB = _wageDB['2025時給'];
    }

    if (!targetDB) return { status: "ERROR", msg: "該当年度の時給シートなし" };

    const nLoc = locName.replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").replace(/\/.*/, "").replace(/[\s ]+/g, "").trim();
    let targetDept = deptName;
    if (nLoc === "亀有" || nLoc === "北葛西") {
      targetDept = (locName.includes("内科") || deptName.includes("内科")) ? "内科" : "小児科";
    }

    let rates = null;
    if (clinicId && targetDB[`${clinicId}_${targetDept}`]) {
      rates = targetDB[`${clinicId}_${targetDept}`];
    } else if (clinicId && targetDB[`${clinicId}_小児科`]) {
      rates = targetDB[`${clinicId}_小児科`];
    } else if (clinicId && targetDB[`${clinicId}_`]) {
      rates = targetDB[`${clinicId}_`];
    } else {
      rates = targetDB[`${nLoc}_${targetDept}`] || targetDB[`${nLoc}_小児科`] || targetDB[`${nLoc}_`];
    }

    if (!rates) return { status: "ERROR", msg: "マスタ未設定" };

    const day = dObj.getDay();
    const isHol = _holidayMap[dateStr] === true || day === 0;
    const prefix = isHol ? "hol" : (day === 6 ? "sat" : "wd");

    let sHour = parseInt(startStr.split(':')[0], 10);
    let eHour = parseInt(endStr.split(':')[0], 10);
    if (eHour < sHour && eHour !== 0) eHour += 24;

    let totalWage = 0;
    let maxHourly = 0;

    for (let h = sHour; h < eHour; h++) {
      if (h === 13 || h === 14) continue;

      let r = 0;
      if (h >= 9 && h < 13) r = rates[`${prefix}_am`];
      else if (h >= 15 && h < 18) r = rates[`${prefix}_pm`];
      else if (h >= 18 && h < 21) r = rates[`${prefix}_nt`];

      if (contractText && contractText.includes("円")) {
        let match = contractText.match(/(\d{1,2})[：:]\d{2}-(\d{1,2})[：:]\d{2}.*?(\d{1,2}(?:,\d{3})*)円/);
        if (match) {
          let spStart = parseInt(match[1], 10);
          let spEnd = parseInt(match[2], 10);
          if (h >= spStart && h < spEnd) r = parseInt(match[3].replace(/,/g, ''), 10);
        }
      }
      
      totalWage += r;
      if (r > maxHourly) maxHourly = r;
    }

    return { status: "SUCCESS", total: totalWage, maxHourly: maxHourly };
  }

  return { calculateDailyTotal };
})();