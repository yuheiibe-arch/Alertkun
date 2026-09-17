/**
 * ==========================================
 * シフトチェッカー Phase 3 & 依頼手当監査
 * ★UPDATE: 過去の募集シフトの時給監査を除外 ＆ アーカイブシートからの除外リスト読み込みを修正
 * ==========================================
 */
function runShiftCheckerPhase3() {
  const startTime = Date.now();
  Logger.log("=== Phase 3: 時給・給与監査 実行開始 ===");

  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet(); 
  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA'; 
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co'; 
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs'; 
  const ATTENDANCE_SS_ID = '1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA'; 
  
  let pasteSs, shiftSs, locSs, attSs;
  try {
    pasteSs = SpreadsheetApp.openById(PASTE_MASTER_ID);
    shiftSs = SpreadsheetApp.openById(SHIFT_MASTER_ID);
    locSs = SpreadsheetApp.openById(LOC_MASTER_ID);
    attSs = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message); return;
  }

  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const errorValues = [], errorBackgrounds = [];
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  
  const EXCLUDED_CLINICS = ['院外勤務（小児科）', '嘱託医業務', '医師会業務', '【関東】バックアップシフト', '有給', '欠勤'];
  const EXCLUDED_DOCTORS = ['⚠休館※医師勤務なし', '休館※医師勤務なし', '⚠ 休館※医師勤務なし'];

  const existingIds = new Set();
  
  // ★修正2: 「アーカイブ」シートを読み込み、除外IDをセットに追加（抜け落ちていた処理）
  const archiveSheet = setupSheet(ACTIVE_SS, "アーカイブ", ALERT_HEADERS);
  const archiveData = archiveSheet.getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) {
    if (archiveData[i][11]) existingIds.add(String(archiveData[i][11]));
  }

  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  const alertData = alertSheet.getDataRange().getValues();
  for (let i = 1; i < alertData.length; i++) {
    if (alertData[i][11]) existingIds.add(String(alertData[i][11]));
  }

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, uniqueId) => {
    if (!existingIds.has(uniqueId)) {
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, "", "", uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, doctor===""?"#eeeeee":null, empType===""?"#eeeeee":null, workTime===""?"#eeeeee":null, null, null, null, null]);
      existingIds.add(uniqueId);
    }
  };

  const jinjerLeavesMap = getJinjerPaidLeaveData();
  const locMaster = getCheckerLocationMaster(locSs);
  const closedDataMap = getCheckerClosedDays(pasteSs, scanStartDate, locMaster.normalize);
  const actualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, locMaster.normalize);

  const contractPeriods = new Map();
  const empTypeCache = new Map();
  
  const getNendo = (d) => d.getMonth() < 3 ? d.getFullYear() - 1 : d.getFullYear();
  const maxForecastDate = new Date(today.getTime());
  maxForecastDate.setMonth(maxForecastDate.getMonth() + 9);
  const startNendo = getNendo(scanStartDate);
  const endNendo = getNendo(maxForecastDate);
  const targetNendos = Array.from(new Set([startNendo, endNendo]));
  
  attSs.getSheets().forEach(sheet => {
    const sName = sheet.getName();
    
    if (sName.includes("年度") && !sName.includes("勤怠")) {
      let eType = sName.includes("常勤") && !sName.includes("非常勤") ? "常勤" : (sName.includes("定期非常勤") ? "定期非常勤" : "");
      if (!eType) return;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const headers = data[0].map(h => String(h).replace(/[\s ]+/g, ""));
      const cName = headers.indexOf("氏名") !== -1 ? headers.indexOf("氏名") : (headers.indexOf("名前") !== -1 ? headers.indexOf("名前") : headers.indexOf("医師名"));
      const cId = headers.indexOf("医籍番号"); 
      const cJoin = headers.findIndex(h => h.includes("入職") || h.includes("契約日"));
      const cLeave = headers.findIndex(h => h.includes("退職") || h.includes("終了"));
      const cSpecial = headers.indexOf("特別時給の内訳");

      if (cName !== -1) {
        for (let r = 1; r < data.length; r++) {
          const dName = String(data[r][cName]).replace(/\s+/g, '');
          const dId = cId !== -1 ? String(data[r][cId]).trim() : "";
          if (!dName && !dId) continue;
          
          let jTime = 0, lTime = 4102444800000;
          if (cJoin !== -1 && data[r][cJoin]) {
            const d = parseDateToSafeDateObj(data[r][cJoin]);
            if (d) jTime = d.getTime();
          }
          if (cLeave !== -1 && data[r][cLeave]) {
            const d = parseDateToSafeDateObj(data[r][cLeave]);
            if (d) lTime = d.getTime();
          }
          const specialText = cSpecial !== -1 ? String(data[r][cSpecial]) : "";
          
          if (dId) {
            if (!contractPeriods.has(`ID_${dId}`)) contractPeriods.set(`ID_${dId}`, []);
            contractPeriods.get(`ID_${dId}`).push({ type: eType, start: jTime, end: lTime, specialText: specialText });
          }
          if (dName) {
            if (!contractPeriods.has(`NAME_${dName}`)) contractPeriods.set(`NAME_${dName}`, []);
            contractPeriods.get(`NAME_${dName}`).push({ type: eType, start: jTime, end: lTime, specialText: specialText });
          }
        }
      }
    }
    
    if (sName.includes("勤怠")) {
      const yearMatch = sName.match(/\d{4}/);
      if (yearMatch && !targetNendos.includes(parseInt(yearMatch[0], 10))) return;

      let eType = sName.includes("常勤") && !sName.includes("非常勤") ? "常勤" : (sName.includes("定期非常勤") ? "定期非常勤" : "");
      if (!eType) return;

      const data = sheet.getDataRange().getDisplayValues();
      if (data.length < 3) return;
      const docNames = [];
      const docIds = [];
      for (let c = 6; c < data[0].length; c++) {
        docNames[c] = String(data[0][c]).replace(/\s+/g, '');
        docIds[c] = String(data[1][c]).trim(); 
      }

      for (let r = 2; r < data.length; r++) {
        const dObj = parseDateToSafeDateObj(data[r][0]);
        if (!dObj) continue;
        const dStr = toYYYYMMDD(dObj);
        for (let c = 6; c < data[0].length; c++) {
          const dName = docNames[c];
          const dId = docIds[c];
          if (!dName && !dId) continue;
          
          const text = String(data[r][c]).trim();
          if (text !== "-") {
            if (dId) empTypeCache.set(`${dStr}_ID_${dId}`, eType);
            if (dName) empTypeCache.set(`${dStr}_NAME_${dName}`, eType);
          }
        }
      }
    }
  });

  const clinicDateShifts = new Map();
  for (const actData of actualShiftsMap.values()) {
    actData.shifts.forEach(shift => {
      const cKey = `${shift.dateStr}_${shift.normClinic}`;
      if (!clinicDateShifts.has(cKey)) clinicDateShifts.set(cKey, []);
      clinicDateShifts.get(cKey).push(shift);
    });
  }

  for (const [key, actData] of actualShiftsMap.entries()) {
    actData.shifts.forEach(shift => {
      const docClean = shift.doctorName || key.split('_')[1];
      const docId = shift.doctorId; 
      const dateStr = key.split('_')[0];
      const displayDate = `${dateStr}(${jpDays[new Date(dateStr).getDay()]})`;
      const shiftTimeMs = new Date(dateStr).getTime();

      if (EXCLUDED_CLINICS.includes(shift.rawClinic) || EXCLUDED_DOCTORS.includes(docClean)) return;
      
      const closedInfo = closedDataMap.get(`${dateStr}_${shift.normClinic}`) || closedDataMap.get(`${dateStr}_全拠点`);
      if (closedInfo && closedInfo.type === "closed") {
        if (shift.sourceSheet === "募集" || docClean === "募集") {
          // 未来の休館日募集シフトの場合のみ警告
          if (shiftTimeMs >= today.getTime()) {
            const maxWage = Math.max(...shift.wages);
            const isWageEntered = maxWage > 0 || shift.wageTotal > 0;
            if (isWageEntered) {
              const actualWageStr = `￥${Math.max(maxWage, shift.wageTotal).toLocaleString()}`;
              addError(`[募集] 休館日・時給消し忘れ`, displayDate, shift.rawClinic, shift.dept, "募集", "スポット", `${shift.startStr}-${shift.endStr}`, `正：￥0\n実：${actualWageStr}`, `休館日募集時給_${dateStr}_${shift.normClinic}_${shift.startStr}`);
            }
          }
        }
        return; 
      }

      // ★修正1: 過去の募集シフトは時給監査を完全にスキップ（過去の空枠エラー防止）
      if (shift.sourceSheet === "募集" || docClean === "募集") {
        if (shiftTimeMs < today.getTime()) {
          return;
        }
      }

      if (shift.sourceSheet === "募集") {
        const pStatus = shift.publishStatus || "";
        const isUnpublished = pStatus.includes("未掲載") || pStatus.includes("非公開") || pStatus.includes("非掲載");

        if (isUnpublished) {
          if (shift.startMin === 9 * 60 && shift.endMin === 21 * 60) return;

          const cShifts = clinicDateShifts.get(`${dateStr}_${shift.normClinic}`) || [];
          const is2ndDoc = cShifts.some(c => (c.sourceSheet === "確定" || c.sourceSheet === "実績" || c.sourceSheet === "応募") && c.startMin < shift.endMin && c.endMin > shift.startMin);
          
          if (is2ndDoc) {
            const maxWageCheck = Math.max(...shift.wages);
            if (maxWageCheck === 0 && shift.wageTotal === 0) {
              return; 
            }
          }
        }
      }

      const pfx = (shift.sourceSheet === "募集" || shift.sourceSheet === "応募") ? `[${shift.sourceSheet}] ` : "";

      let empType = null;
      let specialText = "";
      
      let periods = docId ? contractPeriods.get(`ID_${docId}`) : null;
      if (!periods) periods = contractPeriods.get(`NAME_${docClean}`);
      
      if (periods && periods.length > 0) {
        const validPeriod = periods.find(p => shiftTimeMs >= p.start && shiftTimeMs <= p.end);
        if (validPeriod) {
          empType = validPeriod.type;
          specialText = validPeriod.specialText;
        } else {
          empType = "スポット"; 
        }
      }
      
      if (!empType || docClean === "募集") {
        empType = (docId ? empTypeCache.get(`${dateStr}_ID_${docId}`) : null) || empTypeCache.get(`${dateStr}_NAME_${docClean}`) || "スポット";
      }

      const isKekkin = shift.rawClinic.includes("欠勤") || shift.type.includes("欠勤");
      const isYuku = shift.rawClinic.includes("有給") || shift.type.includes("有休") || shift.type.includes("有給");
      const maxWage = Math.max(...shift.wages);
      const actualWageStr = `￥${Math.max(maxWage, shift.wageTotal).toLocaleString()}`;
      const isWageEntered = maxWage > 0 || shift.wageTotal > 0;

      const jinjerKey = `${dateStr}_${docClean}`;
      if (jinjerLeavesMap.has(jinjerKey)) {
        if (isYuku) {
          jinjerLeavesMap.get(jinjerKey).found = true; 
        } else if (shift.sourceSheet !== "募集" && shift.sourceSheet !== "応募") {
          addError(`Jinjer有給未反映`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正：有休\n実：通常勤務(Jinjer済)`, `Jinjer未反映_${dateStr}_${shift.normClinic}_${docClean}`);
          jinjerLeavesMap.get(jinjerKey).found = true; 
        }
      }

      let allowedMax = 20000;
      let expectedTotalStr = "算出不可"; 
      
      if (empType !== "常勤") {
        try {
          const res = NewWageEngine.calculateDailyTotal(shift.clinicId, shift.normClinic, shift.dept, dateStr, shift.startStr, shift.endStr, specialText);
          if (res.status === "SUCCESS") {
            expectedTotalStr = `￥${Math.round(res.total).toLocaleString()}`;
            allowedMax = res.maxHourly + 2000;
          } else {
            expectedTotalStr = res.msg;
          }
        } catch(e) { 
          expectedTotalStr = `算出エラー`;
        }
      }

      if (isKekkin) {
        if (isWageEntered) {
          addError(`${pfx}欠勤・時給消し忘れ`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正：￥0\n実：${actualWageStr}`, `${shift.sourceSheet}_時給消忘れ_${dateStr}_${shift.normClinic}_${docClean}`);
        }
        return; 
      }

      if (isYuku) {
        if (empType === "常勤") {
          if (isWageEntered) addError(`${pfx}有給時給エラー(常勤)`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正：￥0\n実：${actualWageStr}`, `${shift.sourceSheet}_有給時給_${dateStr}_${shift.normClinic}_${docClean}`);
        } else {
          if (!isWageEntered) {
            addError(`${pfx}有給時給・未入力`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正：${expectedTotalStr}\n実：￥0`, `${shift.sourceSheet}_有給未入力_${dateStr}_${shift.normClinic}_${docClean}`);
          }
        }
        return; 
      }

      if (empType === "常勤") {
        if (isWageEntered && shift.sourceSheet !== "応募") {
          addError(`${pfx}常勤時給入力エラー`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正：￥0\n実：${actualWageStr}`, `${shift.sourceSheet}_常勤時給_${dateStr}_${shift.normClinic}_${docClean}`);
        }
        return;
      } else {
        if (!isWageEntered) {
          addError(`${pfx}時給・日給未入力`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正：${expectedTotalStr}\n実：￥0`, `${shift.sourceSheet}_未入力_${dateStr}_${shift.normClinic}_${docClean}`);
        }
      }

      if (maxWage > 20000 && maxWage > allowedMax) {
        addError(`${pfx}上限時給超過`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正(上限)：￥${allowedMax.toLocaleString()}\n実：￥${maxWage.toLocaleString()}`, `${shift.sourceSheet}_上限超過_${dateStr}_${shift.normClinic}_${docClean}`);
      }
    });
  }

  jinjerLeavesMap.forEach((leave, key) => {
    if (!leave.found) {
      const [dStr, doc] = key.split('_');
      const dObj = new Date(dStr);
      if (dObj >= scanStartDate) {
        const dispD = `${dStr}(${jpDays[dObj.getDay()]})`;
        addError(`Jinjer有給シフト未作成`, dispD, "不明", "不明", doc, "不明", `${leave.start}-${leave.end}`, `正：有給シフトあり\n実：シフト存在せず`, `Jinjer未登録_${dStr}_${doc}`);
      }
    }
  });

  if (errorValues.length > 0) {
    const combined = errorValues.map((val, i) => ({ val, bg: errorBackgrounds[i] }));
    combined.sort((a, b) => a.val[2].localeCompare(b.val[2])); 
    const sortedValues = combined.map(item => item.val);
    const sortedBackgrounds = combined.map(item => item.bg);

    const startRow = alertSheet.getLastRow() + 1;
    const requiredRows = startRow + sortedValues.length - 1;
    if (alertSheet.getMaxRows() < requiredRows) {
      alertSheet.insertRowsAfter(alertSheet.getMaxRows(), requiredRows - alertSheet.getMaxRows());
    }

    alertSheet.getRange(startRow, 1, sortedValues.length, ALERT_HEADERS.length).setValues(sortedValues).setBackgrounds(sortedBackgrounds).setHorizontalAlignment("left");
    alertSheet.getRange(startRow, 1, sortedValues.length, 1).insertCheckboxes(); 
    Logger.log(`✅ ${sortedValues.length}件のエラー(時給・有休監査)を末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規エラーはありませんでした。");
  }
}

/**
 * ==========================================
 * シフトチェッカー 依頼手当監査（本番用）
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