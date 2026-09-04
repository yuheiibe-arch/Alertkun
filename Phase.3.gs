/**
 * ==========================================
 * シフトチェッカー Phase 3: 時給・給与・有給コンプライアンス監査
 * (過去1ヶ月 ＆ Jinjer有給完全突合 搭載版)
 * ★UPDATE: 年度動的判定 ＆ 医籍番号照合 ＆ 末尾追記型対応
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
  
  const EXCLUDED_DOCTORS = ['⚠休館※医師勤務なし', '休館※医師勤務なし', '⚠ 休館※医師勤務なし'];

  const existingIds = new Set();
  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  const alertData = alertSheet.getDataRange().getValues();
  for (let i = 1; i < alertData.length; i++) if (alertData[i][11]) existingIds.add(String(alertData[i][11]));

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, uniqueId) => {
    if (!existingIds.has(uniqueId)) {
      // Phase1,2と列数を合わせるため、対応指示/メモは空で埋める
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, "", "", uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, doctor===""?"#eeeeee":null, empType===""?"#eeeeee":null, workTime===""?"#eeeeee":null, null, null, null, null]);
      existingIds.add(uniqueId);
    }
  };

  const jinjerLeavesMap = getJinjerPaidLeaveData();
  const locMaster = getCheckerLocationMaster(locSs);
  const actualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, locMaster.normalize);

  const contractPeriods = new Map();
  const empTypeCache = new Map();
  
  // ★ 年度の自動計算（Phase 2と同様）
  const getNendo = (d) => d.getMonth() < 3 ? d.getFullYear() - 1 : d.getFullYear();
  const maxForecastDate = new Date(today.getTime());
  maxForecastDate.setMonth(maxForecastDate.getMonth() + 9);
  const startNendo = getNendo(scanStartDate);
  const endNendo = getNendo(maxForecastDate);
  const targetNendos = Array.from(new Set([startNendo, endNendo]));
  
  // 契約情報（年度シート）と勤怠シートの両方をスキャン
  attSs.getSheets().forEach(sheet => {
    const sName = sheet.getName();
    
    // 契約シート（例: 常勤2026年度）
    if (sName.includes("年度") && !sName.includes("勤怠")) {
      let eType = sName.includes("常勤") && !sName.includes("非常勤") ? "常勤" : (sName.includes("定期非常勤") ? "定期非常勤" : "");
      if (!eType) return;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const headers = data[0].map(h => String(h).replace(/[\s ]+/g, ""));
      const cName = headers.indexOf("氏名") !== -1 ? headers.indexOf("氏名") : (headers.indexOf("名前") !== -1 ? headers.indexOf("名前") : headers.indexOf("医師名"));
      const cId = headers.indexOf("医籍番号"); // ★追加
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
            const d = new Date(data[r][cJoin]);
            if (!isNaN(d.getTime())) jTime = d.getTime();
          }
          if (cLeave !== -1 && data[r][cLeave]) {
            const d = new Date(data[r][cLeave]);
            if (!isNaN(d.getTime())) lTime = d.getTime();
          }
          const specialText = cSpecial !== -1 ? String(data[r][cSpecial]) : "";
          
          // IDと名前の両方をキーとして登録
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
    
    // 勤怠シート（例: 常勤勤怠2026）※ターゲット年度のみ
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
        docIds[c] = String(data[1][c]).trim(); // 2行目からID取得
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

  for (const [key, actData] of actualShiftsMap.entries()) {
    actData.shifts.forEach(shift => {
      const docClean = shift.doctorName || key.split('_')[1];
      const docId = shift.doctorId; // getCheckerActualShiftsで取得済みのID
      const dateStr = key.split('_')[0];
      const displayDate = `${dateStr}(${jpDays[new Date(dateStr).getDay()]})`;
      const shiftTimeMs = new Date(dateStr).getTime();

      if (EXCLUDED_DOCTORS.includes(docClean)) return;

      const pfx = (shift.sourceSheet === "募集" || shift.sourceSheet === "応募") ? `[${shift.sourceSheet}] ` : "";

      let empType = null;
      let specialText = "";
      
      // ★ ID優先で契約期間を引く
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
        if (isWageEntered) {
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

  // ★ アラートリストへの書き出し（末尾追記）
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
  
  Logger.log(`⏱ Phase 3 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}