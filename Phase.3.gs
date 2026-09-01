/**
 * ==========================================
 * シフトチェッカー Phase 3: 時給・給与・有給コンプライアンス監査
 * (過去1ヶ月 ＆ Jinjer有給完全突合 搭載版)
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
  
  // ★ 監査範囲を「今日」から「過去30日(約1ヶ月)」へ拡張
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "ユニークキー"];
  
  const EXCLUDED_DOCTORS = ['⚠休館※医師勤務なし', '休館※医師勤務なし', '⚠ 休館※医師勤務なし'];

  const existingIds = new Set();
  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  if (alertSheet.getLastRow() > 1) {
    alertSheet.getRange(2, 1, alertSheet.getLastRow() - 1, alertSheet.getMaxColumns())
              .clearContent().removeCheckboxes().setBackground(null);
  }

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, uniqueId) => {
    if (!existingIds.has(uniqueId)) {
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, doctor===""?"#eeeeee":null, empType===""?"#eeeeee":null, workTime===""?"#eeeeee":null, null, null]);
      existingIds.add(uniqueId);
    }
  };

  // ★ Jinjer 有給データのロード
  const jinjerLeavesMap = getJinjerPaidLeaveData();
  const locMaster = getCheckerLocationMaster(locSs);
  // ★ scanStartDateを起点にシフトを抽出
  const actualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, locMaster.normalize);

  const contractPeriods = new Map();
  const empTypeCache = new Map();
  
  attSs.getSheets().forEach(sheet => {
    const sName = sheet.getName();
    
    if (sName.includes("年度") && !sName.includes("勤怠")) {
      let eType = sName.includes("常勤") && !sName.includes("非常勤") ? "常勤" : (sName.includes("定期非常勤") ? "定期非常勤" : "");
      if (!eType) return;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const headers = data[0].map(h => String(h).replace(/[\s ]+/g, ""));
      const cName = headers.indexOf("氏名") !== -1 ? headers.indexOf("氏名") : (headers.indexOf("名前") !== -1 ? headers.indexOf("名前") : headers.indexOf("医師名"));
      const cJoin = headers.findIndex(h => h.includes("入職") || h.includes("契約日"));
      const cLeave = headers.findIndex(h => h.includes("退職") || h.includes("終了"));
      const cSpecial = headers.indexOf("特別時給の内訳");

      if (cName !== -1) {
        for (let r = 1; r < data.length; r++) {
          const dName = String(data[r][cName]).replace(/\s+/g, '');
          if (!dName) continue;
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
          
          if (!contractPeriods.has(dName)) contractPeriods.set(dName, []);
          contractPeriods.get(dName).push({ type: eType, start: jTime, end: lTime, specialText: specialText });
        }
      }
    }
    
    if (sName.includes("勤怠")) {
      let eType = sName.includes("常勤") && !sName.includes("非常勤") ? "常勤" : (sName.includes("定期非常勤") ? "定期非常勤" : "");
      if (!eType) return;

      const data = sheet.getDataRange().getDisplayValues();
      if (data.length < 3) return;
      const docNames = [];
      for (let c = 6; c < data[0].length; c++) docNames[c] = String(data[0][c]).replace(/\s+/g, '');

      for (let r = 2; r < data.length; r++) {
        const dObj = parseDateToSafeDateObj(data[r][0]);
        if (!dObj) continue;
        const dStr = toYYYYMMDD(dObj);
        for (let c = 6; c < data[0].length; c++) {
          const dName = docNames[c];
          if (!dName) continue;
          const text = String(data[r][c]).trim();
          if (text !== "-") empTypeCache.set(`${dStr}_${dName}`, eType);
        }
      }
    }
  });

  for (const [key, actData] of actualShiftsMap.entries()) {
    actData.shifts.forEach(shift => {
      const docClean = key.split('_')[1];
      const dateStr = key.split('_')[0];
      const displayDate = `${dateStr}(${jpDays[new Date(dateStr).getDay()]})`;
      const shiftTimeMs = new Date(dateStr).getTime();

      if (EXCLUDED_DOCTORS.includes(docClean) || EXCLUDED_DOCTORS.includes(shift.doctorName)) return;

      const pfx = (shift.sourceSheet === "募集" || shift.sourceSheet === "応募") ? `[${shift.sourceSheet}] ` : "";

      let empType = null;
      let specialText = "";
      const periods = contractPeriods.get(docClean);
      
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
        empType = empTypeCache.get(`${dateStr}_${docClean}`) || "スポット";
      }

      const isKekkin = shift.rawClinic.includes("欠勤") || shift.type.includes("欠勤");
      const isYuku = shift.rawClinic.includes("有給") || shift.type.includes("有休") || shift.type.includes("有給");
      const maxWage = Math.max(...shift.wages);
      const actualWageStr = `￥${Math.max(maxWage, shift.wageTotal).toLocaleString()}`;
      const isWageEntered = maxWage > 0 || shift.wageTotal > 0;

      // ★ Jinjer 有給突合チェック
      const jinjerKey = `${dateStr}_${docClean}`;
      if (jinjerLeavesMap.has(jinjerKey)) {
        if (isYuku) {
          jinjerLeavesMap.get(jinjerKey).found = true; // 正常に有給シフトとして反映済み
        } else if (shift.sourceSheet !== "募集" && shift.sourceSheet !== "応募") {
          // 「済」なのに通常シフトとして残っている（募集・応募シートは除外）
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

  // ★ Jinjer側で「済」となっているのに、シフト自体が一つも存在しないケースの拾い上げ
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

    const requiredRows = sortedValues.length + 1;
    if (alertSheet.getMaxRows() < requiredRows) {
      alertSheet.insertRowsAfter(alertSheet.getMaxRows(), requiredRows - alertSheet.getMaxRows());
    }

    alertSheet.getRange(2, 1, sortedValues.length, ALERT_HEADERS.length).setValues(sortedValues).setBackgrounds(sortedBackgrounds).setHorizontalAlignment("left");
    alertSheet.getRange(2, 1, sortedValues.length, 1).insertCheckboxes(); 
    Logger.log(`✅ ${sortedValues.length}件のエラー(時給・有休監査)を日付順に出力しました。`);
  } else {
    Logger.log("✅ 新規エラーはありませんでした。");
  }
  
  Logger.log(`⏱ Phase 3 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}