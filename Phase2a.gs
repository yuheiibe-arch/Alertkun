/**
 * ==========================================
 * シフトチェッカー Phase 2a: 募集枠・空き枠監査
 * ★UPDATE: 亀有・北葛西の「内科」「小児科」独立計算（空き枠スルー問題の修正）
 * ★UPDATE: 2診目の未掲載エラー（金額入り未掲載含む）を完全撤廃 ＆ 同時間帯の別募集枠も重複として判定
 * ==========================================
 */
function runShiftCheckerPhase2_Recruit() {
  const startTime = Date.now();
  Logger.log("=== Phase 2.5: 募集枠・空き枠監査 実行開始 ===");

  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet(); 
  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA'; 
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co'; 
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs'; 
  
  let pasteSs, shiftSs, locSs;
  try {
    pasteSs = SpreadsheetApp.openById(PASTE_MASTER_ID);
    shiftSs = SpreadsheetApp.openById(SHIFT_MASTER_ID);
    locSs = SpreadsheetApp.openById(LOC_MASTER_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message); return;
  }

  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const errorValues = [], errorBackgrounds = [];
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayYYYYMMDD = toYYYYMMDD(today); 
  
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  
  const scanEndDate = new Date(today.getFullYear(), today.getMonth() + 3, 0);
  
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  
  const EXCLUDED_CLINICS = ['院外勤務（小児科）', '嘱託医業務', '医師会業務', '【関東】バックアップシフト', '有給', '欠勤'];

  const existingIds = new Set();
  const archiveData = setupSheet(ACTIVE_SS, "アーカイブ", ALERT_HEADERS).getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) if (archiveData[i][11]) existingIds.add(String(archiveData[i][11])); 
  
  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  const alertData = alertSheet.getDataRange().getValues();
  for (let i = 1; i < alertData.length; i++) if (alertData[i][11]) existingIds.add(String(alertData[i][11]));

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId) => {
    if (!existingIds.has(uniqueId)) {
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, "", uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, doctor===""?"#eeeeee":null, empType===""?"#eeeeee":null, workTime===""?"#eeeeee":null, null, null, null, null]);
      existingIds.add(uniqueId);
    }
  };

  const locMaster = getCheckerLocationMaster(locSs);
  const closedDataMap = getCheckerClosedDays(pasteSs, scanStartDate, locMaster.normalize);

  // --- 「設定」シートの動的読み込み ---
  let holidaySkip = { start: 0, end: 0, isActive: false };
  let excludedSpecificClinics = new Set();
  
  const settingsSheet = ACTIVE_SS.getSheetByName("設定");
  if (settingsSheet) {
    const sData = settingsSheet.getDataRange().getValues();
    for (let r = 0; r < sData.length; r++) {
      const keyStr = String(sData[r][0]).trim();
      if (keyStr === "年末年始") {
        const sDate = parseDateToSafeDateObj(sData[r][1]);
        const eDate = parseDateToSafeDateObj(sData[r][2]);
        const isSkip = String(sData[r][3]).trim() === "検知なし";
        if (sDate && eDate && isSkip) {
          holidaySkip = { start: sDate.getTime(), end: eDate.getTime(), isActive: true };
        }
      } 
      else if (keyStr === "特定拠点") {
        for (let c = 1; c < sData[r].length; c++) {
          const cName = String(sData[r][c]).trim();
          if (cName) excludedSpecificClinics.add(locMaster.normalize(cName));
        }
      }
    }
  }

  const rawActualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, locMaster.normalize);

  // --- ★修正ポイント(1)：シフトをグループ化する際、亀有と北葛西は「科」もキーに含めて分離する ---
  const coverageMap = new Map();
  for (const [mapKey, actData] of rawActualShiftsMap.entries()) {
    const dateStr = mapKey.split('_')[0]; 
    const dObj = parseDateToSafeDateObj(dateStr);
    
    if (!dObj || dObj > scanEndDate) continue; 

    actData.shifts.forEach(shift => {
      if (!shift.normClinic || isNaN(shift.startMin) || isNaN(shift.endMin)) return;
      if (EXCLUDED_CLINICS.includes(shift.rawClinic)) return;
      
      const isMultiDept = shift.normClinic.includes("亀有") || shift.normClinic.includes("北葛西");
      let shiftDept = "";
      if (isMultiDept) {
         // 診療科カラムに内科とあれば内科、それ以外は小児科として扱う
         shiftDept = shift.dept.includes("内科") ? "内科" : "小児科"; 
      }
      // 亀有・北葛西の場合は「日付_拠点名_科」をキーにする
      const covKey = isMultiDept ? `${dateStr}_${shift.normClinic}_${shiftDept}` : `${dateStr}_${shift.normClinic}`;
      
      if (!coverageMap.has(covKey)) coverageMap.set(covKey, []);
      coverageMap.get(covKey).push(shift);
    });
  }

  const SYSTEM_SHEET_NAME = "SystemData_MissingBlocks";
  let systemSheet = ACTIVE_SS.getSheetByName(SYSTEM_SHEET_NAME);
  if (!systemSheet) {
    systemSheet = ACTIVE_SS.insertSheet(SYSTEM_SHEET_NAME);
    systemSheet.hideSheet();
  }
  
  let previousMissingData = {};
  const sysVal = systemSheet.getRange(1, 1).getValue();
  if (sysVal) {
    try {
      const parsed = JSON.parse(sysVal);
      if (Array.isArray(parsed)) {
        parsed.forEach(key => previousMissingData[key] = todayYYYYMMDD);
      } else if (typeof parsed === 'object') {
        previousMissingData = parsed;
      }
    } catch(e) {}
  }
  const currentMissingData = {};

  // --- 1. 未掲載エラー ＆ 1診目未掲載エラーの検知 ---
  coverageMap.forEach((shifts, covKey) => {
    const parts = covKey.split('_');
    const dateStr = parts[0];
    const normClinic = parts[1];
    const covDept = parts[2] || ""; // 亀有/北葛西の場合は「内科」か「小児科」が入る

    const dObj = new Date(dateStr);
    const msTime = dObj.getTime();

    if (msTime < today.getTime()) return; 

    if (excludedSpecificClinics.has(normClinic)) return;
    if (holidaySkip.isActive && msTime >= holidaySkip.start && msTime <= holidaySkip.end) return;

    const closedInfo = closedDataMap.get(`${dateStr}_${normClinic}`) || closedDataMap.get(`${dateStr}_全拠点`);
    if (closedInfo && closedInfo.type === "closed") return;

    const displayDate = `${dateStr}(${jpDays[dObj.getDay()]})`;
    const recruitShifts = shifts.filter(s => s.sourceSheet === "募集");
    
    // アラートの拠点名に「亀有(小児科)」のように出力するためのラベル
    const clinicLabel = covDept ? `${normClinic}(${covDept})` : normClinic;

    recruitShifts.forEach((rShift, index) => {
      const pStatus = rShift.publishStatus || "";
      const isUnpublished = pStatus.includes("未掲載") || pStatus.includes("非公開") || pStatus.includes("非掲載");

      if (isUnpublished) {
        const overlappingOthers = shifts.filter(s => 
          s !== rShift && s.startMin < rShift.endMin && s.endMin > rShift.startMin
        );
        const isFirstDoc = (overlappingOthers.length === 0);

        if (isFirstDoc) {
          const uniqueId = `1診目未掲載_${dateStr}_${clinicLabel}_${rShift.startStr}_${index}`;
          const firstDetectedDate = previousMissingData[uniqueId];
          
          if (!firstDetectedDate) {
            currentMissingData[uniqueId] = todayYYYYMMDD; 
          } else {
            currentMissingData[uniqueId] = firstDetectedDate;
            if (firstDetectedDate !== todayYYYYMMDD) { 
              const actionMsg = `規定の営業時間に対して必要なシフトが存在しません。\nシフトに空欄があります。確認・募集シフトを作成してください。`;
              const dispDept = covDept ? covDept : rShift.dept;
              addError("募集忘れ(未掲載)", displayDate, clinicLabel, dispDept, "未登録", "スポット", `${rShift.startStr}-${rShift.endStr}`, `1診目が未掲載`, actionMsg, uniqueId);
            }
          }
        }
      }
    });
  });

  // --- 2. 募集忘れ の検知（完全に空枠の場合） ---
  const activeClinics = Array.from(locMaster.locSystemStatus.entries())
                             .filter(([_, isActive]) => isActive)
                             .map(([name, _]) => name);

  for (let d = new Date(today.getTime()); d <= scanEndDate; d.setDate(d.getDate() + 1)) {
    const msTime = d.getTime();
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const displayDate = `${dateStr}(${jpDays[d.getDay()]})`;

    if (holidaySkip.isActive && msTime >= holidaySkip.start && msTime <= holidaySkip.end) continue;

    activeClinics.forEach(normClinic => {
      if (excludedSpecificClinics.has(normClinic)) return;

      let closedInfo = closedDataMap.get(`${dateStr}_${normClinic}`) || closedDataMap.get(`${dateStr}_全拠点`);
      if (closedInfo && closedInfo.type === "closed") return;

      // ★修正ポイント(2)：亀有と北葛西の場合は、内科用と小児科用の「2回」空き枠チェックを回す
      const isMultiDept = normClinic.includes("亀有") || normClinic.includes("北葛西");
      const deptsToCheck = isMultiDept ? ["内科", "小児科"] : [""];

      deptsToCheck.forEach(targetDept => {
        const covKey = isMultiDept ? `${dateStr}_${normClinic}_${targetDept}` : `${dateStr}_${normClinic}`;
        const dayShifts = coverageMap.get(covKey) || [];
        const isKitaKasai = normClinic.includes("北葛西");
        
        let requiredBlocks = [];
        if (closedInfo && closedInfo.type === "irregular") {
          requiredBlocks = closedInfo.ranges.map(r => ({
            start: r.startMin, end: r.endMin,
            startStr: formatMinutesToHHMM(r.startMin), endStr: formatMinutesToHHMM(r.endMin)
          }));
        } else {
          requiredBlocks = [
            { start: 9 * 60, end: 13 * 60, startStr: "09:00", endStr: "13:00" },
            { start: 15 * 60, end: 18 * 60, startStr: "15:00", endStr: "18:00" },
            { start: 18 * 60, end: isKitaKasai ? 20 * 60 : 21 * 60, startStr: "18:00", endStr: isKitaKasai ? "20:00" : "21:00" }
          ];
        }

        const missingBlocks = [];
        requiredBlocks.forEach(block => {
          const coveringShifts = dayShifts.filter(s => s.startMin < block.end && s.endMin > block.start);
          if (coveringShifts.length === 0) {
            missingBlocks.push(`${block.startStr}-${block.endStr}`);
          }
        });

        if (missingBlocks.length > 0) {
          const timeStr = missingBlocks.join(", ");
          const prefix = closedInfo && closedInfo.type === "irregular" ? "変則営業時間" : "規定の営業時間";
          const actionMsg = `${prefix}に対して必要なシフトが存在しません。\nシフトに空欄があります。確認・募集シフトを作成してください。`;
          
          // 表示ラベル（例: 亀有(小児科)）
          const displayDept = targetDept !== "" ? targetDept : (normClinic.includes("内科") ? "内科" : "小児科");
          const clinicLabel = targetDept !== "" ? `${normClinic}(${targetDept})` : normClinic;
          const uniqueId = `募集忘れ_${dateStr}_${clinicLabel}_${timeStr}`;

          const firstDetectedDate = previousMissingData[uniqueId];
          
          if (!firstDetectedDate) {
            currentMissingData[uniqueId] = todayYYYYMMDD;
          } else {
            currentMissingData[uniqueId] = firstDetectedDate;
            if (firstDetectedDate !== todayYYYYMMDD) {
              addError("募集忘れ", displayDate, clinicLabel, displayDept, "未登録", "", timeStr, `空き: ${timeStr}`, actionMsg, uniqueId);
            }
          }
        }
      });
    });
  }

  systemSheet.getRange(1, 1).setValue(JSON.stringify(currentMissingData));

  // --- 3. アラートリストへの書き出し ---
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
    Logger.log(`✅ ${sortedValues.length}件の[募集枠・空き枠エラー]を末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規の募集枠・空き枠エラーはありませんでした。");
  }
  
  Logger.log(`⏱ Phase 2.5 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}