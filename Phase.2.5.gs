/**
 * ==========================================
 * シフトチェッカー Phase 2.5: 募集枠・空き枠監査
 * (応募日タイムラグ除外 ＆ 全除外ルール ＆ 診療科区別 ＆ 2診要望ダブルマッチング 完成版)
 * ★UPDATE: 最新アラートリスト共通ルール（末尾追記型）対応
 * ==========================================
 */
function runShiftCheckerPhase2_Recruit() {
  const startTime = Date.now();
  Logger.log("=== Phase 2.5: 募集枠・空き枠監査 実行開始 ===");

  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet(); 
  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA'; 
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co'; 
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs'; 
  const TWO_DOC_SS_ID = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk'; 
  
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
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);
  
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  
  // 重複出力防止のための既存IDチェック（アーカイブ＆アラートリスト両方）
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
  const closedDataMap = getCheckerClosedDays(ACTIVE_SS, scanStartDate, locMaster.normalize);
  const actualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, locMaster.normalize);
  const twoDocRequestsMap = getCheckerTwoDoctorRequests(TWO_DOC_SS_ID, scanStartDate, locMaster.normalize);

  const coverageMap = new Map();
  for (const [mapKey, actData] of actualShiftsMap.entries()) {
    const dateStr = mapKey.split('_')[0]; 
    actData.shifts.forEach(shift => {
      if (!shift.normClinic || isNaN(shift.startMin) || isNaN(shift.endMin)) return;
      const covKey = `${dateStr}_${shift.normClinic}`;
      if (!coverageMap.has(covKey)) coverageMap.set(covKey, []);
      coverageMap.get(covKey).push(shift);
    });
  }

  const isNewYearHoliday = (dObj) => {
    const m = dObj.getMonth() + 1;
    const d = dObj.getDate();
    return (m === 12 && d >= 29) || (m === 1 && d <= 3);
  };

  // --- 1. 余剰募集の検知 ---
  coverageMap.forEach((shifts, covKey) => {
    const [dateStr, normClinic] = covKey.split('_');
    const displayDate = `${dateStr}(${jpDays[new Date(dateStr).getDay()]})`;

    if (new Date(dateStr) < today) return; // 過去日は対象外

    const confirmedShifts = shifts.filter(s => s.sourceSheet === "確定" || s.sourceSheet === "実績" || s.sourceSheet === "応募");
    const recruitShifts = shifts.filter(s => s.sourceSheet === "募集");

    recruitShifts.forEach(rShift => {
      const maxWage = Math.max(...rShift.wages, rShift.wageTotal);
      if (maxWage === 0) return; 

      if (rShift.rawClinic.includes("バックアップシフト")) return;
      if (normClinic.includes("北葛西") && (
        ((rShift.startStr === "17:00" || rShift.startStr === "18:00") && rShift.endStr === "20:00") ||
        (rShift.startStr === "15:00" && (rShift.endStr === "17:00" || rShift.endStr === "18:00"))
      )) return;
      if (rShift.remarks && (rShift.remarks.includes("併行募集") || rShift.remarks.includes("並行募集"))) return;

      let isTwoDocRequested = false;
      const reqs = [];
      if (rShift.clinicId && twoDocRequestsMap.has(`${dateStr}_ID_${rShift.clinicId}`)) {
        reqs.push(...twoDocRequestsMap.get(`${dateStr}_ID_${rShift.clinicId}`));
      }
      if (twoDocRequestsMap.has(`${dateStr}_NAME_${normClinic}`)) {
        reqs.push(...twoDocRequestsMap.get(`${dateStr}_NAME_${normClinic}`));
      }
      if (reqs.length > 0) {
        isTwoDocRequested = reqs.some(req => {
          if (isNaN(req.startMin) || isNaN(req.endMin)) return true;
          return (rShift.startMin < req.endMin && rShift.endMin > req.startMin);
        });
      }
      if (isTwoDocRequested) return;

      const isOverlapping = confirmedShifts.some(cShift => {
        if (cShift.dept === rShift.dept && rShift.startMin < cShift.endMin && rShift.endMin > cShift.startMin) {
          if (cShift.sourceSheet === "応募" && cShift.applyDateObj && toYYYYMMDD(cShift.applyDateObj) === todayYYYYMMDD) {
            return false;
          }
          const overlapMins = Math.min(rShift.endMin, cShift.endMin) - Math.max(rShift.startMin, cShift.startMin);
          return overlapMins > 60; 
        }
        return false;
      });

      if (isOverlapping) {
        const actionMsg = `すでに確定シフトが存在しますが、募集枠が残っています。該当の募集枠を取り下げ（クローズ）してください。`;
        addError("余剰募集", displayDate, rShift.rawClinic, rShift.dept, "募集枠", "", `${rShift.startStr}-${rShift.endStr}`, `確定シフトと重複`, actionMsg, `余剰募集_${dateStr}_${normClinic}_${rShift.startStr}`);
      }
    });
  });

  // --- 2. 募集忘れ ＆ 未掲載エラー の検知 ---
  const endDate = new Date(today.getFullYear(), today.getMonth() + 1, today.getDate());
  const activeClinics = Array.from(locMaster.locSystemStatus.entries())
                             .filter(([_, isActive]) => isActive)
                             .map(([name, _]) => name);

  for (let d = new Date(today.getTime()); d <= endDate; d.setDate(d.getDate() + 1)) {
    if (isNewYearHoliday(d)) continue;

    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const displayDate = `${dateStr}(${jpDays[d.getDay()]})`;

    activeClinics.forEach(normClinic => {
      let closedInfo = closedDataMap.get(`${dateStr}_${normClinic}`) || closedDataMap.get(`${dateStr}_全拠点`);
      if (closedInfo && closedInfo.type === "closed") return;

      const dayShifts = coverageMap.get(`${dateStr}_${normClinic}`) || [];
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
      const unpublishedBlocks = [];

      requiredBlocks.forEach(block => {
        const coveringShifts = dayShifts.filter(s => s.startMin < block.end && s.endMin > block.start);
        if (coveringShifts.length === 0) {
          missingBlocks.push(`${block.startStr}-${block.endStr}`);
        } else {
          const isProperlyCovered = coveringShifts.some(s => {
            if (s.sourceSheet === "確定" || s.sourceSheet === "実績" || s.sourceSheet === "応募") return true;
            if (s.sourceSheet === "募集") {
              const maxWage = Math.max(...s.wages, s.wageTotal);
              if (maxWage === 0) return true; 
              const pStatus = s.publishStatus || "";
              if (!pStatus.includes("未掲載") && !pStatus.includes("非公開")) return true; 
            }
            return false;
          });

          if (!isProperlyCovered) {
            unpublishedBlocks.push(`${block.startStr}-${block.endStr}`);
          }
        }
      });

      const dept = normClinic.includes("内科") ? "内科" : "小児科";

      if (missingBlocks.length > 0) {
        const timeStr = missingBlocks.join(", ");
        const prefix = closedInfo && closedInfo.type === "irregular" ? "変則営業時間" : "規定の営業時間";
        const actionMsg = `${prefix}に対して必要なシフトが存在しません。\nシフトに空欄があります。確認・募集シフトを作成してください。`;
        addError("募集忘れ", displayDate, normClinic, dept, "未登録", "", timeStr, `空き: ${timeStr}`, actionMsg, `募集忘れ_${dateStr}_${normClinic}_${timeStr}`);
      }

      if (unpublishedBlocks.length > 0) {
        const timeStr = unpublishedBlocks.join(", ");
        const actionMsg = `募集シフトは存在しますが、未掲載状態になっています。\n時給が設定されているため意図的でない可能性があります。募集を公開してください。`;
        addError("未掲載エラー", displayDate, normClinic, dept, "未登録", "", timeStr, `金額入り未掲載: ${timeStr}`, actionMsg, `未掲載_${dateStr}_${normClinic}_${timeStr}`);
      }
    });
  }

  // --- 3. アラートリストへの書き出し（末尾追記） ---
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