/**
 * ==========================================
 * シフトチェッカー Phase 1: 休館日・反映漏れ・Wブッキング検知
 * (対応指示・メモ欄追加 ＆ 休館日例外ルール対応版)
 * ★UPDATE: 採用くん機能のPhase4分離 ＆ 医籍番号によるWブッキング誤検知防止
 * ==========================================
 */
function runShiftCheckerPhase1() {
  const startTime = Date.now();
  Logger.log("=== Phase 1: シフトチェッカー 実行開始 ===");

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
    Logger.log("❌ スプレッドシート読み込みエラー: " + e.message); return;
  }

  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const errorValues = [], errorBackgrounds = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);

  const EXCLUDED_CLINICS = ['院外勤務（小児科）', '嘱託医業務', '医師会業務', '【関東】バックアップシフト', '有給', '欠勤'];
  const EXCLUDED_DOCTORS = ['⚠ 休館※医師勤務なし'];
  
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];

  const existingIds = new Set();
  const archiveData = setupSheet(ACTIVE_SS, "アーカイブ", ALERT_HEADERS).getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) if (archiveData[i][11]) existingIds.add(String(archiveData[i][11]));

  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  const alertData = alertSheet.getDataRange().getValues();
  for (let i = 1; i < alertData.length; i++) if (alertData[i][11]) existingIds.add(String(alertData[i][11]));

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId) => {
    if (!existingIds.has(uniqueId)) {
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, "", uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, doctor === "" ? "#eeeeee" : null, empType === "" ? "#eeeeee" : null, workTime === "" ? "#eeeeee" : null, null, null, null, null]);
      existingIds.add(uniqueId);
    }
  };

  const locMaster = getCheckerLocationMaster(locSs);
  const closedDataMap = getCheckerClosedDays(ACTIVE_SS, scanStartDate, locMaster.normalize);

  const unifiedShifts = [];

  const loadShifts = (sheet, isCurrentNextMonth, sourceName) => {
    if (!sheet) return;
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;
    
    // ★ 医籍番号を追加取得
    const cols = getColumnIndices(data[0], ['医籍番号', '名前', 'クリニック名', '診療科', '勤務日', '勤務開始時間', '勤務終了時間']);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateObj = parseDateToSafeDateObj(row[cols['勤務日']]);
      if (!dateObj || dateObj < scanStartDate) continue;
      
      if (sourceName !== "募集") {
        if (isCurrentNextMonth && dateObj >= thresholdDate) continue; 
        if (!isCurrentNextMonth && dateObj < thresholdDate) continue; 
      }

      const dateStr = toYYYYMMDD(dateObj);
      const displayDate = `${dateStr}(${jpDays[dateObj.getDay()]})`;
      const rawClinic = cols['クリニック名'] !== -1 ? String(row[cols['クリニック名']]).trim() : "";
      const dept = cols['診療科'] !== -1 ? String(row[cols['診療科']]).trim() : "";
      let doctor = cols['名前'] !== -1 ? String(row[cols['名前']]).trim() : "";
      const docId = cols['医籍番号'] !== -1 ? String(row[cols['医籍番号']]).trim() : ""; 
      
      if (sourceName === "募集" && !doctor) doctor = "募集";
      if (!rawClinic || !doctor) continue;

      const docClean = doctor.replace(/\s+/g, '');
      const normClinic = locMaster.normalize(rawClinic);
      
      const startStr = String(row[cols['勤務開始時間']]).trim();
      const endStr = String(row[cols['勤務終了時間']]).trim();
      const startMin = parseTimeToMinutes(startStr);
      const endMin = parseTimeToMinutes(endStr);

      if (EXCLUDED_CLINICS.includes(rawClinic) || EXCLUDED_DOCTORS.includes(doctor)) continue;
      
      unifiedShifts.push({ sourceSheet: sourceName, dateStr, displayDate, clinic: rawClinic, normClinic, dept, doctor: docClean, docId: docId, startStr, endStr, startMin, endMin });

      if (closedDataMap.has(`${dateStr}_${normClinic}`)) {
        if (sourceName === "募集" && startStr === "09:00" && endStr === "21:00") {
           // 例外許容
        } else {
           const actionMsg = `休館日です。${docClean}先生がアサインされています。\n該当シフトを削除し、医師へ連絡をお願いします。\n対象：${docClean}先生（${startStr}-${endStr}）`;
           addError("休館日アサイン", displayDate, rawClinic, dept, doctor, "", `${startStr}-${endStr}`, `休館日にアサインあり`, actionMsg, `休館_${dateStr}_${normClinic}_${docClean}_${i}`);
        }
      }
    }
  };

  loadShifts(pasteSs.getSheetByName("貼付用"), true, "実績");
  loadShifts(shiftSs.getSheetByName("確定シフト"), false, "確定");
  loadShifts(shiftSs.getSheetByName("募集シフト"), false, "募集"); 
  
  // --- Wブッキング検知 (20:00-21:00) ---
  const groupedShifts = {};
  unifiedShifts.forEach(shift => {
    if (shift.sourceSheet === "募集") return; 
    if (isNaN(shift.startMin) || isNaN(shift.endMin)) return;
    const key = `${shift.dateStr}_${shift.clinic}_${shift.dept}_${shift.displayDate}`; 
    if (!groupedShifts[key]) groupedShifts[key] = [];
    groupedShifts[key].push(shift);
  });

  const TARGET_START = 20 * 60, TARGET_END = 21 * 60;   
  for (const key in groupedShifts) {
    const targetDoctorsMap = new Map();
    const details = [];

    groupedShifts[key].forEach(shift => {
      if (shift.startMin < TARGET_END && shift.endMin > TARGET_START) {
        const docKey = shift.docId ? `ID_${shift.docId}` : `NAME_${shift.doctor}`;
        targetDoctorsMap.set(docKey, shift.doctor);
        details.push(`${shift.doctor}：${shift.startStr}-${shift.endStr}`);
      }
    });

    if (targetDoctorsMap.size >= 2) {
      const [dStr, cName, dName, displayDate] = key.split('_');
      const sortedDocs = Array.from(targetDoctorsMap.values()).sort().join('_');
      
      const docCount = targetDoctorsMap.size;
      const detailsStr = details.join('\n');
      const actionMsg = `同時間帯（20:00-21:00等）に${docCount}名の医師がアサインされています。\n${detailsStr}\n許容される２診であればA列にチェックを。\nミスの場合は対応をお願いします。`;
      
      addError("Wブッキング", displayDate, cName, dName, Array.from(targetDoctorsMap.values()).join(', '), "", details.join(', '), `複数アサイン`, actionMsg, `Wブッキング_${dStr}_${cName}_${sortedDocs}`);
    }
  }

  if (errorValues.length > 0) {
    const combined = errorValues.map((val, i) => ({ val, bg: errorBackgrounds[i] }));
    combined.sort((a, b) => a.val[11].localeCompare(b.val[11])); 
    const sortedValues = combined.map(item => item.val);
    const sortedBackgrounds = combined.map(item => item.bg);

    const startRow = alertSheet.getLastRow() + 1;
    alertSheet.getRange(startRow, 1, sortedValues.length, ALERT_HEADERS.length).setValues(sortedValues).setBackgrounds(sortedBackgrounds).setHorizontalAlignment("left");
    alertSheet.getRange(startRow, 1, sortedValues.length, 1).insertCheckboxes(); 
    Logger.log(`✅ ${sortedValues.length}件の新規エラーを末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規エラーはありません。");
  }
  
  Logger.log(`⏱ Phase 1 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}