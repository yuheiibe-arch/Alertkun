/**
 * ==========================================
 * シフトチェッカー Phase 1: 休館日・反映漏れ・Wブッキング検知
 * (対応指示・メモ欄追加 ＆ 休館日例外ルール対応版)
 * ==========================================
 */
function runShiftCheckerPhase1() {
  const startTime = Date.now();
  Logger.log("=== Phase 1: シフトチェッカー 実行開始 ===");

  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet(); 
  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA'; 
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co'; 
  const RECRUIT_MASTER_ID = '1NfSmzERA-ee-8ToDYSwL_PYkxvMSxRv5Y_CehfeAQcQ'; 
  const HUB_MASTER_ID = '1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE'; 
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs'; 
  
  let pasteSs, shiftSs, recruitSs, hubSs, locSs;
  try {
    pasteSs = SpreadsheetApp.openById(PASTE_MASTER_ID);
    shiftSs = SpreadsheetApp.openById(SHIFT_MASTER_ID);
    recruitSs = SpreadsheetApp.openById(RECRUIT_MASTER_ID);
    hubSs = SpreadsheetApp.openById(HUB_MASTER_ID);
    locSs = SpreadsheetApp.openById(LOC_MASTER_ID);
  } catch (e) {
    Logger.log("❌ スプレッドシート読み込みエラー: " + e.message); return;
  }

  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const errorValues = [], errorBackgrounds = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  // Phase 1も統一して過去30日を起点に設定
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);

  const EXCLUDED_CLINICS = ['院外勤務（小児科）', '嘱託医業務', '医師会業務', '【関東】バックアップシフト', '有給', '欠勤'];
  const EXCLUDED_DOCTORS = ['⚠ 休館※医師勤務なし'];
  
  // ★ ヘッダーを拡張
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];

  // --- 0. 既存エラーIDの読み込み（重複防止・追記用） ---
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

  // --- 1. 共通ヘルパーの呼び出し（拠点マスタ＆休館日） ---
  const locMaster = getCheckerLocationMaster(locSs);
  const closedDataMap = getCheckerClosedDays(ACTIVE_SS, scanStartDate, locMaster.normalize);

  // --- 2. ハブ（紹介会社応募表）の読み込み ---
  const hubMap = new Map();
  const hubSheet = hubSs.getSheetByName("紹介会社応募表");
  if (hubSheet) {
    const hData = hubSheet.getDataRange().getDisplayValues();
    const hCols = getColumnIndices(hData[0], ['名前', '勤務日', 'クリニック名']);
    for (let i = 1; i < hData.length; i++) {
      const dateObj = parseDateToSafeDateObj(hData[i][hCols['勤務日']]);
      if (!dateObj || dateObj < scanStartDate) continue;
      
      const dateStr = toYYYYMMDD(dateObj);
      const docClean = String(hData[i][hCols['名前']]).replace(/\s+/g, '');
      const clinic = String(hData[i][hCols['クリニック名']]).trim();

      if (docClean && clinic) {
        const key = `${dateStr}_${docClean}`;
        if (!hubMap.has(key)) hubMap.set(key, new Set());
        hubMap.get(key).add(clinic);
      }
    }
  }

  // --- 3. メインシフトの読み込み ---
  const unifiedShifts = [];
  const allCommentsArray = []; 
  const fixedShiftFallbackSet = new Set(); 

  const loadShifts = (sheet, isCurrentNextMonth, sourceName) => {
    if (!sheet) return;
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;
    const cols = getColumnIndices(data[0], ['名前', 'クリニック名', '診療科', '勤務日', '勤務開始時間', '勤務終了時間', 'スタッフコメント1', 'スタッフコメント2', 'スタッフコメント3', 'スタッフコメント4', 'スタッフコメント5']);

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
      
      if (sourceName === "募集" && !doctor) doctor = "募集";
      if (!rawClinic || !doctor) continue;

      const docClean = doctor.replace(/\s+/g, '');
      const normClinic = locMaster.normalize(rawClinic);
      
      fixedShiftFallbackSet.add(`${dateStr}_${rawClinic}_${docClean}`);

      const startStr = String(row[cols['勤務開始時間']]).trim();
      const endStr = String(row[cols['勤務終了時間']]).trim();
      const startMin = parseTimeToMinutes(startStr);
      const endMin = parseTimeToMinutes(endStr);

      [1, 2, 3, 4, 5].forEach(num => {
        const idx = cols[`スタッフコメント${num}`];
        if (idx !== -1 && row[idx]) allCommentsArray.push(row[idx]);
      });

      if (EXCLUDED_CLINICS.includes(rawClinic) || EXCLUDED_DOCTORS.includes(doctor)) continue;
      
      unifiedShifts.push({ sourceSheet: sourceName, dateStr, displayDate, clinic: rawClinic, normClinic, dept, doctor: docClean, startStr, endStr, startMin, endMin });

      // ★ 休館日チェック（募集シフトかつ09:00-21:00なら除外する例外ルール追加）
      if (closedDataMap.has(`${dateStr}_${normClinic}`)) {
        if (sourceName === "募集" && startStr === "09:00" && endStr === "21:00") {
           // 例外として許容するためスルー
        } else {
           const actionMsg = `休館日です。${docClean}先生がアサインされています。\n該当シフトを削除し、医師へ連絡をお願いします。\n対象：${docClean}先生（${startStr}-${endStr}）`;
           addError("休館日アサイン", displayDate, rawClinic, dept, doctor, "", `${startStr}-${endStr}`, `休館日にアサインあり`, actionMsg, `休館_${dateStr}_${normClinic}_${docClean}_${i}`);
        }
      }
    }
  };

  loadShifts(pasteSs.getSheetByName("貼付用"), true, "実績");
  loadShifts(shiftSs.getSheetByName("確定シフト"), false, "確定");
  loadShifts(shiftSs.getSheetByName("募集シフト"), false, "募集"); // 休館日例外用に追加
  
  const megaCommentsString = allCommentsArray.join("|||"); 

  // --- 4. 採用くん（反映漏れチェック） ---
  const recruitSheet = recruitSs.getSheetByName("処理済み");
  if (recruitSheet) {
    const rData = recruitSheet.getDataRange().getDisplayValues();
    const idColIdx = rData[0].map(h => String(h).replace(/\r?\n/g, '').trim()).findIndex(h => h.includes('識別番号'));
    const rCols = getColumnIndices(rData[0], ['医師名', '拠点名', '診療科', '勤務希望日', '採用可否', '勤務開始時間', '勤務終了時間']);
    
    for (let i = 1; i < rData.length; i++) {
      const row = rData[i];
      if (String(row[rCols['採用可否']]).trim() !== '採用') continue;
      
      const dateObj = parseDateToSafeDateObj(row[rCols['勤務希望日']]);
      if (!dateObj || dateObj < scanStartDate) continue;

      const uKey = idColIdx !== -1 ? String(row[idColIdx]).trim() : "";
      if (!uKey) continue;

      const doctor = String(row[rCols['医師名']]).trim();
      const rawClinic = String(row[rCols['拠点名']]).trim();
      const dept = String(row[rCols['診療科']]).trim();
      const sTime = rCols['勤務開始時間'] !== -1 ? String(row[rCols['勤務開始時間']]).trim() : "";
      const eTime = rCols['勤務終了時間'] !== -1 ? String(row[rCols['勤務終了時間']]).trim() : "";
      
      const dateStr = toYYYYMMDD(dateObj);
      const displayDate = `${dateStr}(${jpDays[dateObj.getDay()]})`;
      const docClean = doctor.replace(/\s+/g, '');
      const hubKey = `${dateStr}_${docClean}`;
      
      let isReflected = false, standardClinic = rawClinic;
      if (hubMap.has(hubKey)) standardClinic = Array.from(hubMap.get(hubKey))[0];

      if (megaCommentsString.includes(uKey)) {
        isReflected = true;
      } else if (hubMap.has(hubKey)) {
        for (const sClinic of hubMap.get(hubKey)) {
          if (fixedShiftFallbackSet.has(`${dateStr}_${sClinic}_${docClean}`)) {
            isReflected = true; break;
          }
        }
      }

      if (!isReflected) {
        const actionMsg = `採用くんで採用となっていますが、シフトが登録されていません。\n勤務なら勤務作成\n欠勤なら欠勤シフトを作成してください。\n\n通常勤務：${sTime}-${eTime}\n現在の登録：なし`;
        addError("シフト未登録", displayDate, standardClinic, dept, doctor, "", `${sTime}-${eTime}`, `確定シフトに存在しません`, actionMsg, `漏れ_${dateStr}_${standardClinic}_${docClean}`);
      }
    }
  }

  // --- 5. Wブッキング検知 (20:00-21:00) ---
  const groupedShifts = {};
  unifiedShifts.forEach(shift => {
    if (shift.sourceSheet === "募集") return; // 募集枠同士のWブッキングは別で処理するため除外
    if (isNaN(shift.startMin) || isNaN(shift.endMin)) return;
    const key = `${shift.dateStr}_${shift.clinic}_${shift.dept}_${shift.displayDate}`; 
    if (!groupedShifts[key]) groupedShifts[key] = [];
    groupedShifts[key].push(shift);
  });

  const TARGET_START = 20 * 60, TARGET_END = 21 * 60;   
  for (const key in groupedShifts) {
    const targetDoctors = new Set();
    const details = [];

    groupedShifts[key].forEach(shift => {
      if (shift.startMin < TARGET_END && shift.endMin > TARGET_START) {
        targetDoctors.add(shift.doctor);
        details.push(`${shift.doctor}：${shift.startStr}-${shift.endStr}`);
      }
    });

    if (targetDoctors.size >= 2) {
      const [dStr, cName, dName, displayDate] = key.split('_');
      const sortedDocs = Array.from(targetDoctors).sort().join('_');
      
      const docCount = targetDoctors.size;
      const detailsStr = details.join('\n');
      const actionMsg = `同時間帯（20:00-21:00等）に${docCount}名の医師がアサインされています。\n${detailsStr}\n許容される２診であればA列にチェックを。\nミスの場合は対応をお願いします。`;
      
      addError("Wブッキング", displayDate, cName, dName, Array.from(targetDoctors).join(', '), "", details.join(', '), `複数アサイン`, actionMsg, `Wブッキング_${dStr}_${cName}_${sortedDocs}`);
    }
  }

  // --- 6. アラートリストへの書き出し（クリアせず末尾追記＆昇順ソート） ---
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