/**
 * ==========================================
 * シフトチェッカー Phase 4: 採用くん連携監査
 * (住所部分一致 ＆ 毒マップ完全排除 ＆ URLリッチテキスト化)
 * ==========================================
 */
function runShiftCheckerPhase4() {
  const startTime = Date.now();
  Logger.log("=== Phase 4: 採用くん連携監査 実行開始 ===");

  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet(); 
  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA'; 
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co'; 
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs'; 
  const RECRUIT_MASTER_ID = '1NfSmzERA-ee-8ToDYSwL_PYkxvMSxRv5Y_CehfeAQcQ'; 
  const TWO_DOC_SS_ID = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk'; 
  const HUB_MASTER_ID = '1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE'; 
  
  let pasteSs, shiftSs, locSs, recruitSs, hubSs;
  try {
    pasteSs = SpreadsheetApp.openById(PASTE_MASTER_ID);
    shiftSs = SpreadsheetApp.openById(SHIFT_MASTER_ID);
    locSs = SpreadsheetApp.openById(LOC_MASTER_ID);
    recruitSs = SpreadsheetApp.openById(RECRUIT_MASTER_ID);
    hubSs = SpreadsheetApp.openById(HUB_MASTER_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message); return;
  }

  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const errorValues = [], errorBackgrounds = [], errorRichTexts = [];
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);
  
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  
  const archivedIds = new Set();
  const archiveData = setupSheet(ACTIVE_SS, "アーカイブ", ALERT_HEADERS).getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) if (archiveData[i][11]) archivedIds.add(String(archiveData[i][11])); 

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId, targetUrl = null) => {
    if (!archivedIds.has(uniqueId)) {
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, "", uniqueId]);
      errorBackgrounds.push([null, null, null, null, null, doctor===""?"#eeeeee":null, empType===""?"#eeeeee":null, workTime===""?"#eeeeee":null, null, null, null, null]);
      
      const builder = SpreadsheetApp.newRichTextValue().setText(actionMsg);
      if (targetUrl) {
        const startIdx = actionMsg.indexOf(targetUrl);
        if (startIdx !== -1) {
          builder.setLinkUrl(startIdx, startIdx + targetUrl.length, targetUrl);
        }
      }
      errorRichTexts.push(builder.build());
    }
  };

  const locMaster = getCheckerLocationMaster(locSs);
  const twoDocRequests = getCheckerTwoDoctorRequests(TWO_DOC_SS_ID, scanStartDate, locMaster.normalize);
  
  // --- 0. 拠点マスタから住所マッピング用の辞書を作成 ---
  const locData = locSs.getSheetByName("拠点名").getDataRange().getDisplayValues();
  const locHeaders = locData[0].map(h => String(h).replace(/[\s ]+/g, ''));
  const idxFormal = locHeaders.indexOf("正規記載");
  const idxAddress = locHeaders.findIndex(h => h.includes("住所"));
  
  const addressMap = [];
  if (idxFormal !== -1 && idxAddress !== -1) {
    for (let i = 1; i < locData.length; i++) {
      const formal = String(locData[i][idxFormal]).trim();
      const address = String(locData[i][idxAddress]).trim();
      if (formal && address) addressMap.push({ formal, address });
    }
  }

  // --- 1. 紹介会社応募表（ハブ）の読み込み（フォールバック検証用のみに使用）---
  const hubMap = new Map();
  const hubSheet = hubSs.getSheetByName("紹介会社応募表");
  if (hubSheet) {
    const hData = hubSheet.getDataRange().getDisplayValues();
    const hCols = getColumnIndices(hData[0], ['名前', '勤務日', 'クリニック名']);
    for (let i = 1; i < hData.length; i++) {
      const dateObj = parseDateToSafeDateObj(hData[i][hCols['勤務日']]);
      if (!dateObj || dateObj < scanStartDate) continue;
      const dateStr = toYYYYMMDD(dateObj);
      const docClean = String(hData[i][hCols['名前']]).replace(/[\s ]+/g, '');
      const clinic = String(hData[i][hCols['クリニック名']]).trim();
      if (docClean && clinic) {
        const key = `${dateStr}_${docClean}`;
        if (!hubMap.has(key)) hubMap.set(key, new Set());
        hubMap.get(key).add(clinic);
      }
    }
  }

  // --- 2. スタッフコメント＆フォールバックセットの読み込み ---
  const allCommentsArray = []; 
  const fixedShiftFallbackSet = new Set(); 
  const scanShiftsForComments = (sheet) => {
    if (!sheet) return;
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;
    const cols = getColumnIndices(data[0], ['名前', 'クリニック名', '勤務日', 'スタッフコメント1', 'スタッフコメント2', 'スタッフコメント3', 'スタッフコメント4', 'スタッフコメント5']);
    
    for (let i = 1; i < data.length; i++) {
      const dateObj = parseDateToSafeDateObj(data[i][cols['勤務日']]);
      if (!dateObj || dateObj < scanStartDate) continue;
      const dateStr = toYYYYMMDD(dateObj);
      const rawClinic = cols['クリニック名'] !== -1 ? String(data[i][cols['クリニック名']]).trim() : "";
      const docClean = cols['名前'] !== -1 ? String(data[i][cols['名前']]).replace(/[\s ]+/g, '') : "";
      
      if (rawClinic && docClean) fixedShiftFallbackSet.add(`${dateStr}_${rawClinic}_${docClean}`);
      
      [1, 2, 3, 4, 5].forEach(num => {
        const idx = cols[`スタッフコメント${num}`];
        if (idx !== -1 && data[i][idx]) allCommentsArray.push(data[i][idx]);
      });
    }
  };
  scanShiftsForComments(pasteSs.getSheetByName("貼付用"));
  scanShiftsForComments(shiftSs.getSheetByName("確定シフト"));
  const megaCommentsString = allCommentsArray.join("|||"); 

  // --- 3. Wブッキング・ゴースト判定用のシフトマップ構築 ---
  const actualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, locMaster.normalize);
  const clinicDateShifts = new Map();
  for (const actData of actualShiftsMap.values()) {
    actData.shifts.forEach(shift => {
      const key = `${shift.dateStr}_${shift.normClinic}`;
      if (!clinicDateShifts.has(key)) clinicDateShifts.set(key, []);
      clinicDateShifts.get(key).push(shift);
    });
  }

  // --- 4. 採用くん（処理済み）データの走査 ---
  let recruitSheet = null;
  recruitSs.getSheets().forEach(s => {
    if (s.getSheetId() == '1155146432' || s.getName().includes("処理済み")) recruitSheet = s;
  });
  if (!recruitSheet) return;
  
  const rData = recruitSheet.getDataRange().getDisplayValues();
  if (rData.length < 2) return;
  
  const rCols = getColumnIndices(rData[0], ['識別番号', '採用可否', '医師名', '勤務希望日', '拠点名', '該当時間', '着信日', '対応時間']);
  const sheetId = recruitSheet.getSheetId();

  // メインループ
  for (let i = 1; i < rData.length; i++) {
    const row = rData[i];
    if (String(row[rCols['採用可否']]).trim() !== '採用') continue;
    
    const dateObj = parseDateToSafeDateObj(row[rCols['勤務希望日']]);
    if (!dateObj || dateObj < scanStartDate) continue;

    // 当日処理のラグ回避
    let isRecentlyProcessed = false;
    if (rCols['着信日'] !== -1 && row[rCols['着信日']]) {
      const receivedDate = parseDateToSafeDateObj(row[rCols['着信日']]);
      if (receivedDate && receivedDate.getTime() >= today.getTime()) isRecentlyProcessed = true;
    }
    if (rCols['対応時間'] !== -1 && row[rCols['対応時間']]) {
      const processedDate = parseDateToSafeDateObj(row[rCols['対応時間']]);
      if (processedDate && processedDate.getTime() >= today.getTime()) isRecentlyProcessed = true;
    }
    if (isRecentlyProcessed && dateObj > today) {
      continue; 
    }

    const dateStr = toYYYYMMDD(dateObj);
    const displayDate = `${dateStr}(${jpDays[dateObj.getDay()]})`;
    const rawClinic = String(row[rCols['拠点名']]).trim();
    const doctor = String(row[rCols['医師名']]).trim();
    const docClean = doctor.replace(/[\s ]+/g, '');
    const uKey = String(row[rCols['識別番号']]).trim();
    
    // ★拠点名の照合ロジック（毎ループ完全リセット）
    let standardClinic = "";

    if (rawClinic) {
      const cleanRaw = rawClinic.replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").replace(/[\s ]+/g, "");
      
      const addrMatch = addressMap.find(m => {
        const cleanMaster = m.address.replace(/[\s ]+/g, "");
        return cleanMaster.includes(cleanRaw) || cleanRaw.includes(cleanMaster);
      });
      
      if (addrMatch) {
        standardClinic = addrMatch.formal;
      } else {
        standardClinic = rawClinic; // 判定不能な場合は生データをそのまま残す
      }
    }
    
    const normClinic = locMaster.normalize(standardClinic);
    const dept = normClinic.includes("内科") ? "内科" : "小児科";

    const timeStr = String(row[rCols['該当時間']]).trim();
    const timeParts = timeStr.split(/[〜～-]/);
    let sStr = "不明", eStr = "不明", sMin = NaN, eMin = NaN;
    if (timeParts.length >= 2) {
      sStr = timeParts[0].trim();
      eStr = timeParts[1].trim();
      sMin = parseTimeToMinutes(sStr);
      eMin = parseTimeToMinutes(eStr);
    }
    
    // ① 反映漏れ（シフト未登録）チェック
    let isReflected = false;
    
    const dayClinicShifts = clinicDateShifts.get(`${dateStr}_${normClinic}`) || [];
    const hasConfirmed = dayClinicShifts.some(s => 
      (s.sourceSheet === "確定" || s.sourceSheet === "実績") && s.doctorName === docClean
    );
    if (hasConfirmed) isReflected = true;

    if (!isReflected && uKey && megaCommentsString.includes(uKey)) {
      isReflected = true;
    } 
    else if (!isReflected && hubMap.has(`${dateStr}_${docClean}`)) {
      for (const sClinic of hubMap.get(`${dateStr}_${docClean}`)) {
        if (fixedShiftFallbackSet.has(`${dateStr}_${sClinic}_${docClean}`)) {
          isReflected = true; break;
        }
      }
    }

    if (!isReflected) {
      const targetUrl = `https://docs.google.com/spreadsheets/d/${RECRUIT_MASTER_ID}/edit#gid=${sheetId}&range=${i+1}:${i+1}`;
      const actionMsg = `採用くんで採用となっていますが、確定シフトに登録がありません。\n確認してください。\n応募取り下げの場合は、採用可否のステータスを不採用にしてください。\n\n詳細(採用くん): ${targetUrl}`;
      
      addError("採用枠 未反映", displayDate, normClinic, dept, doctor, "スポット", `${sStr}-${eStr}`, "シフト未作成", actionMsg, `採用漏れ_${dateStr}_${normClinic}_${docClean}_${i}`, targetUrl);
      continue; 
    }

    // ② Wブッキング & ③ ゴースト募集 チェック
    if (!twoDocRequests.has(`${dateStr}_${normClinic}`) && !isNaN(sMin) && !isNaN(eMin)) {
      const overlappingOthers = dayClinicShifts.filter(s => 
        (s.sourceSheet === "確定" || s.sourceSheet === "実績") &&
        s.doctorName !== docClean &&
        s.startMin < eMin && s.endMin > sMin
      );

      if (overlappingOthers.length > 0) {
        const otherNames = overlappingOthers.map(s => s.doctorName).join(", ");
        const actionMsg = `採用くんの決定枠ですが、シフト上で別の医師（${otherNames}先生）がアサインされています。\n許容される２診であればA列にチェックを。ミスの場合は対応をお願いします。`;
        addError("紹介枠 Wブッキング", displayDate, normClinic, dept, doctor, "スポット", `${sStr}-${eStr}`, `重複: ${otherNames}`, actionMsg, `紹介重複_${dateStr}_${normClinic}_${docClean}_${i}`);
      }
    }

    if (!isNaN(sMin) && !isNaN(eMin)) {
      const remainingRecruits = dayClinicShifts.filter(s => 
        s.sourceSheet === "募集" &&
        s.startMin < eMin && s.endMin > sMin
      );

      if (remainingRecruits.length > 0) {
        const actionMsg = `採用くんで決定済みですが、同時間帯の募集枠が残っています。他媒体からの重複応募を防ぐため、該当の募集枠を取り下げ（クローズ）してください。`;
        addError("募集枠 消し忘れ", displayDate, normClinic, dept, doctor, "スポット", `${sStr}-${eStr}`, "募集枠が残存", actionMsg, `募集残存_${dateStr}_${normClinic}_${docClean}_${i}`);
      }
    }
  }

  // アラート出力
  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  if (errorValues.length > 0) {
    const combined = errorValues.map((val, i) => ({ val, bg: errorBackgrounds[i], rt: errorRichTexts[i] }));
    combined.sort((a, b) => a.val[11].localeCompare(b.val[11])); 
    
    const sortedValues = combined.map(item => item.val);
    const sortedBackgrounds = combined.map(item => item.bg);
    
    // ★ 修正: setRichTextValuesに渡すために、1次元配列を2次元配列に変換 [item.rt]
    const sortedRichTexts = combined.map(item => [item.rt]);

    const startRow = alertSheet.getLastRow() + 1;
    const requiredRows = startRow + sortedValues.length - 1;
    if (alertSheet.getMaxRows() < requiredRows) {
      alertSheet.insertRowsAfter(alertSheet.getMaxRows(), requiredRows - alertSheet.getMaxRows());
    }

    // 背景色と通常の値をセット
    alertSheet.getRange(startRow, 1, sortedValues.length, ALERT_HEADERS.length).setValues(sortedValues).setBackgrounds(sortedBackgrounds).setHorizontalAlignment("left");
    alertSheet.getRange(startRow, 1, sortedValues.length, 1).insertCheckboxes(); 
    
    // J列（10列目・対応指示）をリッチテキスト（部分リンク）で上書き
    alertSheet.getRange(startRow, 10, sortedRichTexts.length, 1).setRichTextValues(sortedRichTexts);
    
    Logger.log(`✅ ${sortedValues.length}件のエラー(採用くん監査)を末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規エラーはありませんでした。");
  }
  
  Logger.log(`⏱ Phase 4 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}