/**
 * ============================================================================
 * 【Phase4-B_採用監査_データ統合・引当】
 * ・シフトデータの診療科分離・1対1引き当て用リスト構築
 * ・採用くん「処理済み」シート（本日以降）の読込 ＆ 同一識別番号のグループ統合
 * ・2段階の厳格1対1引き当て（Pass 1: 識別番号一致 → Pass 2: 同日＋正規拠点＋診療科＋医師名）
 * ============================================================================
 */

/**
 * 期間ルール厳守のシフトデータを読み込み、確定配列と拠点・日付・診療科別マップを構築する
 */
function preparePhase4ShiftStructures_(pasteSs, shiftSs, scanStartDate, thresholdDate, strictNormalize) {
  const actualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, strictNormalize);
  const confirmedShifts = [];
  const clinicDateShifts = new Map();

  for (const actData of actualShiftsMap.values()) {
    actData.shifts.forEach(shift => {
      const isMultiDept = shift.normClinic.includes("亀有") || shift.normClinic.includes("北葛西");
      const shiftDept = isMultiDept
        ? (shift.dept.includes("内科") ? "内科" : "小児科")
        : (shift.normClinic.includes("内科") ? "内科" : "小児科");

      shift.normDept = shiftDept;
      shift.isLeaveOrAbsent = (
        shift.rawClinic.includes("欠勤") || shift.rawClinic.includes("有給") || shift.rawClinic.includes("有休") ||
        shift.type.includes("欠勤") || shift.type.includes("有給") || shift.type.includes("有休")
      );
      shift.usedByRecruitKey = null;

      const covKey = isMultiDept
        ? `${shift.dateStr}_${shift.normClinic}_${shiftDept}`
        : `${shift.dateStr}_${shift.normClinic}`;
      if (!clinicDateShifts.has(covKey)) clinicDateShifts.set(covKey, []);
      clinicDateShifts.get(covKey).push(shift);

      if (shift.sourceSheet === "実績" || shift.sourceSheet === "確定") {
        confirmedShifts.push(shift);
      }
    });
  }

  return { confirmedShifts, clinicDateShifts };
}

/**
 * 採用くん「処理済み」シートから本日以降の「採用」行を読み込み、同一識別番号ごとに1案件へ統合する
 */
function loadAndGroupPhase4Adoptions_(recruitSs, today, identifyClinic) {
  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  let recruitSheet = null;
  recruitSs.getSheets().forEach(s => {
    if (s.getSheetId() == '1155146432' || s.getName().includes("処理済み")) recruitSheet = s;
  });
  if (!recruitSheet) return null;

  const rData = getCachedDisplayValues(recruitSheet);
  if (rData.length < 2) return null;

  const rCols = getColumnIndices(rData[0], [
    '識別番号', '採用可否', '医師名', '勤務希望日', '拠点名', '診療科', '該当時間', '処理状況', '着信日', '対応時間', '対応者'
  ]);
  const sheetId = recruitSheet.getSheetId();
  const groupedAdoptions = new Map();

  for (let i = 1; i < rData.length; i++) {
    const row = rData[i];
    const saiyoVal = rCols['採用可否'] !== -1 ? String(row[rCols['採用可否']]).trim() : "";
    if (saiyoVal !== '採用') continue;

    const colIVal = String(row[8] || "").trim();
    const statusVal = rCols['処理状況'] !== -1 ? String(row[rCols['処理状況']]).trim() : "";
    const personVal = rCols['対応者'] !== -1 ? String(row[rCols['対応者']]).trim() : "";
    if (colIVal.includes("キャンセル") || statusVal.includes("キャンセル") || statusVal.includes("お断り") || statusVal.includes("掲載停止") || personVal === "不採用") {
      continue;
    }

    const dateObj = parseDateToSafeDateObj(row[rCols['勤務希望日']]);
    // 既に終了した過去日程（昨日以前）の採用行は監査対象外にする
    if (!dateObj || dateObj < today) continue;

    let isRecentlyProcessed = false;
    if (rCols['着信日'] !== -1 && row[rCols['着信日']]) {
      const receivedDate = parseDateToSafeDateObj(row[rCols['着信日']]);
      if (receivedDate && receivedDate.getTime() >= today.getTime()) isRecentlyProcessed = true;
    }
    if (rCols['対応時間'] !== -1 && row[rCols['対応時間']]) {
      const processedDate = parseDateToSafeDateObj(row[rCols['対応時間']]);
      if (processedDate && processedDate.getTime() >= today.getTime()) isRecentlyProcessed = true;
    }

    const dateStr = toYYYYMMDD(dateObj);
    const displayDate = `${dateStr}(${jpDays[dateObj.getDay()]})`;
    const rawClinic = String(row[rCols['拠点名']]).trim();
    const clinicInfo = identifyClinic(rawClinic);
    const normClinic = clinicInfo.formal;

    const rawDept = rCols['診療科'] !== -1 ? String(row[rCols['診療科']]).trim() : "";
    const isMultiDept = normClinic.includes("亀有") || normClinic.includes("北葛西");
    const dept = isMultiDept
      ? ((rawDept.includes("内科") || rawClinic.includes("内科")) ? "内科" : "小児科")
      : (normClinic.includes("内科") ? "内科" : "小児科");
    const clinicLabel = isMultiDept ? `${normClinic}(${dept})` : normClinic;

    const doctor = String(row[rCols['医師名']]).trim();
    const docClean = doctor.replace(/[\s ]+/g, '');
    const uKey = String(row[rCols['識別番号']]).trim().replace('Ｃ', 'C');

    const timeStr = String(row[rCols['該当時間']]).trim();
    const timeParts = timeStr.split(/[〜～~-]/);
    let sStr = "不明", eStr = "不明", sMin = NaN, eMin = NaN;
    if (timeParts.length >= 2) {
      sStr = timeParts[0].trim();
      eStr = timeParts[1].trim();
      sMin = parseTimeToMinutes(sStr);
      eMin = parseTimeToMinutes(eStr);
    }

    const dedupKey = uKey ? `UKEY_${uKey}` : `NOKEY_${dateStr}_${normClinic}_${docClean}_${sStr}-${eStr}`;

    if (!groupedAdoptions.has(dedupKey)) {
      groupedAdoptions.set(dedupKey, {
        dedupKey,
        rowIndices: [i],
        latestSheetRow: i + 1,
        isRecentlyProcessed,
        dateObj,
        dateStr,
        displayDate,
        rawClinic,
        normClinic,
        clinicResolved: clinicInfo.resolved,
        isMultiDept,
        dept,
        clinicLabel,
        doctor,
        docClean,
        uKey,
        sStr,
        eStr,
        sMin,
        eMin
      });
    } else {
      const existing = groupedAdoptions.get(dedupKey);
      existing.rowIndices.push(i);
      existing.latestSheetRow = i + 1;
      if (isRecentlyProcessed) existing.isRecentlyProcessed = true;
    }
  }

  return { sheetId, groupedAdoptions };
}

/**
 * 統合された採用案件群を確定シフトへ2段階で厳格に1対1引き当てする
 */
function matchPhase4AdoptionsToShifts_(groupedAdoptions, confirmedShifts) {
  const matchedShiftByGroup = new Map();

  // Pass 1: スタッフコメント内の「識別番号(uKey)」による確定引き当て
  groupedAdoptions.forEach(rec => {
    if (!rec.uKey || rec.uKey.length < 4) return;

    let hit = confirmedShifts.find(s =>
      !s.usedByRecruitKey && s.dateStr === rec.dateStr && s.staffComments && s.staffComments.includes(rec.uKey)
    );
    if (!hit) {
      hit = confirmedShifts.find(s =>
        !s.usedByRecruitKey && s.staffComments && s.staffComments.includes(rec.uKey)
      );
    }
    if (hit) {
      hit.usedByRecruitKey = rec.dedupKey;
      matchedShiftByGroup.set(rec.dedupKey, hit);
    }
  });

  // Pass 2: 補足照合「同日 ＋ 正規拠点名 ＋ 医師名（亀有・北葛西は診療科も一致）」で1対1引き当て
  groupedAdoptions.forEach(rec => {
    if (matchedShiftByGroup.has(rec.dedupKey)) return;
    if (!rec.clinicResolved) return;

    const candidates = confirmedShifts.filter(s => {
      if (s.usedByRecruitKey || s.isLeaveOrAbsent) return false;
      if (s.dateStr !== rec.dateStr) return false;
      if (s.doctorName !== rec.docClean) return false;
      if (s.normClinic !== rec.normClinic) return false;
      if (rec.isMultiDept && s.normDept !== rec.dept) return false;
      return true;
    });

    if (candidates.length > 0) {
      let hit = null;
      if (!isNaN(rec.sMin) && !isNaN(rec.eMin)) {
        hit = candidates.find(s => !isNaN(s.startMin) && !isNaN(s.endMin) && s.startMin < rec.eMin && s.endMin > rec.sMin);
      }
      if (!hit) hit = candidates[0];

      hit.usedByRecruitKey = rec.dedupKey;
      matchedShiftByGroup.set(rec.dedupKey, hit);
    }
  });

  return matchedShiftByGroup;
}