/**
 * ============================================================================
 * 【0-B_共通_マスタ・シフト読込】
 * ・拠点名マスタ（正規記載・システム反映・開院日判定）の読込
 * ・休館日・変則営業マスタの読込
 * ・全シフトシート（貼付用・確定・応募・募集）の期間別統合読込
 * ・2診要望一覧シートの読込
 * ============================================================================
 */

function getCheckerLocationMaster(locSs) {
  const cache = getAlertkunRuntimeCache_();
  const ssId = locSs.getId();
  if (cache.locMasterBySsId.has(ssId)) {
    return cache.locMasterBySsId.get(ssId);
  }

  const locMap = new Map();
  const locSystemStatus = new Map();
  const locOpenDates = new Map();
  const locSheet = locSs.getSheetByName("拠点名");

  if (locSheet) {
    const lData = getCachedDisplayValues(locSheet);
    const lHeaders = lData[0].map(h => String(h).replace(/[\s ]+/g, ''));
    const idxFormal = lHeaders.indexOf("正規記載");
    const idxStatus = lHeaders.indexOf("システム反映");
    let idxOpenDate = lHeaders.indexOf("開院日");
    if (idxOpenDate === -1) {
      idxOpenDate = lHeaders.findIndex(h => h.includes("開院") || h.includes("オープン日") || h.includes("開設日"));
    }
    const aliasIndices = lHeaders.map((h, i) => (h.includes("表記揺れ") || h.includes("クリニックNo")) ? i : -1).filter(i => i !== -1);

    if (idxFormal > -1) {
      for (let i = 1; i < lData.length; i++) {
        const formalName = String(lData[i][idxFormal]).replace(/[\s ]+/g, '');
        if (!formalName) continue;
        if (idxStatus > -1) locSystemStatus.set(formalName, String(lData[i][idxStatus]).trim() === "済");

        if (idxOpenDate > -1 && lData[i][idxOpenDate]) {
          const openDateObj = parseDateToSafeDateObj(lData[i][idxOpenDate]);
          if (openDateObj) {
            openDateObj.setHours(0, 0, 0, 0);
            locOpenDates.set(formalName, openDateObj);
          }
        }

        locMap.set(formalName, formalName);
        aliasIndices.forEach(idx => {
          const alias = String(lData[i][idx]).replace(/[\s ]+/g, '');
          if (alias) locMap.set(alias, formalName);
        });
      }
    }
  }

  const normalize = (rawName) => {
    const clean = String(rawName).replace(/[\s 【】]+/g, "");
    return locMap.has(clean) ? locMap.get(clean) : clean;
  };

  const isBeforeOpen = (clinicName, targetDate) => {
    const norm = normalize(clinicName);
    const openDate = locOpenDates.get(norm);
    if (!openDate) return false;
    const dObj = (targetDate instanceof Date) ? targetDate : parseDateToSafeDateObj(targetDate);
    if (!dObj) return false;
    const checkMs = new Date(dObj.getFullYear(), dObj.getMonth(), dObj.getDate()).getTime();
    return checkMs < openDate.getTime();
  };

  const result = {
    locSystemStatus,
    locOpenDates,
    normalize,
    isBeforeOpen
  };
  cache.locMasterBySsId.set(ssId, result);
  return result;
}

function getCheckerClosedDays(pasteSs, scanStartDate, normalizeFunc) {
  const closedDataMap = new Map();

  const parseTimeRanges = (timeStr) => {
    const ranges = [];
    const parts = timeStr.split(/[,、\s]+/);
    parts.forEach(p => {
      const t = p.split(/[-〜～]/);
      if (t.length === 2) {
        const startMin = parseTimeToMinutes(t[0]);
        const endMin = parseTimeToMinutes(t[1]);
        if (!isNaN(startMin) && !isNaN(endMin)) ranges.push({ startMin, endMin });
      }
    });
    return ranges;
  };

  const closedSheet = pasteSs.getSheetByName("休館日");
  if (closedSheet) {
    const cData = getCachedValues(closedSheet);
    if (cData.length > 0) {
      const headers = cData[0].map(h => String(h).replace(/[\s ]+/g, ''));
      const cDate = Math.max(0, headers.findIndex(h => h.includes('日付') || h.includes('対象日')));
      const cTime = Math.max(1, headers.findIndex(h => h.includes('時間') || h.includes('営業')));
      const cClinic = Math.max(2, headers.findIndex(h => h.includes('拠点名') || h.includes('拠点')));

      for (let i = 1; i < cData.length; i++) {
        const dateObj = parseDateToSafeDateObj(cData[i][cDate]);
        if (!dateObj || dateObj < scanStartDate) continue;

        const rawClinicName = String(cData[i][cClinic]).trim();
        const timeStr = String(cData[i][cTime]).trim();
        if (rawClinicName) {
          const clinicKey = rawClinicName.includes("全拠点") ? "全拠点" : normalizeFunc(rawClinicName);
          const key = `${toYYYYMMDD(dateObj)}_${clinicKey}`;

          if (!timeStr || timeStr.includes("休館") || timeStr.includes("全休") || timeStr.includes("全日")) {
            closedDataMap.set(key, { type: "closed" });
          } else {
            const ranges = parseTimeRanges(timeStr);
            if (ranges.length > 0) closedDataMap.set(key, { type: "irregular", ranges });
            else closedDataMap.set(key, { type: "closed" });
          }
        }
      }
    }
  }

  const irrSheet = pasteSs.getSheetByName("変則営業");
  if (irrSheet) {
    const iData = getCachedValues(irrSheet);
    if (iData.length > 0) {
      const headers = iData[0].map(h => String(h).replace(/[\s ]+/g, ''));
      const iDate = Math.max(0, headers.findIndex(h => h.includes('日付') || h.includes('対象日')));
      const iTime = Math.max(1, headers.findIndex(h => h.includes('時間') || h.includes('営業')));
      const iClinic = Math.max(2, headers.findIndex(h => h.includes('拠点名') || h.includes('拠点')));

      for (let i = 1; i < iData.length; i++) {
        const dateObj = parseDateToSafeDateObj(iData[i][iDate]);
        if (!dateObj || dateObj < scanStartDate) continue;

        const rawClinicName = String(iData[i][iClinic]).trim();
        const timeStr = String(iData[i][iTime]).trim();
        if (rawClinicName && timeStr) {
          const clinicKey = rawClinicName.includes("全拠点") ? "全拠点" : normalizeFunc(rawClinicName);
          const ranges = parseTimeRanges(timeStr);
          if (ranges.length > 0) {
            closedDataMap.set(`${toYYYYMMDD(dateObj)}_${clinicKey}`, { type: "irregular", ranges });
          }
        }
      }
    }
  }

  return closedDataMap;
}

function getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, normalizeFunc) {
  const actualShiftsMap = new Map();
  const today = new Date();
  const currentMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);

  const load = (sheet, sourceName) => {
    if (!sheet) return;
    const data = getCachedDisplayValues(sheet);
    if (data.length < 2) return;

    const cols = getColumnIndices(data[0], [
      '医籍番号', '名前', 'クリニックNo', 'クリニック名', '診療科', '勤務種別', '勤務日',
      '勤務開始時間', '勤務終了時間', '時給1', '時給2', '時給3', '時給4', '時給合計', '合計日給', '掲載ステータス', '業務内容', '備考', '応募日',
      'スタッフコメント1', 'スタッフコメント2', 'スタッフコメント3', 'スタッフコメント4', 'スタッフコメント5', '特別時給設定'
    ]);

    const totalCol = cols['時給合計'] !== -1 ? cols['時給合計'] : cols['合計日給'];
    const remarksCol = cols['業務内容'] !== -1 ? cols['業務内容'] : cols['備考'];
    const applyCol = cols['応募日'];
    const comCols = [1, 2, 3, 4, 5].map(n => cols[`スタッフコメント${n}`]).filter(idx => idx !== -1);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateObj = parseDateToSafeDateObj(row[cols['勤務日']]);
      if (!dateObj || dateObj < scanStartDate) continue;

      if (sourceName === "実績") {
        if (dateObj < currentMonthStart || dateObj >= thresholdDate) continue;
      } else if (sourceName === "確定") {
        if (dateObj >= currentMonthStart && dateObj < thresholdDate) continue;
      }

      if (sourceName === "応募") {
        const idNo = cols['医籍番号'] !== -1 ? String(row[cols['医籍番号']]).trim() : "";
        const docName = cols['名前'] !== -1 ? String(row[cols['名前']]).trim() : "";
        if (idNo === "310036" || docName.includes("橋本浩")) {
          continue;
        }
      }

      const dateStr = toYYYYMMDD(dateObj);
      const rawClinic = cols['クリニック名'] !== -1 ? String(row[cols['クリニック名']]) : "";
      let doctor = cols['名前'] !== -1 ? String(row[cols['名前']]).replace(/[\s ]+/g, '') : "";

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

      let applyDateObj = null;
      if (sourceName === "応募" && applyCol !== -1 && row[applyCol]) {
        applyDateObj = parseDateToSafeDateObj(row[applyCol]);
      }

      const staffComments = comCols.map(idx => String(row[idx]).trim()).filter(v => v !== "").join(" / ");

      record.shifts.push({
        sourceSheet: sourceName,
        dateStr: dateStr,
        doctorName: doctor,
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
          Number(String(row[cols['時給1']] || '').replace(/[^\d.-]/g, '')) || 0,
          Number(String(row[cols['時給2']] || '').replace(/[^\d.-]/g, '')) || 0,
          Number(String(row[cols['時給3']] || '').replace(/[^\d.-]/g, '')) || 0,
          Number(String(row[cols['時給4']] || '').replace(/[^\d.-]/g, '')) || 0
        ],
        wageTotal: totalCol !== -1 ? Number(String(row[totalCol] || '').replace(/[^\d.-]/g, '')) || 0 : 0,
        publishStatus: cols['掲載ステータス'] !== -1 ? String(row[cols['掲載ステータス']]).trim() : "",
        specialWageFlag: cols['特別時給設定'] !== -1 ? String(row[cols['特別時給設定']]).trim() : "",
        remarks: remarksCol !== -1 ? String(row[remarksCol]) : "",
        applyDateObj: applyDateObj,
        staffComments: staffComments
      });
    }
  };

  load(pasteSs.getSheetByName("貼付用"), "実績");
  load(shiftSs.getSheetByName("確定シフト"), "確定");
  load(shiftSs.getSheetByName("応募シフト"), "応募");
  load(shiftSs.getSheetByName("募集シフト"), "募集");

  return actualShiftsMap;
}

function getCheckerTwoDoctorRequests(twoDocSsId, scanStartDate, normalizeFunc) {
  const requestsMap = new Map();
  const ss = getCachedSs(twoDocSsId);
  if (!ss) return requestsMap;

  const sheet = ss.getSheetByName("２診要望一覧") || ss.getSheetByName("2診要望一覧") || ss.getSheets()[0];
  if (!sheet) return requestsMap;

  const data = getCachedDisplayValues(sheet);
  if (data.length < 2) return requestsMap;

  const headers = data[0].map(h => String(h).replace(/[\s ]+/g, '').toLowerCase());
  const colDate = headers.findIndex(h => h.includes('日付') || h.includes('対象日'));
  const colClinicId = headers.findIndex(h => h.includes('clinicid'));
  const colClinicName = headers.findIndex(h => h.includes('拠点') || h.includes('クリニック'));
  const colStart = headers.findIndex(h => h.includes('開始'));
  const colEnd = headers.findIndex(h => h.includes('終了'));

  if (colDate === -1) return requestsMap;

  for (let i = 1; i < data.length; i++) {
    const dObj = parseDateToSafeDateObj(data[i][colDate]);
    if (!dObj || dObj < scanStartDate) continue;

    const dateStr = toYYYYMMDD(dObj);
    const cId = colClinicId !== -1 ? String(data[i][colClinicId]).trim() : "";
    const rawName = colClinicName !== -1 ? String(data[i][colClinicName]) : "";
    const normName = normalizeFunc(rawName);

    let startMin = NaN, endMin = NaN;
    if (colStart !== -1 && colEnd !== -1) {
      startMin = parseTimeToMinutes(data[i][colStart]);
      endMin = parseTimeToMinutes(data[i][colEnd]);
    }

    const addRecord = (key) => {
      if (!requestsMap.has(key)) requestsMap.set(key, []);
      requestsMap.get(key).push({ startMin, endMin });
    };

    if (cId) addRecord(`${dateStr}_ID_${cId}`);
    if (normName) addRecord(`${dateStr}_NAME_${normName}`);
  }

  return requestsMap;
}