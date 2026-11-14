/**
 * ==========================================
 * シフトチェッカー Phase 3: 時給・有休監査（高速キャッシュ・募集複数枠対応版）
 * ★UPDATE: 過去の募集シフトの時給監査を除外 ＆ アーカイブシートからの除外リスト読み込みを修正
 * ★UPDATE: 過去年度の保存シートを無視し、対象年度(最新)のシートのみを読み込むよう修正
 * ★UPDATE: 募集シフトの「特別時給設定」を判定し、1円でも高ければ厳格にエラーを出す仕様に変更
 * ★UPDATE: 診療科、または業務内容/備考に「ワクチン」が含まれる特殊シフトは時給エラーから除外
 * ★UPDATE: 年末年始（12/29〜1/3）のみ特別時給として上限を22,000円に引き上げ（特別設定フラグ不問・22,000円超で高時給エラー）
 * ★UPDATE: 募集シフトの同日複数コマで時給エラーが潰れないよう時間帯別ユニークキーに対応（旧アーカイブ互換維持）
 * ★UPDATE: オンメモリキャッシュ共用による高速化 ＆ Phase2と重複していた死蔵Jinjer読込を整理
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

  const openSs = (id) => (typeof getCachedSs === 'function') ? getCachedSs(id) : SpreadsheetApp.openById(id);
  const readValues = (sh) => (typeof getCachedValues === 'function') ? getCachedValues(sh) : sh.getDataRange().getValues();
  const readDisplayValues = (sh) => (typeof getCachedDisplayValues === 'function') ? getCachedDisplayValues(sh) : sh.getDataRange().getDisplayValues();

  let pasteSs, shiftSs, locSs, attSs;
  try {
    pasteSs = openSs(PASTE_MASTER_ID);
    shiftSs = openSs(SHIFT_MASTER_ID);
    locSs = openSs(LOC_MASTER_ID);
    attSs = openSs(ATTENDANCE_SS_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message);
    return;
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

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, uniqueId, legacyUid = null) => {
    if (existingIds.has(uniqueId)) return;
    if (legacyUid && existingIds.has(legacyUid)) return;

    errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, "", "", uniqueId]);
    errorBackgrounds.push([null, null, null, null, null, doctor === "" ? "#eeeeee" : null, empType === "" ? "#eeeeee" : null, workTime === "" ? "#eeeeee" : null, null, null, null, null]);
    existingIds.add(uniqueId);
  };

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

    const yearMatch = sName.match(/\d{4}/);
    if (yearMatch && !targetNendos.includes(parseInt(yearMatch[0], 10))) {
      return;
    }

    if (sName.includes("年度") && !sName.includes("勤怠")) {
      const eType = sName.includes("常勤") && !sName.includes("非常勤") ? "常勤" : (sName.includes("定期非常勤") ? "定期非常勤" : "");
      if (!eType) return;

      const data = readValues(sheet);
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
      const eType = sName.includes("常勤") && !sName.includes("非常勤") ? "常勤" : (sName.includes("定期非常勤") ? "定期非常勤" : "");
      if (!eType) return;

      const data = readDisplayValues(sheet);
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
      const shiftDateObj = parseDateToSafeDateObj(dateStr);
      if (!shiftDateObj) return;
      const displayDate = `${dateStr}(${jpDays[shiftDateObj.getDay()]})`;
      const shiftTimeMs = shiftDateObj.getTime();

      if (EXCLUDED_CLINICS.includes(shift.rawClinic) || EXCLUDED_DOCTORS.includes(docClean)) return;

      const closedInfo = closedDataMap.get(`${dateStr}_${shift.normClinic}`) || closedDataMap.get(`${dateStr}_全拠点`);
      if (closedInfo && closedInfo.type === "closed") {
        if (shift.sourceSheet === "募集" || docClean === "募集") {
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

      if (shift.sourceSheet === "募集" || docClean === "募集") {
        if (shiftTimeMs < today.getTime()) return;
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
      const recruitTimeSuffix = shift.sourceSheet === "募集" ? `_${shift.startStr}` : "";

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

      // ★ 年末年始（12/29〜1/3）判定：この期間のみ特別時給上限を22,000円に設定
      const shiftMonth = shiftDateObj.getMonth() + 1;
      const shiftDay = shiftDateObj.getDate();
      const isNewYearPeriod = (shiftMonth === 12 && shiftDay >= 29) || (shiftMonth === 1 && shiftDay <= 3);
      const baseUpperLimit = isNewYearPeriod ? 22000 : 20000;

      let allowedMax = baseUpperLimit;
      let expectedTotalStr = "算出不可";
      let expectedHourly = 0;
      let expectedTotal = 0;

      if (empType !== "常勤") {
        try {
          const res = NewWageEngine.calculateDailyTotal(shift.clinicId, shift.normClinic, shift.dept, dateStr, shift.startStr, shift.endStr, specialText);
          if (res.status === "SUCCESS") {
            expectedTotal = res.total;
            expectedTotalStr = `￥${Math.round(res.total).toLocaleString()}`;
            allowedMax = isNewYearPeriod ? 22000 : (res.maxHourly + 2000);
            expectedHourly = res.maxHourly;
          } else {
            expectedTotalStr = res.msg;
          }
        } catch (e) {
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
          const baseUid = `${shift.sourceSheet}_未入力_${dateStr}_${shift.normClinic}_${docClean}`;
          addError(`${pfx}時給・日給未入力`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正：${expectedTotalStr}\n実：￥0`, `${baseUid}${recruitTimeSuffix}`, baseUid);
        } else if (shift.sourceSheet === "募集" && expectedHourly > 0) {
          const isUnpublished = shift.publishStatus && (shift.publishStatus.includes("未掲載") || shift.publishStatus.includes("非公開") || shift.publishStatus.includes("非掲載"));
          const isIrregularTime = (shift.startMin % 60 !== 0 || shift.endMin % 60 !== 0);
          const isVaccineShift = shift.dept.includes("ワクチン") || shift.remarks.includes("ワクチン");

          if (!isUnpublished && !isIrregularTime && !isVaccineShift) {
            const actualHourly = maxWage > 0 ? maxWage : (shift.wageTotal > 0 && shift.wageTotal <= baseUpperLimit ? shift.wageTotal : 0);

            const rawSpecialFlag = shift.specialWageFlag || "";
            const isSpecialWage = rawSpecialFlag.includes("特別時給設定あり") || rawSpecialFlag.includes("TRUE") || rawSpecialFlag.includes("true") || rawSpecialFlag.includes("あり") || rawSpecialFlag.includes("☑");

            let isTooCheap = false;
            let isTooExpensive = false;

            if ((actualHourly > 0 && actualHourly < expectedHourly) || (shift.wageTotal > baseUpperLimit && expectedTotal > 0 && shift.wageTotal < expectedTotal)) {
              isTooCheap = true;
            }

            // 年末年始(12/29〜1/3)は特別時給設定フラグが無くても22,000円以下なら許容（22,000円超は下の「上限時給超過」で検知）
            if (!isNewYearPeriod && actualHourly > 0 && actualHourly > expectedHourly) {
              if (!isSpecialWage) {
                isTooExpensive = true;
              }
            }

            if (isTooCheap || isTooExpensive) {
              const errorTitle = isTooExpensive ? `${pfx}時給設定エラー(高額/特別設定なし)` : `${pfx}時給設定エラー(不一致/安い)`;

              let correctMsg = `正：時給￥${expectedHourly.toLocaleString()}`;
              let actualMsg = `実：時給￥${actualHourly > 0 ? actualHourly.toLocaleString() : maxWage.toLocaleString()}`;

              if (expectedTotal > 0 && shift.wageTotal > 0) {
                correctMsg += ` (日給￥${expectedTotal.toLocaleString()})`;
                actualMsg += ` (日給￥${shift.wageTotal.toLocaleString()})`;
              }

              const baseUid = `${shift.sourceSheet}_時給不一致_${dateStr}_${shift.normClinic}_${docClean}`;
              addError(errorTitle, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `${correctMsg}\n${actualMsg}`, `${baseUid}${recruitTimeSuffix}`, baseUid);
            }
          }
        }
      }

      // ★上限時給超過チェック（年末年始は22,000円超でエラー、通常期は20,000円超かつallowedMax超でエラー）
      if (isNewYearPeriod) {
        if (maxWage > 22000) {
          const baseUid = `${shift.sourceSheet}_上限超過_${dateStr}_${shift.normClinic}_${docClean}`;
          addError(`${pfx}上限時給超過`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正(上限)：￥22,000\n実：￥${maxWage.toLocaleString()}`, `${baseUid}${recruitTimeSuffix}`, baseUid);
        }
      } else {
        if (maxWage > 20000 && maxWage > allowedMax) {
          const baseUid = `${shift.sourceSheet}_上限超過_${dateStr}_${shift.normClinic}_${docClean}`;
          addError(`${pfx}上限時給超過`, displayDate, shift.rawClinic, shift.dept, docClean, empType, `${shift.startStr}-${shift.endStr}`, `正(上限)：￥${allowedMax.toLocaleString()}\n実：￥${maxWage.toLocaleString()}`, `${baseUid}${recruitTimeSuffix}`, baseUid);
        }
      }
    });
  }

  if (errorValues.length > 0) {
    const combined = errorValues.map((val, i) => ({ val, bg: errorBackgrounds[i] }));
    combined.sort((a, b) => a.val[2].localeCompare(b.val[2]));
    const sortedValues = combined.map(item => item.val);
    const sortedBackgrounds = combined.map(item => item.bg);

    if (typeof appendAlertRowsSafe === 'function') {
      appendAlertRowsSafe(alertSheet, ALERT_HEADERS.length, sortedValues, sortedBackgrounds);
    } else {
      const startRow = alertSheet.getLastRow() + 1;
      const requiredRows = startRow + sortedValues.length - 1;
      if (alertSheet.getMaxRows() < requiredRows) {
        alertSheet.insertRowsAfter(alertSheet.getMaxRows(), requiredRows - alertSheet.getMaxRows());
      }
      alertSheet.getRange(startRow, 1, sortedValues.length, ALERT_HEADERS.length)
        .setValues(sortedValues)
        .setBackgrounds(sortedBackgrounds)
        .setHorizontalAlignment("left");
      alertSheet.getRange(startRow, 1, sortedValues.length, 1).insertCheckboxes();
    }
    Logger.log(`✅ ${sortedValues.length}件のエラー(時給・有休監査)を末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規エラーはありませんでした。");
  }

  Logger.log(`⏱ Phase 3 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}