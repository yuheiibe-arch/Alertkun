/**
 * ============================================================================
 * 【0-C_共通_時給計算エンジン】
 * ・NewWageEngine（年度・上期/下期時給テーブル ＆ 祝日カレンダー判定・日給計算）
 * ============================================================================
 */

var NewWageEngine = (function() {
  const LOC_WAGE_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  const HOLIDAY_MASTER_ID = '1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw';
  let _isInitialized = false;
  let _wageDB = {};
  let _holidayMap = {};

  function _init() {
    if (_isInitialized) return;
    try {
      const ss = getCachedSs(LOC_WAGE_MASTER_ID);
      const sheets = ss.getSheets();

      sheets.forEach(sheet => {
        const sName = sheet.getName();
        if (!sName.includes("時給") || sName.includes("特別")) return;

        _wageDB[sName] = {};
        const data = getCachedDisplayValues(sheet);
        for (let i = 1; i < data.length; i++) {
          const cId = String(data[i][0]).trim();
          const rawLoc = String(data[i][1]).trim();
          const rawDept = String(data[i][2]).trim();
          if (!cId && !rawLoc) continue;

          const getNum = (val) => Number(String(val).replace(/[^\d]/g, '')) || 0;

          const rateData = {
            wd_am: getNum(data[i][3]), wd_pm: getNum(data[i][4]), wd_nt: getNum(data[i][5]),
            sat_am: getNum(data[i][6]), sat_pm: getNum(data[i][7]), sat_nt: getNum(data[i][8]),
            hol_am: getNum(data[i][9]), hol_pm: getNum(data[i][10]), hol_nt: getNum(data[i][11])
          };

          if (cId) _wageDB[sName][`${cId}_${rawDept}`] = rateData;
          const nLoc = rawLoc
            .replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "")
            .replace(/\/.*/, "")
            .replace(/[\s ]+/g, "")
            .trim();
          _wageDB[sName][`${nLoc}_${rawDept}`] = rateData;
        }
      });

      const holSs = getCachedSs(HOLIDAY_MASTER_ID);
      holSs.getSheets().forEach(sheet => {
        const hData = getCachedValues(sheet);
        for (let i = 2; i < hData.length; i++) {
          const dObj = hData[i][0];
          if (dObj instanceof Date) {
            const dStr = toYYYYMMDD(dObj);
            const isHol = String(hData[i][3]).trim() !== "" || String(hData[i][5]).trim() !== "";
            _holidayMap[dStr] = isHol;
          }
        }
      });

      _isInitialized = true;
    } catch (e) {
      throw new Error("時給マスタ初期化失敗: " + e.message);
    }
  }

  function calculateDailyTotal(clinicId, locName, deptName, dateStr, startStr, endStr, contractText) {
    _init();

    const dObj = new Date(dateStr);
    const m = dObj.getMonth() + 1;
    const y = dObj.getFullYear();
    const fy = (m >= 4) ? y : y - 1;
    const half = (m >= 4 && m <= 9) ? "上期" : "下期";

    const targetSheetNames = [`${fy}${half}時給`, `${fy}時給`];
    let targetDB = null;
    for (const name of targetSheetNames) {
      if (_wageDB[name]) {
        targetDB = _wageDB[name];
        break;
      }
    }

    if (contractText && (contractText.includes("2025年度") || contractText.includes("2025年"))) {
      if (_wageDB['2025時給']) targetDB = _wageDB['2025時給'];
    }

    if (!targetDB) return { status: "ERROR", msg: "該当年度の時給シートなし" };

    const nLoc = locName
      .replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "")
      .replace(/\/.*/, "")
      .replace(/[\s ]+/g, "")
      .trim();
    let targetDept = deptName;
    if (nLoc === "亀有" || nLoc === "北葛西") {
      targetDept = (locName.includes("内科") || deptName.includes("内科")) ? "内科" : "小児科";
    }

    let rates = null;
    if (clinicId && targetDB[`${clinicId}_${targetDept}`]) {
      rates = targetDB[`${clinicId}_${targetDept}`];
    } else if (clinicId && targetDB[`${clinicId}_小児科`]) {
      rates = targetDB[`${clinicId}_小児科`];
    } else if (clinicId && targetDB[`${clinicId}_`]) {
      rates = targetDB[`${clinicId}_`];
    } else {
      rates = targetDB[`${nLoc}_${targetDept}`] || targetDB[`${nLoc}_小児科`] || targetDB[`${nLoc}_`];
    }

    if (!rates) return { status: "ERROR", msg: "マスタ未設定" };

    const day = dObj.getDay();
    const isHol = _holidayMap[dateStr] === true || day === 0;
    const prefix = isHol ? "hol" : (day === 6 ? "sat" : "wd");

    const sHour = parseInt(startStr.split(':')[0], 10);
    let eHour = parseInt(endStr.split(':')[0], 10);
    if (eHour < sHour && eHour !== 0) eHour += 24;

    let totalWage = 0;
    let maxHourly = 0;

    for (let h = sHour; h < eHour; h++) {
      if (h === 13 || h === 14) continue;

      let r = 0;
      if (h >= 9 && h < 13) r = rates[`${prefix}_am`];
      else if (h >= 15 && h < 18) r = rates[`${prefix}_pm`];
      else if (h >= 18 && h < 21) r = rates[`${prefix}_nt`];

      if (contractText && contractText.includes("円")) {
        const match = contractText.match(/(\d{1,2})[：:]\d{2}-(\d{1,2})[：:]\d{2}.*?(\d{1,2}(?:,\d{3})*)円/);
        if (match) {
          const spStart = parseInt(match[1], 10);
          const spEnd = parseInt(match[2], 10);
          if (h >= spStart && h < spEnd) r = parseInt(match[3].replace(/,/g, ''), 10);
        }
      }

      totalWage += r;
      if (r > maxHourly) maxHourly = r;
    }

    return { status: "SUCCESS", total: totalWage, maxHourly: maxHourly };
  }

  return { calculateDailyTotal };
})();