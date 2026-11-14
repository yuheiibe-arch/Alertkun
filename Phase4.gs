/**
 * ==========================================
 * シフトチェッカー Phase 4: 採用くん連携監査（抜本改修・厳格照合版）
 * ★UPDATE: 誤判定の原因だった紹介会社応募表(hubMap)・無制限コメント結合フォールバックを撤廃
 * ★UPDATE: 拠点名・エムスリー住所表記（漢数字/全角数字/一般内科/施設名付き等）の100%正規同定エンジンを搭載
 * ★UPDATE: 採用くん「処理済み」の同一識別番号行（応募通知受信＋勤務確定）を1案件へ統合
 * ★UPDATE: 識別番号一致(最優先) ＋ 同日・正規拠点・診療科・医師名(1対1引当) による2段階厳格照合
 * ★UPDATE: シート並び替えでもリセットされない識別番号ベースの固定ユニークキーへ変更（旧アーカイブ互換維持）
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

  let pasteSs, shiftSs, locSs, recruitSs;
  try {
    pasteSs = SpreadsheetApp.openById(PASTE_MASTER_ID);
    shiftSs = SpreadsheetApp.openById(SHIFT_MASTER_ID);
    locSs = SpreadsheetApp.openById(LOC_MASTER_ID);
    recruitSs = SpreadsheetApp.openById(RECRUIT_MASTER_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message);
    return;
  }

  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const errorValues = [], errorBackgrounds = [], errorRichTexts = [];

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayYYYYMMDD = toYYYYMMDD(today);

  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);

  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  const EXCLUDED_NON_CLINICS = ['院外勤務（小児科）', '嘱託医業務', '医師会業務', '【関東】バックアップシフト', '有給', '有休', '欠勤'];

  // --- アーカイブ済みユニークキーの読み込み（旧形式の行番号付きキーとの後方互換セットも構築） ---
  const archivedIds = new Set();
  const legacyArchivedPrefixes = new Set();
  const archiveData = setupSheet(ACTIVE_SS, "アーカイブ", ALERT_HEADERS).getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) {
    const uid = String(archiveData[i][11] || "").trim();
    if (!uid) continue;
    archivedIds.add(uid);
    const legacyMatch = uid.match(/^((?:採用漏れ|紹介重複|募集残存)_\d{4}\/\d{2}\/\d{2}_.+_[^_]+)_\d+$/);
    if (legacyMatch) {
      legacyArchivedPrefixes.add(legacyMatch[1]);
    }
  }

  const isArchivedError = (uniqueId, legacyPrefix, rowIndices) => {
    if (archivedIds.has(uniqueId)) return true;
    if (rowIndices && rowIndices.some(idx => archivedIds.has(`${legacyPrefix}_${idx}`))) return true;
    if (legacyArchivedPrefixes.has(legacyPrefix)) return true;
    return false;
  };

  const addError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId, legacyPrefix, rowIndices, targetUrl = null) => {
    if (!isArchivedError(uniqueId, legacyPrefix, rowIndices)) {
      errorValues.push([false, type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, "", uniqueId]);
      errorBackgrounds.push([
        null, null, null, null, null,
        doctor === "" ? "#eeeeee" : null,
        empType === "" ? "#eeeeee" : null,
        workTime === "" ? "#eeeeee" : null,
        null, null, null, null
      ]);

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

  // --- タイムラグ考慮（2日連続検知）用システムシートの準備 ---
  const SYSTEM_SHEET_NAME = "SystemData_Phase4Lag";
  let systemSheet = ACTIVE_SS.getSheetByName(SYSTEM_SHEET_NAME);
  if (!systemSheet) {
    systemSheet = ACTIVE_SS.insertSheet(SYSTEM_SHEET_NAME);
    systemSheet.hideSheet();
  }
  let previousDelayData = {};
  const sysVal = systemSheet.getRange(1, 1).getValue();
  if (sysVal) {
    try {
      const parsed = JSON.parse(sysVal);
      if (parsed && typeof parsed === 'object') previousDelayData = parsed;
    } catch (e) {}
  }
  const currentDelayData = {};

  // 旧形式キーで昨日記録されていた検知日も引き継げるヘルパー
  const findPreviousDetectedDate = (uniqueId, legacyPrefix, rowIndices) => {
    if (previousDelayData[uniqueId]) return previousDelayData[uniqueId];
    if (rowIndices) {
      for (const idx of rowIndices) {
        if (previousDelayData[`${legacyPrefix}_${idx}`]) {
          return previousDelayData[`${legacyPrefix}_${idx}`];
        }
      }
    }
    return null;
  };

  const processLagError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId, legacyPrefix, rowIndices, targetUrl = null) => {
    const firstDetectedDate = findPreviousDetectedDate(uniqueId, legacyPrefix, rowIndices);
    if (!firstDetectedDate) {
      currentDelayData[uniqueId] = todayYYYYMMDD;
    } else {
      currentDelayData[uniqueId] = firstDetectedDate;
      if (firstDetectedDate !== todayYYYYMMDD) {
        addError(type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId, legacyPrefix, rowIndices, targetUrl);
      }
    }
  };

  // --- 1. 拠点マスタ読み込み ＆ 100%拠点同定エンジンの構築 ---
  const locData = locSs.getSheetByName("拠点名").getDataRange().getDisplayValues();
  const locHeaders = locData[0].map(h => String(h).replace(/[\s ]+/g, ''));
  const idxFormal = locHeaders.indexOf("正規記載");
  const idxAddress = locHeaders.findIndex(h => h.includes("住所"));
  const aliasIndices = locHeaders.map((h, i) => (h.includes("表記揺れ") || h.includes("クリニックNo")) ? i : -1).filter(i => i !== -1);

  const normalizeAddrText = (str) => {
    if (!str) return "";
    let s = String(str)
      .replace(/[【】\(（][^【】\(（\)）]*(内科|小児科)[^【】\(（\)）]*[\)）]/g, "")
      .replace(/[【】\(（]?(一般内科|内科|小児科)[\)）]?/g, "")
      .replace(/[\s ]+/g, "")
      .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
      .replace(/[ヶｹ]/g, "ケ")
      .replace(/[−ー―‐]/g, "-");
    const kanjiMap = { '〇': 0, '零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
    s = s.replace(/[〇零一二三四五六七八九十]+(?=丁目|番|号)/g, match => {
      if (match === '十') return '10';
      let total = 0, cur = 0;
      for (const ch of match) {
        if (ch === '十') { total += (cur === 0 ? 1 : cur) * 10; cur = 0; }
        else { cur = kanjiMap[ch]; }
      }
      return String(total + cur);
    });
    return s;
  };

  const formalList = [];
  const nameToFormalMap = new Map();
  const addressEntries = [];

  for (let i = 1; i < locData.length; i++) {
    const formal = idxFormal !== -1 ? String(locData[i][idxFormal]).replace(/[\s ]+/g, '') : "";
    if (!formal) continue;
    formalList.push({ formal, key: formal.replace(/[ヶｹ]/g, 'ケ') });

    const cleanKey = (v) => String(v).replace(/[\s 【】]+/g, '').replace(/[ヶｹ]/g, 'ケ');
    nameToFormalMap.set(formal, formal);
    nameToFormalMap.set(cleanKey(formal), formal);

    aliasIndices.forEach(idx => {
      const alias = String(locData[i][idx]).replace(/[\s ]+/g, '');
      if (alias) {
        nameToFormalMap.set(alias, formal);
        nameToFormalMap.set(cleanKey(alias), formal);
      }
    });

    if (idxAddress !== -1 && locData[i][idxAddress]) {
      const rawAddr = String(locData[i][idxAddress]).trim();
      const normAddr = normalizeAddrText(rawAddr);
      const hyphenAddr = normAddr.replace(/(\d+)丁目(\d+)?/g, (_, c1, c2) => c2 ? `${c1}-${c2}` : `${c1}-`);
      const townMatch = normAddr.match(/^(.+?[都道府県].+?[市区町村].+?\d+丁目)/);
      const townPrefix = townMatch ? townMatch[1] : "";
      if (normAddr) addressEntries.push({ formal, rawAddr, normAddr, hyphenAddr, townPrefix });
    }
  }
  formalList.sort((a, b) => b.key.length - a.key.length);

  const identifyClinic = (rawInput) => {
    const raw = String(rawInput || "").trim();
    if (!raw) return { resolved: false, formal: "" };
    if (EXCLUDED_NON_CLINICS.some(ex => raw.includes(ex))) {
      return { resolved: false, formal: raw };
    }

    const stripped = raw
      .replace(/キャップスクリニック|CAPSクリニック|CAPS|クリニック/gi, "")
      .replace(/[【】\(（][^【】\(（\)）]*(内科|小児科)[^【】\(（\)）]*[\)）]/g, "")
      .replace(/[【】\(（]?(一般内科|内科|小児科)[\)）]?/g, "")
      .replace(/[\s ]+/g, "");
    const rawClean = raw.replace(/[\s 【】]+/g, "");

    // 1. 正規記載・表記揺れ完全一致
    for (const c of [rawClean, rawClean.replace(/[ヶｹ]/g, "ケ"), stripped, stripped.replace(/[ヶｹ]/g, "ケ")]) {
      if (nameToFormalMap.has(c)) return { resolved: true, formal: nameToFormalMap.get(c) };
    }

    // 2. 住所照合（正規化・丁目ハイフン吸収・町名＋丁目一致）
    const normRaw = normalizeAddrText(raw);
    const hyphenRaw = normRaw.replace(/(\d+)丁目(\d+)?/g, (_, c1, c2) => c2 ? `${c1}-${c2}` : `${c1}-`);
    const addrHit = addressEntries.find(e =>
      e.normAddr.includes(normRaw) || normRaw.includes(e.normAddr) ||
      e.hyphenAddr.includes(hyphenRaw) || hyphenRaw.includes(e.hyphenAddr) ||
      (e.townPrefix && normRaw.includes(e.townPrefix))
    );
    if (addrHit) return { resolved: true, formal: addrHit.formal };

    // 3. 商業施設名付き等の部分拠点名一致（文字数降順）
    const strippedKe = stripped.replace(/[ヶｹ]/g, "ケ");
    const partialHit = formalList.find(item => item.key.length >= 2 && strippedKe.includes(item.key));
    if (partialHit) return { resolved: true, formal: partialHit.formal };

    return { resolved: false, formal: raw };
  };

  const strictNormalize = (raw) => identifyClinic(raw).formal;
  const twoDocRequests = getCheckerTwoDoctorRequests(TWO_DOC_SS_ID, scanStartDate, strictNormalize);

  // --- 2. 期間ルール厳守のシフトデータ取得（当月・翌月=貼付用 / 翌々月以降=確定シフト / 募集） ---
  const actualShiftsMap = getCheckerActualShifts(pasteSs, shiftSs, scanStartDate, thresholdDate, strictNormalize);
  const confirmedShifts = []; // 1対1引き当て用の確定・実績シフト配列
  const clinicDateShifts = new Map(); // Wブッキング・ゴースト募集判定用マップ

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

      const covKey = isMultiDept ? `${shift.dateStr}_${shift.normClinic}_${shiftDept}` : `${shift.dateStr}_${shift.normClinic}`;
      if (!clinicDateShifts.has(covKey)) clinicDateShifts.set(covKey, []);
      clinicDateShifts.get(covKey).push(shift);

      if (shift.sourceSheet === "実績" || shift.sourceSheet === "確定") {
        confirmedShifts.push(shift);
      }
    });
  }

  // --- 3. 採用くん「処理済み」シートの読み込み ＆ 同一識別番号のグループ統合 ---
  let recruitSheet = null;
  recruitSs.getSheets().forEach(s => {
    if (s.getSheetId() == '1155146432' || s.getName().includes("処理済み")) recruitSheet = s;
  });
  if (!recruitSheet) return;

  const rData = recruitSheet.getDataRange().getDisplayValues();
  if (rData.length < 2) return;

  const rCols = getColumnIndices(rData[0], ['識別番号', '採用可否', '医師名', '勤務希望日', '拠点名', '診療科', '該当時間', '処理状況', '着信日', '対応時間', '対応者']);
  const sheetId = recruitSheet.getSheetId();

  const groupedAdoptions = new Map(); // dedupKey -> 統合された採用案件オブジェクト

  for (let i = 1; i < rData.length; i++) {
    const row = rData[i];
    const saiyoVal = rCols['採用可否'] !== -1 ? String(row[rCols['採用可否']]).trim() : "";
    if (saiyoVal !== '採用') continue;

    // キャンセル済み痕跡や列ズレ不採用痕跡がある行は安全のためスキップ
    const colIVal = String(row[8] || "").trim();
    const statusVal = rCols['処理状況'] !== -1 ? String(row[rCols['処理状況']]).trim() : "";
    const personVal = rCols['対応者'] !== -1 ? String(row[rCols['対応者']]).trim() : "";
    if (colIVal.includes("キャンセル") || statusVal.includes("キャンセル") || statusVal.includes("お断り") || statusVal.includes("掲載停止") || personVal === "不採用") {
      continue;
    }

    const dateObj = parseDateToSafeDateObj(row[rCols['勤務希望日']]);
    if (!dateObj || dateObj < scanStartDate) continue;

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

    // 同一識別番号（応募通知受信＋勤務確定など）を1案件に統合
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

  // --- 4. 2段階の厳格1対1引き当て（Pass 1: 識別番号最優先 ／ Pass 2: 同日＋正規拠点＋診療科＋医師名） ---
  const matchedShiftByGroup = new Map();

  // Pass 1: スタッフコメント内の「識別番号(uKey)」による確定引き当て（時間・日程変更にも対応）
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
      // 同日同拠点で2枠ある場合に備え、時間帯が重なるシフトを優先的に引き当てる
      let hit = null;
      if (!isNaN(rec.sMin) && !isNaN(rec.eMin)) {
        hit = candidates.find(s => !isNaN(s.startMin) && !isNaN(s.endMin) && s.startMin < rec.eMin && s.endMin > rec.sMin);
      }
      if (!hit) hit = candidates[0];

      hit.usedByRecruitKey = rec.dedupKey;
      matchedShiftByGroup.set(rec.dedupKey, hit);
    }
  });

  // --- 5. 各採用案件の監査判定（①未反映 ／ ②紹介枠Wブッキング ／ ③募集枠消し忘れ） ---
  groupedAdoptions.forEach(rec => {
    // 当日着信・当日対応の未来シフトは入力猶予としてスキップ
    if (rec.isRecentlyProcessed && rec.dateObj > today) return;

    const stableIdPart = rec.uKey ? rec.uKey : `${rec.sStr}-${rec.eStr}`;
    const matchedShift = matchedShiftByGroup.get(rec.dedupKey);

    // ① 反映漏れ（シフト未登録）チェック
    if (!matchedShift) {
      const targetUrl = `https://docs.google.com/spreadsheets/d/${RECRUIT_MASTER_ID}/edit#gid=${sheetId}&range=${rec.latestSheetRow}:${rec.latestSheetRow}`;
      const sameDayOtherShifts = confirmedShifts.filter(s => s.dateStr === rec.dateStr && s.doctorName === rec.docClean && !s.isLeaveOrAbsent);
      const otherHint = sameDayOtherShifts.length > 0
        ? `\n※同日に別シフト（${sameDayOtherShifts.map(s => `${s.normClinic} ${s.startStr}-${s.endStr}`).join(", ")}）あり。拠点変更・案件取り直しの場合は旧行を「不採用」にしてください。`
        : "";

      const actionMsg = `採用くんで採用となっていますが、確定シフトに登録がありません。\n確認してください。\n応募取り下げ・キャンセルの場合は、採用可否のステータスを不採用にしてください。${otherHint}\n\n詳細(採用くん): ${targetUrl}`;
      const legacyPrefix = `採用漏れ_${rec.dateStr}_${rec.normClinic}_${rec.docClean}`;
      const uniqueId = `${legacyPrefix}_${stableIdPart}`;

      processLagError(
        "採用枠 未反映", rec.displayDate, rec.clinicLabel, rec.dept, rec.doctor,
        "スポット", `${rec.sStr}-${rec.eStr}`, "シフト未作成",
        actionMsg, uniqueId, legacyPrefix, rec.rowIndices, targetUrl
      );
      return;
    }

    // 過去日程、または欠勤シフト化されている場合はWブッキング・募集残存チェックは不要
    if (rec.dateObj < today || matchedShift.isLeaveOrAbsent) return;

    // 拠点変更等で別拠点・別日程のシフトに識別番号一致していた場合は、元の採用くん拠点での重複・募集残存を誤検知しないようスキップ
    if (matchedShift.dateStr !== rec.dateStr || matchedShift.normClinic !== rec.normClinic) return;

    const covKey = rec.isMultiDept ? `${rec.dateStr}_${rec.normClinic}_${rec.dept}` : `${rec.dateStr}_${rec.normClinic}`;
    const dayClinicShifts = clinicDateShifts.get(covKey) || [];
    const checkStartMin = !isNaN(matchedShift.startMin) ? matchedShift.startMin : rec.sMin;
    const checkEndMin = !isNaN(matchedShift.endMin) ? matchedShift.endMin : rec.eMin;

    if (isNaN(checkStartMin) || isNaN(checkEndMin)) return;

    // ② 紹介枠 Wブッキング チェック（2診要望一覧の正規キー形式 _NAME_ / _ID_ に対応）
    const hasTwoDocReq = (
      twoDocRequests.has(`${rec.dateStr}_NAME_${rec.normClinic}`) ||
      (matchedShift.clinicId && twoDocRequests.has(`${rec.dateStr}_ID_${matchedShift.clinicId}`))
    );

    if (!hasTwoDocReq) {
      const overlappingOthers = dayClinicShifts.filter(s =>
        (s.sourceSheet === "確定" || s.sourceSheet === "実績") &&
        !s.isLeaveOrAbsent &&
        s.doctorName !== rec.docClean &&
        s.startMin < checkEndMin && s.endMin > checkStartMin
      );

      if (overlappingOthers.length > 0) {
        const otherNames = Array.from(new Set(overlappingOthers.map(s => s.doctorName))).join(", ");
        const actionMsg = `採用くんの決定枠ですが、シフト上で別の医師（${otherNames}先生）がアサインされています。\n許容される２診であればA列にチェックを。ミスの場合は対応をお願いします。`;
        const legacyPrefix = `紹介重複_${rec.dateStr}_${rec.normClinic}_${rec.docClean}`;
        const uniqueId = `${legacyPrefix}_${stableIdPart}`;

        processLagError(
          "紹介枠 Wブッキング", rec.displayDate, rec.clinicLabel, rec.dept, rec.doctor,
          "スポット", `${rec.sStr}-${rec.eStr}`, `重複: ${otherNames}`,
          actionMsg, uniqueId, legacyPrefix, rec.rowIndices
        );
      }
    }

    // ③ ゴースト募集（募集枠 消し忘れ）チェック
    const remainingRecruits = dayClinicShifts.filter(s =>
      s.sourceSheet === "募集" &&
      s.startMin < checkEndMin && s.endMin > checkStartMin
    );

    if (remainingRecruits.length > 0) {
      const actionMsg = `採用くんで決定済みですが、同時間帯の募集枠が残っています。他媒体からの重複応募を防ぐため、該当の募集枠を取り下げ（クローズ）してください。`;
      const legacyPrefix = `募集残存_${rec.dateStr}_${rec.normClinic}_${rec.docClean}`;
      const uniqueId = `${legacyPrefix}_${stableIdPart}`;

      processLagError(
        "募集枠 消し忘れ", rec.displayDate, rec.clinicLabel, rec.dept, rec.doctor,
        "スポット", `${rec.sStr}-${rec.eStr}`, "募集枠が残存",
        actionMsg, uniqueId, legacyPrefix, rec.rowIndices
      );
    }
  });

  // --- 6. 最新の検知リスト（タイムラグ記録）をシステムシートに保存 ---
  systemSheet.getRange(1, 1).setValue(JSON.stringify(currentDelayData));

  // --- 7. アラートリストへの出力 ---
  const alertSheet = setupSheet(ACTIVE_SS, "アラートリスト", ALERT_HEADERS);
  if (errorValues.length > 0) {
    const combined = errorValues.map((val, i) => ({ val, bg: errorBackgrounds[i], rt: errorRichTexts[i] }));
    combined.sort((a, b) => a.val[11].localeCompare(b.val[11]));

    const sortedValues = combined.map(item => item.val);
    const sortedBackgrounds = combined.map(item => item.bg);
    const sortedRichTexts = combined.map(item => [item.rt]);

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

    // J列（10列目・対応指示）をリッチテキスト（部分リンク）で上書き
    alertSheet.getRange(startRow, 10, sortedRichTexts.length, 1).setRichTextValues(sortedRichTexts);

    Logger.log(`✅ ${sortedValues.length}件のエラー(採用くん監査)を末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規エラーはありませんでした。");
  }

  Logger.log(`⏱ Phase 4 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}