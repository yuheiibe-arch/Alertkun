/**
 * ============================================================================
 * 【Phase4-A_採用監査_基盤・拠点同定】
 * ・エントリーポイント runShiftCheckerPhase4（master.gs から呼び出し）
 * ・アーカイブ済みキー互換照合 ＆ 2日連続検知（SystemData_Phase4Lag）状態管理
 * ・拠点名・エムスリー住所表記（漢数字/全角数字/一般内科/施設名付き等）の100%正規同定エンジン
 * ============================================================================
 */

function runShiftCheckerPhase4() {
  const startTime = Date.now();
  Logger.log("=== Phase 4: 採用くん連携監査 実行開始 ===");

  const ACTIVE_SS = SpreadsheetApp.getActiveSpreadsheet();
  // ドライラン（仮想シート）実行中かどうかを自動判定（ドライラン中は外部シートへの書き込みを行わない）
  const isDryRun = (typeof ACTIVE_SS.getSheets !== 'function');

  const PASTE_MASTER_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA';
  const SHIFT_MASTER_ID = '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co';
  const LOC_MASTER_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  const RECRUIT_MASTER_ID = '1NfSmzERA-ee-8ToDYSwL_PYkxvMSxRv5Y_CehfeAQcQ';
  const TWO_DOC_SS_ID = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk';

  let pasteSs, shiftSs, locSs, recruitSs;
  try {
    pasteSs = getCachedSs(PASTE_MASTER_ID);
    shiftSs = getCachedSs(SHIFT_MASTER_ID);
    locSs = getCachedSs(LOC_MASTER_ID);
    recruitSs = getCachedSs(RECRUIT_MASTER_ID);
  } catch (e) {
    Logger.log("❌ マスタ読み込みエラー: " + e.message);
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const scanStartDate = new Date(today.getTime());
  scanStartDate.setDate(scanStartDate.getDate() - 30);
  const thresholdDate = new Date(today.getFullYear(), today.getMonth() + 2, 1);

  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  const EXCLUDED_NON_CLINICS = ['院外勤務（小児科）', '嘱託医業務', '医師会業務', '【関東】バックアップシフト', '有給', '有休', '欠勤'];

  // 1. アーカイブ互換判定 ＆ 2日連続検知（Lag）管理の初期化
  const state = initPhase4ArchiveAndLagState_(ACTIVE_SS, ALERT_HEADERS, toYYYYMMDD(today));

  // 2. 100%拠点同定エンジンの構築 ＆ 2診要望一覧の取得
  const clinicEngine = buildPhase4ClinicIdentifier_(locSs, EXCLUDED_NON_CLINICS);
  const twoDocRequests = getCheckerTwoDoctorRequests(TWO_DOC_SS_ID, scanStartDate, clinicEngine.strictNormalize);

  // 3. 期間ルール厳守のシフトデータ取得・構造化 [Phase4-B]
  const shiftPack = preparePhase4ShiftStructures_(pasteSs, shiftSs, scanStartDate, thresholdDate, clinicEngine.strictNormalize);

  // 4. 採用くん「処理済み」シート読込 ＆ 同一識別番号のグループ統合 [Phase4-B]
  const recruitPack = loadAndGroupPhase4Adoptions_(recruitSs, today, clinicEngine.identifyClinic);
  if (!recruitPack) return;

  // 5. 3段階の厳格引き当て（識別番号 → 同日同拠点 → 同日同時間帯の拠点移動・旧行判定） [Phase4-B]
  const matchResult = matchPhase4AdoptionsToShifts_(recruitPack.groupedAdoptions, shiftPack.confirmedShifts);

  // 5.5 拠点移動・求人取り直し旧行を「処理済み」シートへ自動記載 [Phase4-B]
  applyPhase4RecruitSheetUpdates_(recruitPack, matchResult.sheetUpdates, isDryRun);

  // 6. 各採用案件の3種監査判定（①未反映 ／ ②紹介枠Wブッキング ／ ③募集枠消し忘れ） [Phase4-C]
  evaluatePhase4AdoptionAudits_({
    groupedAdoptions: recruitPack.groupedAdoptions,
    sheetId: recruitPack.sheetId,
    recruitMasterId: RECRUIT_MASTER_ID,
    today: today,
    matchedShiftByGroup: matchResult.matchedShiftByGroup,
    confirmedShifts: shiftPack.confirmedShifts,
    clinicDateShifts: shiftPack.clinicDateShifts,
    twoDocRequests: twoDocRequests,
    processLagError: state.processLagError
  });

  // 7. タイムラグ記録の保存 ＆ アラートリストへの出力 [Phase4-C]
  finalizePhase4Output_(state, ALERT_HEADERS.length, startTime);
}

function initPhase4ArchiveAndLagState_(activeSs, alertHeaders, todayYYYYMMDD) {
  const errorValues = [];
  const errorBackgrounds = [];
  const errorRichTexts = [];

  const archivedIds = new Set();
  const legacyArchivedPrefixes = new Set();
  const archiveData = setupSheet(activeSs, "アーカイブ", alertHeaders).getDataRange().getValues();
  for (let i = 1; i < archiveData.length; i++) {
    const uid = String(archiveData[i][11] || "").trim();
    if (!uid) continue;
    archivedIds.add(uid);
    const legacyMatch = uid.match(/^((?:採用漏れ|紹介重複|募集残存)_\d{4}\/\d{2}\/\d{2}_.+_[^_]+)_\d+$/);
    if (legacyMatch) {
      legacyArchivedPrefixes.add(legacyMatch[1]);
    }
  }

  const alertSheet = setupSheet(activeSs, "アラートリスト", alertHeaders);
  const alertData = alertSheet.getDataRange().getValues();
  for (let i = 1; i < alertData.length; i++) {
    if (alertData[i][11]) archivedIds.add(String(alertData[i][11]).trim());
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
      archivedIds.add(uniqueId);
    }
  };

  const SYSTEM_SHEET_NAME = "SystemData_Phase4Lag";
  let systemSheet = activeSs.getSheetByName(SYSTEM_SHEET_NAME);
  if (!systemSheet) {
    systemSheet = activeSs.insertSheet(SYSTEM_SHEET_NAME);
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

  const processLagError = (type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId, legacyPrefix, rowIndices, targetUrl = null, forceImmediate = false) => {
    const firstDetectedDate = findPreviousDetectedDate(uniqueId, legacyPrefix, rowIndices);
    if (forceImmediate) {
      currentDelayData[uniqueId] = firstDetectedDate || "AGED";
      addError(type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId, legacyPrefix, rowIndices, targetUrl);
      return;
    }

    if (!firstDetectedDate) {
      currentDelayData[uniqueId] = todayYYYYMMDD;
    } else {
      currentDelayData[uniqueId] = firstDetectedDate;
      if (firstDetectedDate !== todayYYYYMMDD) {
        addError(type, displayDate, clinic, dept, doctor, empType, workTime, errorDetail, actionMsg, uniqueId, legacyPrefix, rowIndices, targetUrl);
      }
    }
  };

  return {
    alertSheet,
    systemSheet,
    currentDelayData,
    errorValues,
    errorBackgrounds,
    errorRichTexts,
    processLagError
  };
}

function buildPhase4ClinicIdentifier_(locSs, excludedNonClinics) {
  const locData = getCachedDisplayValues(locSs.getSheetByName("拠点名"));
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
    if (excludedNonClinics.some(ex => raw.includes(ex))) {
      return { resolved: false, formal: raw };
    }

    const stripped = raw
      .replace(/キャップスクリニック|CAPSクリニック|CAPS|クリニック/gi, "")
      .replace(/[【】\(（][^【】\(（\)）]*(内科|小児科)[^【】\(（\)）]*[\)）]/g, "")
      .replace(/[【】\(（]?(一般内科|内科|小児科)[\)）]?/g, "")
      .replace(/[\s ]+/g, "");
    const rawClean = raw.replace(/[\s 【】]+/g, "");

    for (const c of [rawClean, rawClean.replace(/[ヶｹ]/g, "ケ"), stripped, stripped.replace(/[ヶｹ]/g, "ケ")]) {
      if (nameToFormalMap.has(c)) return { resolved: true, formal: nameToFormalMap.get(c) };
    }

    const normRaw = normalizeAddrText(raw);
    const hyphenRaw = normRaw.replace(/(\d+)丁目(\d+)?/g, (_, c1, c2) => c2 ? `${c1}-${c2}` : `${c1}-`);
    const addrHit = addressEntries.find(e =>
      e.normAddr.includes(normRaw) || normRaw.includes(e.normAddr) ||
      e.hyphenAddr.includes(hyphenRaw) || hyphenRaw.includes(e.hyphenAddr) ||
      (e.townPrefix && normRaw.includes(e.townPrefix))
    );
    if (addrHit) return { resolved: true, formal: addrHit.formal };

    const strippedKe = stripped.replace(/[ヶｹ]/g, "ケ");
    const partialHit = formalList.find(item => item.key.length >= 2 && strippedKe.includes(item.key));
    if (partialHit) return { resolved: true, formal: partialHit.formal };

    return { resolved: false, formal: raw };
  };

  return {
    identifyClinic,
    strictNormalize: (raw) => identifyClinic(raw).formal
  };
}