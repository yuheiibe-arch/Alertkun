/**
 * ============================================================================
 * 【Phase4-B_採用監査_データ統合・引当】
 * ・シフトデータの診療科分離・1対1引き当て用リスト構築
 * ・採用くん「処理済み」シート（本日以降）の読込 ＆ 同一識別番号のグループ統合
 * ・3段階の厳格引き当て：
 *    Pass 1: 識別番号一致（別拠点移動時はJ列メモ自動追記）
 *    Pass 2: 同日＋正規拠点＋診療科＋医師名 一致
 *    Pass 3A: 事務局依頼での拠点移動（同日＋同時間帯重複＋未引当シフト → 採用維持のままJ列へ移動先記載）
 *    Pass 3B: 拠点変更・求人取り直し旧行（同日＋同時間帯重複＋同社別採用行で引当済 → 不採用＆旧行記載）
 * ============================================================================
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
    '紹介会社名', '識別番号', '採用可否', '医師名', '勤務希望日', '拠点名', '診療科', '該当時間', '処理状況', '着信日', '対応時間', '対応者'
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
    if (colIVal.includes("キャンセル") || statusVal.includes("キャンセル") || statusVal.includes("お断り") || statusVal.includes("掲載停止") || statusVal.includes("旧行") || personVal === "不採用") {
      continue;
    }

    const dateObj = parseDateToSafeDateObj(row[rCols['勤務希望日']]);
    if (!dateObj || dateObj < today) continue;

    let isRecentlyProcessed = false;
    let actionDateObj = null;

    if (rCols['着信日'] !== -1 && row[rCols['着信日']]) {
      const receivedDate = parseDateToSafeDateObj(row[rCols['着信日']]);
      if (receivedDate) {
        actionDateObj = receivedDate;
        if (receivedDate.getTime() >= today.getTime()) isRecentlyProcessed = true;
      }
    }
    if (rCols['対応時間'] !== -1 && row[rCols['対応時間']]) {
      const processedDate = parseDateToSafeDateObj(row[rCols['対応時間']]);
      if (processedDate) {
        if (!actionDateObj || processedDate.getTime() > actionDateObj.getTime()) {
          actionDateObj = processedDate;
        }
        if (processedDate.getTime() >= today.getTime()) isRecentlyProcessed = true;
      }
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

    let agency = rCols['紹介会社名'] !== -1 ? String(row[rCols['紹介会社名']]).trim() : "";
    if (!agency && /^C\d{8,}/.test(uKey)) agency = "エムスリー";

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
        agency,
        rowIndices: [i],
        latestSheetRow: i + 1,
        isRecentlyProcessed,
        actionDateObj,
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
        eMin,
        supersededByShift: null
      });
    } else {
      const existing = groupedAdoptions.get(dedupKey);
      existing.rowIndices.push(i);
      existing.latestSheetRow = i + 1;
      if (!existing.agency && agency) existing.agency = agency;
      if (isRecentlyProcessed) existing.isRecentlyProcessed = true;
      if (actionDateObj && (!existing.actionDateObj || actionDateObj.getTime() > existing.actionDateObj.getTime())) {
        existing.actionDateObj = actionDateObj;
      }
    }
  }

  return {
    recruitSheet,
    rData,
    statusColIdx: rCols['処理状況'],
    saiyoColIdx: rCols['採用可否'],
    sheetId,
    groupedAdoptions
  };
}

function matchPhase4AdoptionsToShifts_(groupedAdoptions, confirmedShifts) {
  const matchedShiftByGroup = new Map();
  const sheetUpdates = []; // { rec, mode: 'RELINK_MOVE' | 'SUPERSEDED_OLD', newStatus }

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

      // 識別番号一致で別拠点へ移動していた場合も「処理済み」シートのJ列で移動先が分かるよう記録予約
      if (rec.clinicResolved && hit.normClinic && hit.normClinic !== rec.normClinic && !hit.isLeaveOrAbsent) {
        sheetUpdates.push({
          rec,
          mode: 'RELINK_MOVE',
          newStatus: `拠点移動済(→${hit.normClinic} ${hit.startStr}-${hit.endStr})`
        });
      }
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

  // Pass 3A: 事務局依頼での拠点移動（同日 ＋ 該当時間帯が重なる未引当の実シフトが別拠点にある場合 → 採用維持のまま移動済記載）
  groupedAdoptions.forEach(rec => {
    if (matchedShiftByGroup.has(rec.dedupKey)) return;
    if (isNaN(rec.sMin) || isNaN(rec.eMin)) return;

    const hit = confirmedShifts.find(s =>
      !s.usedByRecruitKey &&
      !s.isLeaveOrAbsent &&
      s.dateStr === rec.dateStr &&
      s.doctorName === rec.docClean &&
      !isNaN(s.startMin) && !isNaN(s.endMin) &&
      s.startMin < rec.eMin && s.endMin > rec.sMin
    );

    if (hit) {
      hit.usedByRecruitKey = rec.dedupKey;
      matchedShiftByGroup.set(rec.dedupKey, hit);
      sheetUpdates.push({
        rec,
        mode: 'RELINK_MOVE',
        newStatus: `拠点移動済(→${hit.normClinic} ${hit.startStr}-${hit.endStr})`
      });
    }
  });

  // Pass 3B: 拠点変更・求人取り直しによる「旧行」の自動整理（同日・同時間帯の実シフトが既に同社の新「採用」行に紐付いている場合）
  groupedAdoptions.forEach(rec => {
    if (matchedShiftByGroup.has(rec.dedupKey)) return;
    if (isNaN(rec.sMin) || isNaN(rec.eMin)) return;

    const claimedShift = confirmedShifts.find(s =>
      Boolean(s.usedByRecruitKey) &&
      !s.isLeaveOrAbsent &&
      s.dateStr === rec.dateStr &&
      s.doctorName === rec.docClean &&
      !isNaN(s.startMin) && !isNaN(s.endMin) &&
      s.startMin < rec.eMin && s.endMin > rec.sMin
    );

    if (claimedShift) {
      const ownerRec = groupedAdoptions.get(claimedShift.usedByRecruitKey);
      const isSameAgency = ownerRec && (!rec.agency || !ownerRec.agency || rec.agency === ownerRec.agency);

      if (isSameAgency) {
        rec.supersededByShift = claimedShift;
        const labelPrefix = (claimedShift.normClinic === rec.normClinic) ? '再登録・旧行' : '拠点変更・旧行';
        sheetUpdates.push({
          rec,
          mode: 'SUPERSEDED_OLD',
          newStatus: `${labelPrefix}(→${claimedShift.normClinic} ${claimedShift.startStr}-${claimedShift.endStr})`
        });
      }
    }
  });

  return { matchedShiftByGroup, sheetUpdates };
}

/**
 * 検知した「拠点移動済（採用維持）」および「拠点変更・再登録の旧行（不採用化）」を
 * 採用くんの「処理済み」シートへ自動反映するヘルパー関数
 */
function applyPhase4RecruitSheetUpdates_(recruitPack, sheetUpdates, isDryRun) {
  if (!sheetUpdates || sheetUpdates.length === 0) return;

  const { recruitSheet, rData, statusColIdx, saiyoColIdx } = recruitPack;
  if (!recruitSheet || statusColIdx === -1 || saiyoColIdx === -1) return;

  const totalCols = recruitSheet.getLastColumn();
  let updatedRowCount = 0;

  sheetUpdates.forEach(item => {
    const { rec, mode, newStatus } = item;

    rec.rowIndices.forEach(idx => {
      const rowNum = idx + 1;
      const curStatus = String(rData[idx][statusColIdx] || "").trim();
      const curSaiyo = String(rData[idx][saiyoColIdx] || "").trim();

      if (mode === 'RELINK_MOVE') {
        // 既に同じ移動先メモが書かれていれば書き込みスキップ（無駄なAPI通信ゼロ）
        if (curStatus === newStatus) return;

        if (isDryRun) {
          Logger.log(`  🔄 [処理済み自動記載プレビュー(採用維持)] 行${rowNum}: ${rec.dateStr} ${rec.doctor} (${rec.normClinic}) -> J列="${newStatus}" / O列="採用(維持)"`);
        } else {
          try {
            recruitSheet.getRange(rowNum, statusColIdx + 1).setValue(newStatus).setBackground('#cfe2f3');
            updatedRowCount++;
          } catch (e) {
            Logger.log(`⚠️ 処理済みシート更新スキップ(行${rowNum}): ${e.message}`);
          }
        }
      } else if (mode === 'SUPERSEDED_OLD') {
        if (curSaiyo === '不採用' && curStatus === newStatus) return;

        if (isDryRun) {
          Logger.log(`  🔄 [処理済み自動切替プレビュー(旧行整理)] 行${rowNum}: ${rec.dateStr} ${rec.doctor} (${rec.normClinic} ID:${rec.uKey}) -> J列="${newStatus}" / O列="不採用"`);
        } else {
          try {
            recruitSheet.getRange(rowNum, statusColIdx + 1).setValue(newStatus);
            recruitSheet.getRange(rowNum, saiyoColIdx + 1).setValue('不採用');
            recruitSheet.getRange(rowNum, 1, 1, totalCols).setBackground('#cfe2f3');
            updatedRowCount++;
          } catch (e) {
            Logger.log(`⚠️ 処理済みシート旧行更新スキップ(行${rowNum}): ${e.message}`);
          }
        }
      }
    });
  });

  if (!isDryRun && updatedRowCount > 0) {
    Logger.log(`✅ 採用くん「処理済み」シートの拠点移動・旧行ステータスを ${updatedRowCount} 行自動更新しました。`);
  }
}