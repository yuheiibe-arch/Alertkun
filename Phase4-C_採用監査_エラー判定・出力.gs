/**
 * ============================================================================
 * 【Phase4-C_採用監査_エラー判定・出力】
 * ・① 採用枠 未反映（対応日から1日以上経過、または当日勤務なら猶予なしで即時出力）
 * ・② 紹介枠 Wブッキング（夜間帯 20:00-21:00 の重複のみ検知・自己重複防止・2診要望除外付き）
 * ・③ 募集枠 消し忘れ（非公開/未掲載クローズ枠の除外付き）
 * ・SystemData_Phase4Lag への保存 ＆ アラートリストへのソート出力
 * ============================================================================
 */

function evaluatePhase4AdoptionAudits_(ctx) {
  const {
    groupedAdoptions,
    sheetId,
    recruitMasterId,
    today,
    matchedShiftByGroup,
    confirmedShifts,
    clinicDateShifts,
    twoDocRequests,
    processLagError
  } = ctx;

  // Phase 1 と同じく、原則2診にしない夜間帯（20:00〜21:00）をWブッキング監査対象とする
  const NIGHT_CHECK_START = 20 * 60;
  const NIGHT_CHECK_END = 21 * 60;

  groupedAdoptions.forEach(rec => {
    // 拠点変更・求人取り直しの旧行として自動整理された案件はエラー対象外
    if (rec.supersededByShift) return;

    if (rec.isRecentlyProcessed && rec.dateObj > today) return;

    // 採用くんの「対応時間（または着信日）」が昨日以前、または当日勤務の場合は猶予なしで即時出力
    const isAlreadyAged = Boolean(
      rec.dateObj.getTime() <= today.getTime() ||
      (rec.actionDateObj && (today.getTime() - rec.actionDateObj.getTime()) >= 24 * 60 * 60 * 1000)
    );

    const stableIdPart = rec.uKey ? rec.uKey : `${rec.sStr}-${rec.eStr}`;
    const matchedShift = matchedShiftByGroup.get(rec.dedupKey);

    // ① 反映漏れ（シフト未登録）チェック
    if (!matchedShift) {
      const targetUrl = `https://docs.google.com/spreadsheets/d/${recruitMasterId}/edit#gid=${sheetId}&range=${rec.latestSheetRow}:${rec.latestSheetRow}`;
      const sameDayOtherShifts = confirmedShifts.filter(s =>
        s.dateStr === rec.dateStr && s.doctorName === rec.docClean && !s.isLeaveOrAbsent
      );
      const otherHint = sameDayOtherShifts.length > 0
        ? `\n※同日に別時間帯・別社シフト（${sameDayOtherShifts.map(s => `${s.normClinic} ${s.startStr}-${s.endStr}`).join(", ")}）あり。取り下げの場合は「不採用」にしてください。`
        : "";

      const actionMsg = `採用くんで採用となっていますが、確定シフトに登録がありません。\n確認してください。\n応募取り下げ・キャンセルの場合は、採用可否のステータスを不採用にしてください。${otherHint}\n\n詳細(採用くん): ${targetUrl}`;
      const legacyPrefix = `採用漏れ_${rec.dateStr}_${rec.normClinic}_${rec.docClean}`;
      const uniqueId = `${legacyPrefix}_${stableIdPart}`;

      processLagError(
        "採用枠 未反映", rec.displayDate, rec.clinicLabel, rec.dept, rec.doctor,
        "スポット", `${rec.sStr}-${rec.eStr}`, "シフト未作成",
        actionMsg, uniqueId, legacyPrefix, rec.rowIndices, targetUrl, isAlreadyAged
      );
      return;
    }

    if (matchedShift.isLeaveOrAbsent) return;
    if (matchedShift.dateStr !== rec.dateStr || matchedShift.normClinic !== rec.normClinic) return;

    const covKey = rec.isMultiDept
      ? `${rec.dateStr}_${rec.normClinic}_${rec.dept}`
      : `${rec.dateStr}_${rec.normClinic}`;
    const dayClinicShifts = clinicDateShifts.get(covKey) || [];
    const checkStartMin = !isNaN(matchedShift.startMin) ? matchedShift.startMin : rec.sMin;
    const checkEndMin = !isNaN(matchedShift.endMin) ? matchedShift.endMin : rec.eMin;

    if (isNaN(checkStartMin) || isNaN(checkEndMin)) return;

    // ② 紹介枠 Wブッキング チェック（夜間帯 20:00〜21:00 に重なっている場合のみ検知）
    const hasTwoDocReq = (
      twoDocRequests.has(`${rec.dateStr}_NAME_${rec.normClinic}`) ||
      (matchedShift.clinicId && twoDocRequests.has(`${rec.dateStr}_ID_${matchedShift.clinicId}`))
    );

    const overlapNightStart = Math.max(checkStartMin, NIGHT_CHECK_START);
    const overlapNightEnd = Math.min(checkEndMin, NIGHT_CHECK_END);

    if (!hasTwoDocReq && overlapNightStart < overlapNightEnd) {
      const overlappingOthers = dayClinicShifts.filter(s => {
        if (s.sourceSheet !== "確定" && s.sourceSheet !== "実績") return false;
        if (s.isLeaveOrAbsent || s === matchedShift) return false;
        if (s.doctorName === rec.docClean || s.doctorName === matchedShift.doctorName) return false;
        if (matchedShift.doctorId && s.doctorId && s.doctorId === matchedShift.doctorId) return false;
        return s.startMin < overlapNightEnd && s.endMin > overlapNightStart;
      });

      if (overlappingOthers.length > 0) {
        const otherNames = Array.from(new Set(overlappingOthers.map(s => s.doctorName))).join(", ");
        const actionMsg = `採用くんの決定枠ですが、夜間帯（20:00-21:00）に別の医師（${otherNames}先生）がアサインされています。\n許容される２診であればA列にチェックを。ミスの場合は対応をお願いします。`;
        const legacyPrefix = `紹介重複_${rec.dateStr}_${rec.normClinic}_${rec.docClean}`;
        const uniqueId = `${legacyPrefix}_${stableIdPart}`;

        processLagError(
          "紹介枠 Wブッキング", rec.displayDate, rec.clinicLabel, rec.dept, rec.doctor,
          "スポット", `${rec.sStr}-${rec.eStr}`, `夜間重複: ${otherNames}`,
          actionMsg, uniqueId, legacyPrefix, rec.rowIndices, null, isAlreadyAged
        );
      }
    }

    // ③ ゴースト募集（募集枠 消し忘れ）チェック（未掲載・非公開にクローズ済みの募集枠は除外）
    const remainingRecruits = dayClinicShifts.filter(s => {
      if (s.sourceSheet !== "募集") return false;
      const pStatus = s.publishStatus || "";
      if (pStatus.includes("未掲載") || pStatus.includes("非公開") || pStatus.includes("非掲載")) return false;
      return s.startMin < checkEndMin && s.endMin > checkStartMin;
    });

    if (remainingRecruits.length > 0) {
      const actionMsg = `採用くんで決定済みですが、同時間帯の募集枠が公開状態で残っています。他媒体からの重複応募を防ぐため、該当の募集枠を取り下げ（クローズ）してください。`;
      const legacyPrefix = `募集残存_${rec.dateStr}_${rec.normClinic}_${rec.docClean}`;
      const uniqueId = `${legacyPrefix}_${stableIdPart}`;

      processLagError(
        "募集枠 消し忘れ", rec.displayDate, rec.clinicLabel, rec.dept, rec.doctor,
        "スポット", `${rec.sStr}-${rec.eStr}`, "募集枠が残存",
        actionMsg, uniqueId, legacyPrefix, rec.rowIndices, null, isAlreadyAged
      );
    }
  });
}

function finalizePhase4Output_(state, headersLength, startTime) {
  const {
    alertSheet,
    systemSheet,
    currentDelayData,
    errorValues,
    errorBackgrounds,
    errorRichTexts
  } = state;

  systemSheet.getRange(1, 1).setValue(JSON.stringify(currentDelayData));

  if (errorValues.length > 0) {
    const combined = errorValues.map((val, i) => ({
      val,
      bg: errorBackgrounds[i],
      rt: errorRichTexts[i]
    }));
    combined.sort((a, b) => a.val[11].localeCompare(b.val[11]));

    const sortedValues = combined.map(item => item.val);
    const sortedBackgrounds = combined.map(item => item.bg);
    const sortedRichTexts = combined.map(item => [item.rt]);

    appendAlertRowsSafe(alertSheet, headersLength, sortedValues, sortedBackgrounds, sortedRichTexts);
    Logger.log(`✅ ${sortedValues.length}件のエラー(採用くん監査)を末尾に追記しました。`);
  } else {
    Logger.log("✅ 新規エラーはありませんでした。");
  }

  Logger.log(`⏱ Phase 4 処理時間: ${(Date.now() - startTime) / 1000}秒`);
}