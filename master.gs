/**
 * ==========================================
 * マスターコントロール (Phase 1〜4 + 依頼手当 独立バトンリレー実行)
 * ★UPDATE: 重複ストッパーを撤廃し、未対応エラー全件を毎日リマインドする仕様
 * ★UPDATE: オンメモリキャッシュによる高速化 ＆ Phase異常終了時のSlack警告通知を追加
 * ==========================================
 */
function runMasterShiftChecker() {
  const startTime = Date.now();
  Logger.log("=== 🚀 マスター監査プロセス 開始 ===");

  if (typeof clearCheckerRuntimeCache === "function") {
    clearCheckerRuntimeCache();
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ALERT_HEADERS = ["転記", "項目", "勤務日", "拠点名", "診療科", "医師名", "雇用区分", "勤務時間", "エラー箇所", "対応指示", "メモ", "ユニークキー"];
  const alertSheet = setupSheet(ss, "アラートリスト", ALERT_HEADERS);

  // 1. アラートリストの初期化 (2行目は空で残す)
  const maxRows = alertSheet.getMaxRows();
  if (maxRows > 2) {
    alertSheet.deleteRows(3, maxRows - 2);
  }
  if (alertSheet.getMaxRows() >= 2) {
    alertSheet.getRange(2, 1, 1, alertSheet.getMaxColumns()).clear();
  }

  const failedPhases = [];

  // 安全な実行のためのヘルパー関数
  const runSafe = (phaseFunc, phaseName) => {
    const pStart = Date.now();
    try {
      phaseFunc();
      SpreadsheetApp.flush();
      Utilities.sleep(300);
      Logger.log(`  ✔ ${phaseName} 完了 (${((Date.now() - pStart) / 1000).toFixed(1)}秒)`);
    } catch (e) {
      const errMsg = `${phaseName}: ${e.message}`;
      failedPhases.push(errMsg);
      Logger.log(`❌ ${phaseName} 実行中にエラーが発生しスキップしました: ` + e.message);
    }
  };

  // 各フェーズを独立して実行
  runSafe(runShiftCheckerPhase1, "Phase 1 (基本監査)");
  runSafe(runShiftCheckerPhase2, "Phase 2 (定期・常勤漏れ)");
  runSafe(runShiftCheckerPhase2_Recruit, "Phase 2.5 (募集枠・空き枠)");
  runSafe(runShiftCheckerPhase3, "Phase 3 (時給・有給監査)");
  runSafe(runShiftCheckerAllowance, "Phase 3.5 (依頼手当監査)");
  runSafe(runShiftCheckerPhase4, "Phase 4 (採用くん連携監査)");

  // 2. 実行後のリストから【すべての未対応エラー】を集計してSlack通知する
  const finalLastRow = alertSheet.getLastRow();
  let totalErrorsCount = 0;
  const breakdown = {};

  if (finalLastRow > 1) {
    const newData = alertSheet.getRange(2, 1, finalLastRow - 1, 12).getValues();
    newData.forEach(row => {
      const errName = row[1]; // B列(2列目)が項目名
      const uid = String(row[11]); // L列(12列目)

      if (uid) {
        totalErrorsCount++;
        if (errName) breakdown[errName] = (breakdown[errName] || 0) + 1;
      }
    });
  }

  // 3. 未対応エラーがある場合、または途中エラーが発生したPhaseがある場合にSlack送信
  if (totalErrorsCount > 0 || failedPhases.length > 0) {
    const sheetUrl = ss.getUrl() + "#gid=" + alertSheet.getSheetId();
    if (typeof sendSlackAlert === "function") {
      sendSlackAlert(totalErrorsCount, breakdown, sheetUrl, failedPhases);
    }
    Logger.log(`✅ Slackへ通知を送信しました (未対応:${totalErrorsCount}件 / 失敗Phase:${failedPhases.length}件)。`);
  } else {
    Logger.log("✅ 未対応エラー0件（全Phase正常完了）のため、Slack通知はスキップしました。");
  }

  Logger.log(`=== 🏁 マスター監査プロセス 完了 (${((Date.now() - startTime) / 1000).toFixed(1)}秒) ===`);
}