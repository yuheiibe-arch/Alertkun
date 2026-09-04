/**
 * ==========================================
 * マスターコントロール (Phase 1〜4 独立バトンリレー実行)
 * ==========================================
 */
function runMasterShiftChecker() {
  const startTime = Date.now();
  Logger.log("=== 🚀 マスター監査プロセス 開始 ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alertSheet = ss.getSheetByName("アラートリスト");
  if (!alertSheet) {
    Logger.log("❌ アラートリストが存在しません。各Phaseを単独実行して初期化してください。");
    return;
  }
  
  // 1. 現在アラートリストに表示されている「既存エラーのユニークキー」を記憶（Slackの重複通知防止のため）
  const lastRowBefore = alertSheet.getLastRow();
  const existingIds = new Set();
  if (lastRowBefore > 1) {
    const oldData = alertSheet.getRange(2, 1, lastRowBefore - 1, 12).getValues();
    oldData.forEach(row => {
      const uid = String(row[11]); // L列(12列目)がユニークキー
      if (uid) existingIds.add(uid);
    });
  }
  
  // 2. アラートリストの2行目以降を完全に削除（修正済みの古いエラーを消去するため）
  const maxRows = alertSheet.getMaxRows();
  if (maxRows > 1) {
    alertSheet.deleteRows(2, maxRows - 1);
  }
  
  // 安全な実行と負荷対策（2秒待機）のためのヘルパー関数
  const runSafe = (phaseFunc, phaseName) => {
    try {
      phaseFunc();
      Utilities.sleep(2000); // サーバーエラー(Service Spreadsheets failed)防止用
    } catch (e) {
      Logger.log(`❌ ${phaseName} 実行中にエラーが発生しスキップしました: ` + e.message);
    }
  };

  // 各フェーズを独立して実行（1つがコケても次へ進む）
  runSafe(runShiftCheckerPhase1, "Phase 1 (基本監査)");
  runSafe(runShiftCheckerPhase2, "Phase 2 (定期・常勤漏れ)");
  runSafe(runShiftCheckerPhase2_Recruit, "Phase 2.5 (募集枠・空き枠)");
  runSafe(runShiftCheckerPhase3, "Phase 3 (時給・有給監査)");
  runSafe(runShiftCheckerPhase4, "Phase 4 (採用くん連携監査)");
  
  // 3. 実行後のリストから「今回初めて出現したエラー」だけを抽出してSlack通知する
  const finalLastRow = alertSheet.getLastRow();
  let newErrorsCount = 0;
  const breakdown = {};
  
  if (finalLastRow > 1) {
    const newData = alertSheet.getRange(2, 1, finalLastRow - 1, 12).getValues();
    newData.forEach(row => {
      const errName = row[1]; // B列(2列目)が項目名
      const uid = String(row[11]); // L列(12列目)
      
      // 実行前に存在しなかった新しいユニークキーの場合のみ「新規」としてカウント
      if (uid && !existingIds.has(uid)) {
        newErrorsCount++;
        if (errName) breakdown[errName] = (breakdown[errName] || 0) + 1;
      }
    });
  }
  
  // 新規エラーがあった場合のみ、内訳を集計してSlackヘルパーへ渡す
  if (newErrorsCount > 0) {
    const sheetUrl = ss.getUrl() + "#gid=" + alertSheet.getSheetId();
    if (typeof sendSlackAlert === "function") {
      sendSlackAlert(newErrorsCount, breakdown, sheetUrl);
    }
    Logger.log(`✅ Slackへ ${newErrorsCount}件の新規エラー通知を送信しました。`);
  } else {
    Logger.log("✅ 新規エラー0件のため、Slack通知はスキップしました。");
  }
  
  Logger.log(`=== 🏁 マスター監査プロセス 完了 (${(Date.now() - startTime) / 1000}秒) ===`);
}