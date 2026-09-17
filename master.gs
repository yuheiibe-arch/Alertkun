/**
 * ==========================================
 * マスターコントロール (Phase 1〜4 + 依頼手当 独立バトンリレー実行)
 * ★UPDATE: 重複ストッパーを撤廃し、未対応エラー全件を毎日リマインドする仕様
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
  
  // 1. アラートリストの初期化 (2行目は空で残す)
  const maxRows = alertSheet.getMaxRows();
  if (maxRows > 2) {
    alertSheet.deleteRows(3, maxRows - 2); 
  }
  if (alertSheet.getMaxRows() >= 2) {
    alertSheet.getRange(2, 1, 1, alertSheet.getMaxColumns()).clear(); 
  }
  
  // 安全な実行と負荷対策（2秒待機）のためのヘルパー関数
  const runSafe = (phaseFunc, phaseName) => {
    try {
      phaseFunc();
      Utilities.sleep(2000); 
    } catch (e) {
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
      
      // ユニークキーが存在するもの（空行以外）は新規・既存問わずすべてカウント
      if (uid) {
        totalErrorsCount++;
        if (errName) breakdown[errName] = (breakdown[errName] || 0) + 1;
      }
    });
  }
  
  // 3. 未対応エラーがある場合のみSlack送信 (slack.gs の sendSlackAlert を呼び出し)
  if (totalErrorsCount > 0) {
    const sheetUrl = ss.getUrl() + "#gid=" + alertSheet.getSheetId();
    if (typeof sendSlackAlert === "function") {
      sendSlackAlert(totalErrorsCount, breakdown, sheetUrl);
    }
    Logger.log(`✅ Slackへ ${totalErrorsCount}件の未対応エラー通知を送信しました。`);
  } else {
    Logger.log("✅ 未対応エラー0件のため、Slack通知はスキップしました。");
  }
  
  Logger.log(`=== 🏁 マスター監査プロセス 完了 (${(Date.now() - startTime) / 1000}秒) ===`);
}