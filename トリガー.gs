/**
 * ==========================================
 * アラートリストの転記（アーカイブ）処理
 * ※ A列にチェックを入れるとアーカイブシートへ自動移動します
 * ★UPDATE: 連続クリックによる衝突を防ぐ「ロック機能（排他制御）」を追加
 * ★UPDATE: 削除時にシートがヘッダーのみになるエラーを防ぐ空行補完処理を追加
 * ==========================================
 */
function processAlertArchiveTrigger(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  
  if (sheet.getName() !== "アラートリスト") return;
  if (e.range.getColumn() > 1) return;
  
  const rowStart = e.range.getRow();
  const numRows = e.range.getNumRows();
  const rowEnd = rowStart + numRows - 1;
  
  if (rowEnd < 2) return;

  // ★追加：ロックを取得（他の処理が走っている間は最大10秒待機して順番を守る）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    // 10秒待っても前の処理が終わらなければ安全のために処理をキャンセル
    Logger.log("同時編集のロックがかかっています。少し待ってから再度チェックを入れてください。");
    return;
  }
  
  try {
    const ss = e.source;
    let archiveSheet = ss.getSheetByName("アーカイブ");
    
    // アーカイブシートが存在しなければ作成
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet("アーカイブ");
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
      archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight("bold");
      archiveSheet.setFrozenRows(1);
    }
    
    const lastCol = sheet.getLastColumn();
    
    // 行ズレ巻き込み削除を防ぐため、必ず「一番下の行から上に向かって」逆順で処理
    for (let r = rowEnd; r >= Math.max(2, rowStart); r--) {
      // その行が本当にチェックされているか、削除直前に再確認する
      const isChecked = sheet.getRange(r, 1).getValue();
      
      if (isChecked === true || isChecked === "TRUE") {
        const sourceRange = sheet.getRange(r, 1, 1, lastCol);
        
        const targetRow = archiveSheet.getLastRow() + 1;
        const targetRange = archiveSheet.getRange(targetRow, 1, 1, lastCol);
        
        // 値、リッチテキスト（URLリンク）、背景色などすべてそのままコピー
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        
        // アラートリストから元の行を削除
        sheet.deleteRow(r);

        // ★追加: 削除した結果、シートが1行(ヘッダーのみ)になった場合は空行を追加する
        if (sheet.getMaxRows() < 2) {
          sheet.insertRowsAfter(1, 1);
        }
      }
    }
  } catch (error) {
    Logger.log("エラーが発生しました: " + error.message);
  } finally {
    // ★追加：処理が終わったら必ずロックを解除し、次の処理に順番を譲る
    lock.releaseLock();
  }
}

function onEdit(e) {
  processAlertArchiveTrigger(e);
}