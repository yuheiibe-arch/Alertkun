/**
 * ==========================================
 * アラートリストの転記（アーカイブ）処理
 * ※ A列にチェックを入れるとアーカイブシートへ自動移動します
 * （★リンク等のリッチテキストや背景色も保持して移動します）
 * ==========================================
 */
function processAlertArchiveTrigger(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  
  if (sheet.getName() !== "アラートリスト") return;
  
  const col = e.range.getColumn();
  const row = e.range.getRow();
  
  // 1列目（転記チェックボックス）が編集され、かつチェックされた場合
  if (col === 1 && row > 1 && e.value === "TRUE") {
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
    const sourceRange = sheet.getRange(row, 1, 1, lastCol);
    
    // アーカイブシートの最終行の次を取得
    const targetRow = archiveSheet.getLastRow() + 1;
    const targetRange = archiveSheet.getRange(targetRow, 1, 1, lastCol);
    
    // ★値だけでなく、リッチテキスト（URLリンク）や背景色もすべてそのままコピーする
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
    // アラートリストから元の行を削除
    sheet.deleteRow(row);
  }
}

function onEdit(e) {
  processAlertArchiveTrigger(e);
}