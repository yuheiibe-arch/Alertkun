/**
 * ==========================================
 * アラートリストの転記（アーカイブ）処理
 * ※ A列にチェックを入れるとアーカイブシートへ自動移動します
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
    
    // 該当行をコピーして削除
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    archiveSheet.appendRow(rowData);
    sheet.deleteRow(row);
  }
}

// ※ すでに他のonEdit関数が存在しない場合は、以下の名前をonEditに変更するか、
// 既存のonEdit関数の中に processAlertArchiveTrigger(e); を追記してください。
function onEdit(e) {
  processAlertArchiveTrigger(e);
}