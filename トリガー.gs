/**
 * ==========================================
 * アラートリストの転記（アーカイブ）処理
 * ※ A列にチェックを入れるとアーカイブシートへ自動移動します
 * ★UPDATE: 連続クリックによる衝突を防ぐ「ロック機能（排他制御）」を追加
 * ★UPDATE: アーカイブシート末尾行不足エラーの自動回避 ＆ チェックOFF編集時の即時スキップを追加
 * ==========================================
 */
function processAlertArchiveTrigger(e) {
  if (!e || !e.range) return;
  if (e.range.getColumn() > 1) return;

  // 単一セル編集でチェックが外された(FALSE)場合などはロックを取らず即終了
  if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && e.value !== "TRUE" && e.value !== true) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== "アラートリスト") return;

  const rowStart = e.range.getRow();
  const numRows = e.range.getNumRows();
  const rowEnd = rowStart + numRows - 1;

  if (rowEnd < 2) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("同時編集のロックがかかっています。少し待ってから再度チェックを入れてください。");
    return;
  }

  try {
    const ss = e.source;
    let archiveSheet = ss.getSheetByName("アーカイブ");

    if (!archiveSheet) {
      archiveSheet = ss.insertSheet("アーカイブ");
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
      archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight("bold");
      archiveSheet.setFrozenRows(1);
    }

    const lastCol = sheet.getLastColumn();

    // 行ズレ巻き込み削除を防ぐため、必ず「一番下の行から上に向かって」逆順で処理
    for (let r = rowEnd; r >= Math.max(2, rowStart); r--) {
      const isChecked = sheet.getRange(r, 1).getValue();

      if (isChecked === true || isChecked === "TRUE") {
        const sourceRange = sheet.getRange(r, 1, 1, lastCol);

        const targetRow = archiveSheet.getLastRow() + 1;
        // アーカイブシートの最大行数が足りない場合は事前に行を追加する
        if (archiveSheet.getMaxRows() < targetRow) {
          archiveSheet.insertRowAfter(archiveSheet.getMaxRows());
        }
        const targetRange = archiveSheet.getRange(targetRow, 1, 1, lastCol);

        // 値、リッチテキスト（URLリンク）、背景色などすべてそのままコピー
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

        // 固定行のみになる削除エラーを回避するため、シート行数が2行以下なら空行を足す
        if (sheet.getMaxRows() <= 2) {
          sheet.insertRowAfter(sheet.getMaxRows());
        }

        sheet.deleteRow(r);
      }
    }
  } catch (error) {
    Logger.log("エラーが発生しました: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function onEdit(e) {
  processAlertArchiveTrigger(e);
}