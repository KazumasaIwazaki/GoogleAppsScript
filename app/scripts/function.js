// スプレッドシートを複写、ダウンロード用の編集を行う
function CopySpreadSheetForDownload(pSpreadSheet) {
  // シートを複写
  var DuplicatedSheet = pSpreadSheet.getActiveSheet().copyTo(pSpreadSheet);

  // 全データを取得して値として貼り付け
  var data = DuplicatedSheet.getRange(1, 1, DuplicatedSheet.getLastRow(), DuplicatedSheet.getLastColumn()).getDisplayValues();
  DuplicatedSheet.getRange(1, 1, DuplicatedSheet.getLastRow(), DuplicatedSheet.getLastColumn()).setValues(data);
  
  SpreadsheetApp.flush()
  return DuplicatedSheet;
}

// 対象シートごとの独自編集を行う
function EditBySheet(pTargetName, pTargetSheet) {
  // 「インボイス」シートの場合はA列を削除
  if (pTargetName == "インボイス") {
    pTargetSheet.deleteColumn(1);
    SpreadsheetApp.flush();
  }
 
}