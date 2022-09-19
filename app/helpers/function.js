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

// 取得データをtsv形式へ整形
function FormatDataForTsv(pTarget) {

  var FormattedString = "";

  for (let row = 0; row <= pTarget.length - 1; row++) {
    var str = "";
    var row_length = pTarget[row].length - 1
    for (let col = 0; col <= row_length; col++) {
        if (col > 0) { str = str.concat('\t') }
        str = str.concat(pTarget[row][col]);
    }

    FormattedString = FormattedString.concat(str, '\n');
  }

  return FormattedString
}

