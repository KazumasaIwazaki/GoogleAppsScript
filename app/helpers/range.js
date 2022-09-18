// 選択中スプレッドシートの全データを取得
function GetActiveSheetAllRange() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // getRange(開始行番号,開始列番号,行数,列数)
    var data = ss.getRange(1, 1, ss.getLastRow(), ss.getLastColumn()).getValues();

    return data;

  }
function GetActiveSheetAllRange_test() {
  var data = GetActiveSheetAllRange();
  Logger.log(data);
}

// 選択中スプレッドシートの全データを取得（A1列基準）
function GetActiveSheetDataForDeliveryPlan() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // A1セルを基準に最終行を取得
    var LastRow = GetLastRow(1);

    // getRange(開始行番号,開始列番号,行数,列数)
    var data = ss.getRange(2, 1, LastRow, ss.getLastColumn()).getValues();

    return data;

  }

  function GetLastRow(pColPosition){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    //A列の最終行のセルから[Ctrl + ↑]
    var lastRow2 = sheet.getRange(sheet.getMaxRows(), pColPosition).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    return lastRow2;
}

// 値ありのセルを基準として最終セルを算出する
function GetLastRowOnValue(pSheetName, pColPosition, pHeaderRowPosition){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(pSheetName);

    // arrayformula関数でデータ取得していることを考慮、「値あり」を行数としてカウント
    var targetRange = '{0}{1}:{0}'.replace(/\{0\}/g, pColPosition).replace('{1}', pHeaderRowPosition);
    var lastRow2 = sheet.getRange(targetRange).getValues().filter(String).length + pHeaderRowPosition - 1;

    return lastRow2;
}
function GetLastRowOnValue_test(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'マスタCSV';

  var result = GetLastRowOnValue(sheetName, 'B', 5);
  Logger.log(result);
}

function GetLastCol(pRowPosition){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getRange(pPosition);

  //1行目の最終列のセルから[Ctrl + ←]
 var lastCol2 = sheet.getRange(pRowPosition, sheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();

  return lastCol2;
}

// 値ありのセルを基準として最終セルを算出する
function GetLastColOnValue(pSheetName, pRowPosition){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(pSheetName);

    //1行目の最終列のセルから[Ctrl + ←]
    // var lastCol2 = sheet.getRange(pRowPosition, sheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
    // var targetRange = '{0}{1}:{0}'.replace(/\{0\}/g, pColPosition).replace('{1}', pHeaderRowPosition);
    var lastCol2 = sheet.getRange('6:6').getValues().filter(String).length;

    return lastCol2;
}
function GetLastColOnValue_test(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'マスタCSV';

  var result = GetLastColOnValue(sheetName, 5);
  Logger.log(result);
}
