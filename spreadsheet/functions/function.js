/**
* 指定した行を対象とし、値が入力された最後の行番号を取得する
* @param {String} pSheetName シート名を文字列（ダブルコーテーションで囲む）で渡します
* @param {String} pPosition 列位置をN1形式で指定します
* @return {Number} 値が入力された最終行の番号
* @customfunction
*/
function GetLastRowByGas(pSheetName, pPosition){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(pSheetName)
    var ColNumber = sheet.getRange(pPosition).getColumn();
    
    //A列の最終行のセルから[Ctrl + ↑]
    var lastRow2 = sheet.getRange(sheet.getMaxRows(), ColNumber).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    return lastRow2;
}
  
function GetLastRowByGasTest(){
  var ret = GetLastRowByGas("納品履歴", "A4")
    Logger.log(ret);
}

/**
* 指定した行を対象とし、値が入力された最後の列番号を取得する
* @param {String} pSheetName シート名を文字列（ダブルコーテーションで囲む）で渡します
* @param {String} pPosition 行位置をN1形式で指定します
* @return {Number} 値が入力された最終列の番号
* @customfunction
*/
function GetLastColByGas(pSheetName, pPosition){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(pSheetName)
    var RowNumber = sheet.getRange(pPosition).getRowIndex();

    //1行目の最終列のセルから[Ctrl + ←]
    var lastCol2 = sheet.getRange(RowNumber, sheet.getMaxColumns()).getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();

    return lastCol2;
}

function GetLastColByGasTest(){
  var ret = GetLastColByGas("納品履歴", "A4")
    Logger.log(ret);
}

