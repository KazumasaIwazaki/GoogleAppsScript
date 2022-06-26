// mark: menu

function onOpen() {
  var MenuEntries = [];
  MenuEntries.push({name: "メニュー表示：サンプルメッセージ表示", functionName: "showMessage"});
  MenuEntries.push({name: "メニュー表示：A1セルの内容をメッセージ表示", functionName: "showA1CellData"});

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("カスタムメニュー", MenuEntries);

}

function showMessage() {
  Browser.msgBox('サンプルメッセージ', 'サンプルメッセージです。\\nプログラムで自由に表示内容を変えることができます。', Browser.Buttons.OK);
}

function showA1CellData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var data = sheet.getRange('A1').getValue();
  Browser.msgBox('A1セルのメッセージ', data, Browser.Buttons.OK);
}