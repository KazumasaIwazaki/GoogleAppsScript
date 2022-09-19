// --------------------------------------------------
// mark: main
// --------------------------------------------------
// 開いているシートの全入力内容を取得してタブ区切りでファイル出力する
function OutputCsvDataByGas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = GetActiveSheetBaseColA();
  // var OutputData = FormatDataForTsv(target)
  var OutputData = target.join('\n')
  var FileName = GetDate() + ss.getActiveSheet().getSheetName() + ".csv"
  var FolderId = GetOutputFolderId(OUTPUT_FOLDER_ID);
  var FileId = OutputCsvData(FolderId, FileName, OutputData);
  ShowDownloadDialog(FileId);
}

// 開いているシートをPDFとして出力する
function CreatePdfByGas() {
  // 現在開いているスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 現在開いているスプレッドシートのIDを取得
  var ssid = ss.getId();
  // 現在開いているスプレッドシートのシートIDを取得
  var sheetid = ss.getActiveSheet().getSheetId();

  // ファイル名に使用する名前を取得
  var customer_name = ss.getActiveSheet().getSheetName();
  // ファイル名に使用するタイムスタンプを取得
  var timestamp = GetDate();

  // PDF作成関数
  var FolderId = GetOutputFolderId(OUTPUT_FOLDER_ID);
  var FileId = CreatePDF(FolderId, ssid, sheetid, timestamp + customer_name);

    // ダウンロード用のダイアログを表示
  ShowDownloadDialog(FileId);
}
 
function CreateExcelSheetByGas() {
  // 現在開いているスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 現在開いているスプレッドシートのIDを取得
  var ssid = ss.getId();
  // 現在開いているスプレッドシートのシートIDを取得
  var sheetid = ss.getActiveSheet().getSheetId();

  // ファイル名に使用する名前を取得
  var customer_name = ss.getActiveSheet().getSheetName();
  // ファイル名に使用するタイムスタンプを取得
  var timestamp = GetDate();
  
  // エクセル作成関数
  var FolderId = GetOutputFolderId(OUTPUT_FOLDER_ID);
  var FileId = CreateXlsx(FolderId, ssid, sheetid , timestamp + customer_name);

  // 複写したシートを削除
  ss.deleteSheet(DuplicatedSheet);

  // ダウンロード用のダイアログを表示
  ShowDownloadDialog(FileId);

}

function CreateExcelByGas() {
  // 現在開いているスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 現在開いているスプレッドシートのIDを取得
  var ssid = ss.getId();

  var FolderId = GetOutputFolderId(OUTPUT_FOLDER_ID);
  var excel_file = ss2xlsx(FolderId, ssid);
  if (excel_file !== undefined) {
    Logger.log("Name:" + excel_file.getName());
  }
}

