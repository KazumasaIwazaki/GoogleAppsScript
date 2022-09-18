// --------------------------------------------------
// mark: declare
// --------------------------------------------------
var OUTPUT_FOLDER_ID = "";

// --------------------------------------------------
// mark: main
// --------------------------------------------------
// 開いているシートの全入力内容を取得してタブ区切りでファイル出力する
function OutputTsvDataByGas() {
    var target = GetActiveSheetDataForDeliveryPlan();
    var OutputData = FormatDataForTsv(target)
    var FileName = GetDate() + "納品プラン作成" + ".txt"
    var FolderId = GetOutputFolderId(OUTPUT_FOLDER_ID);
    var FileId = OutputTsvData(FolderId, FileName, OutputData);
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

  // スプレッドシートを複写、ダウンロード用の編集を行う
  var DuplicatedSheet = CopySpreadSheetForDownload(ss);

  // シートごとの独自編集を行う
  EditBySheet(customer_name, DuplicatedSheet)

  // PDF作成関数
  var FolderId = GetOutputFolderId(OUTPUT_FOLDER_ID);
  var FileId = CreatePDF(FolderId, ssid, DuplicatedSheet.getSheetId(), timestamp + customer_name);

  // 複写したシートを削除
  ss.deleteSheet(DuplicatedSheet);

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
  
  // スプレッドシートを複写、ダウンロード用の編集を行う
  var DuplicatedSheet = CopySpreadSheetForDownload(ss);

  // シートごとの独自編集を行う
  EditBySheet(customer_name, DuplicatedSheet)

  // エクセル作成関数
  var FolderId = GetOutputFolderId(OUTPUT_FOLDER_ID);
  var FileId = CreateXlsx(FolderId, ssid, DuplicatedSheet.getSheetId() , timestamp + customer_name);

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

// --------------------------------------------------
//  mark: function
// --------------------------------------------------

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


function GetTimeStamp() {
    var date = new Date();
    var TimeStamp = Utilities.formatDate( date, 'Asia/Tokyo', 'yyyyMMdd_hhmmss');
    
    return TimeStamp;
}

function GetDate() {
    var date = new Date();
    var TimeStamp = Utilities.formatDate( date, 'Asia/Tokyo', 'yyyyMMdd');
    
    return TimeStamp;
}

function GetOutputFolderId(pOutputFOlderId) {
  if (pOutputFOlderId != "") {return pOutputFOlderId}
  
  return GetParentFolderId();

}

function GetParentFolderId() {

  // 自身のスプレッドシートのIDを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  
  // 親フォルダ(ファイル自身が格納されているフォルダ)を取得
  var parentFolder = DriveApp.getFileById(ssId).getParents();
  var folder = parentFolder.next();

  // フォルダ名を表示
  return folder.getId()
}


