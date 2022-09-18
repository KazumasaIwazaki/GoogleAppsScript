// Googleドライブへファイル出力
function OutputTsvData(pOutputFolderId, pOutputFileName, pData){
    
    var folder = DriveApp.getFolderById(pOutputFolderId);
    var contentType = 'text/plain';
    var charset = 'utf-8';
  
    var blob = Utilities.newBlob('', contentType, pOutputFileName).setDataFromString(pData, charset);
  
    // ファイルに保存
    var FileId = folder.createFile(blob).getId();
    
    return FileId;
  }
  
// PDF作成関数　引数は（folderid:保存先フォルダID, ssid:PDF化するスプレッドシートID, sheetid:PDF化するシートID, filename:PDFの名前）
function CreatePDF(folderid, ssid, sheetid, filename){

    // PDFファイルの保存先となるフォルダをフォルダIDで指定
    var folder = DriveApp.getFolderById(folderid);
  
    // スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
    var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
  
    // PDF作成のオプションを指定
    var opts = {
      exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
      format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
      size:         "A4",     // 用紙サイズの指定 legal / letter / A4
      portrait:     "false",   // true → 縦向き、false → 横向き
      fitw:         "true",   // 幅を用紙に合わせるか
      sheetnames:   "false",  // シート名をPDF上部に表示するか
      printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
      pagenumbers:  "false",  // ページ番号の有無
      gridlines:    "false",  // グリッドラインの表示有無
      fzr:          "false",  // 固定行の表示有無™
      gid:          sheetid   // シートIDを指定 sheetidは引数で取得
    };
    
    var url_ext = [];
    
    // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
    for( optName in opts ){
      url_ext.push( optName + "=" + opts[optName] );
    }
  
    // url_extの各要素を「&」で繋げる
    var options = url_ext.join("&");
  
    // optionsは以下のように作成しても同じです。
    // var ptions = 'exportFormat=pdf&format=pdf'
    // + '&size=A4'                       
    // + '&portrait=true'                    
    // + '&sheetnames=false&printtitle=false' 
    // + '&pagenumbers=false&gridlines=false' 
    // + '&fzr=false'                         
    // + '&gid=' + sheetid;
  
    // API使用のためのOAuth認証
    var token = ScriptApp.getOAuthToken();
  
      // PDF作成
      var response = UrlFetchApp.fetch(url + options, {
        headers: {
          'Authorization': 'Bearer ' +  token
        }
      });
  
      // 
      var blob = response.getBlob().setName(filename + '.pdf');
  
    //}
  
    //　PDFを指定したフォルダに保存
    var FileId = folder.createFile(blob).getId();

    return FileId;
  
  }
  
  // PDF作成関数　引数は（folderid:保存先フォルダID, ssid:PDF化するスプレッドシートID, sheetid:PDF化するシートID, filename:PDFの名前）
function CreateXlsx(folderid, ssid, sheetid, filename){

    // PDFファイルの保存先となるフォルダをフォルダIDで指定
    var folder = DriveApp.getFolderById(folderid);
  
    // スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
    var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
  
    // PDF作成のオプションを指定
    var opts = {
      exportFormat: "xlsx",    // ファイル形式の指定 pdf / csv / xls / xlsx
      format:       "xlsx",    // ファイル形式の指定 pdf / csv / xls / xlsx
      size:         "A4",     // 用紙サイズの指定 legal / letter / A4
      portrait:     "false",   // true → 縦向き、false → 横向き
      fitw:         "true",   // 幅を用紙に合わせるか
      sheetnames:   "false",  // シート名をPDF上部に表示するか
      printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
      pagenumbers:  "false",  // ページ番号の有無
      gridlines:    "false",  // グリッドラインの表示有無
      fzr:          "false",  // 固定行の表示有無™
      gid:          sheetid   // シートIDを指定 sheetidは引数で取得
    };
    
    var url_ext = [];
    
    // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
    for( optName in opts ){
      url_ext.push( optName + "=" + opts[optName] );
    }
  
    // url_extの各要素を「&」で繋げる
    var options = url_ext.join("&");
  
    // optionsは以下のように作成しても同じです。
    // var ptions = 'exportFormat=pdf&format=pdf'
    // + '&size=A4'                       
    // + '&portrait=true'                    
    // + '&sheetnames=false&printtitle=false' 
    // + '&pagenumbers=false&gridlines=false' 
    // + '&fzr=false'                         
    // + '&gid=' + sheetid;
  
    // API使用のためのOAuth認証
    var token = ScriptApp.getOAuthToken();
  
      // PDF作成
      var response = UrlFetchApp.fetch(url + options, {
        headers: {
          'Authorization': 'Bearer ' +  token
        }
      });
  
      // 
      var blob = response.getBlob().setName(filename + '.xlsx');
  
    //}
  
    //　PDFを指定したフォルダに保存
    var FileId = folder.createFile(blob).getId();

    return FileId;
  }

  //SpreadsheetをExcelファイルに変換してドライブに保存、Fileを返す
function ss2xlsx(folderid, spreadsheet_id) {
    // PDFファイルの保存先となるフォルダをフォルダIDで指定
    var folder = DriveApp.getFolderById(folderid);
  
    var new_file;
    var url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?format=xlsx";
    var options = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() == 200) {
      var ss = SpreadsheetApp.openById(spreadsheet_id);
      new_file = folder.createFile(res.getBlob()).setName(ss.getName() + ".xlsx");
    }
    return new_file;
  }
  
  
function ShowDownloadDialog(FileId) {
   // dialog.html をもとにHTMLファイルを生成
  var html = HtmlService.createTemplateFromFile("dialog");
  html.LinkUrl = FileId;
  
  // 上記HTMLファイルをダイアログ出力
  var ui = html.evaluate()
  SpreadsheetApp.getUi().showModalDialog(ui, "ファイルダウンロード");  
}