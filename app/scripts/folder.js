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