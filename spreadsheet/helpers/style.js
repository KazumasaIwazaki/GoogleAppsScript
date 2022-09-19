var fontName = 'Sawarabi Gothic';

// --------------------------------------------------
// mark: function
// --------------------------------------------------
// 合計行の関数更新
// フォント更新
function setFont_test(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sh.getDataRange();
  setFont(range, fontName);
}
function setFont(pTargetRange, pFontName) {
  console.log('■フォントの更新を行います') ;
  console.log('変更後フォント：'.concat(pFontName));
  pTargetRange.setFontFamily(pFontName);
}

// 垂直位置を中央揃え
function setVirticalMiddle_test(pRange) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sh.getDataRange();

  setVirticalMiddle(range);
}
function setVirticalMiddle(pRange) {
  console.log('■垂直方向を中央揃えにします') ;
  pRange.setVerticalAlignment("middle");
}

// 水平位置を中央揃え
function setHorizontalMiddle_test(pRange) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sh.getRange(sheetRow.title, sheetCol.detailNo, 1, sheetCol.amount - (sheetCol.detailNo - 1));

  setHorizontalMiddle(range);
}
function setHorizontalMiddle(pRange) {
  console.log('■水平方向を中央揃えにします') ;
  pRange.setHorizontalAlignment("center");
}

// 明細の背景色を更新
function updateDetailBackColor_test(pTargetSheet, pDetailLastRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  updateDetailBackColor(ss, 35)
}
function updateDetailBackColor(pTargetSheet, pDetailLastRow) {
  var evenColor = '#f2f2f2';  // 偶数
  var oddColor = '#ffffff';   // 奇数
  console.log('■明細背景色の更新を行います')
  for (var i = sheetRow.detailStart; i <= pDetailLastRow + 1; i++) {
    var result = i % 2;
    var colorCode = '';
    if (result == 1) {
      colorCode = oddColor;
    } else {
      colorCode = evenColor
    }
    pTargetSheet.getRange(i, sheetCol.detailNo, 1, sheetCol.amount - (sheetCol.detailNo - 1)).setBackground(colorCode);
  }
}

// 明細行のセル結合、罫線の設定
function updateDetailRowsCellAndBorder_test(pTargetSheet, pDetailLastRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  updateDetailRowsCellAndBorder(ss, 38)
}
function updateDetailRowsCellAndBorder(pTargetSheet, pDetailLastRow) {
  console.log('■明細行のセル結合、罫線の描画を行います')
  for (var i = sheetRow.detailStart; i <= pDetailLastRow + 1; i++) {
    // セル結合
    var range = pTargetSheet.getRange(i, sheetCol.itemNameStart, 1, sheetCol.itemNameEnd - (sheetCol.itemNameStart - 1)).merge();

    // 罫線描画
    pTargetSheet.getRange(i, sheetCol.itemNameStart, 1, sheetCol.amount - sheetCol.detailNo).setBorder(true, true, true, true, true, true);
  }
}

// 数値フォーマット更新
function setNumberFormat_test(pTargetRange) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // G15以降で'合計'文字を含む行まで検索
  var range = sh.getRange(sheetRow.detailStart, sheetCol.unitPrice, 20, 1);
  setNumberFormat(range);
}
function setNumberFormat(pTargetRange) {
  console.log('■数値フォーマットの更新を行います') ;

  var numberFormat = "#,##0";
  pTargetRange.setNumberFormat(numberFormat);
}

