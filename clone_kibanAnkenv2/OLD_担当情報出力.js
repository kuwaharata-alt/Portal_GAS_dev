function a03_displayMember(e) {
  // ===== 特定のセルが編集された際に、情報を表示する =====
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  // 「案件管理表」以外のシートなら処理しない
  if (sheetName !== "案件管理表") {
    return;
  }

  if (col === 10 && row >= 2 && row <= 1000) {
    // J列のセルが編集された場合、担当者情報を表示する
    fc_showMember(sheet, row);
  }else if (col === 5 && row >= 2 && row <= 1000) {
    // E列のセルが編集された場合、案件情報を表示する
    fc_showCase(sheet, row);
  }else{
    return;
  }

}

function fc_showMember(sheet, row) {
  // ===== 担当者情報を出力する ===== 
  var startCol = 11
  var endCol = 18 //メンバーの数 (列数ではない)
  var headers = sheet.getRange(1, startCol, 1, endCol).getValues()[0]; // L1～AC1の見出し
  var rowValues = sheet.getRange(row, startCol, 1, endCol).getValues()[0]; // L～ACの対象行の値

  if (sheet.getRange(row, 7).getValue() === "SV系"){
    var role = [
      "★監督",
      "★管理者",
      "☆作業者",
    ]
  }
  if (sheet.getRange(row, 7).getValue() === "PC系"){
    var role = [
      "★総合管理者",
      "★倉庫-管理",
      "★現地-管理",
      "☆倉庫＋現地作業",
      "〇倉庫-マスター",
      "☆倉庫-作業",
      "☆現地-作業",
    ]
  }
  var members = [];
  
  // 各役職に対応する空の配列を作成
  for (var i = 0; i < role.length; i++) {
    members[role[i]] = [];
  }

  // データを振り分ける
  for (var i = 0; i < rowValues.length; i++) {
    if (role.includes(rowValues[i])) {
      members[rowValues[i]].push(headers[i]); // 該当する役職のキーに見出しを追加
    }
  }

  // 出力するHTMLフォーマットを作成
  var html = "<p><b>【" + sheet.getRange(row, 7).getValue() + "】" + "<br>" + "【" + sheet.getRange(row, 3).getValue() + "】" + sheet.getRange(row, 4).getValue() + "</b></p>";
  for (var j = 0; j < role.length; j++) {
    html += "<p>" + role[j] + "<br>" + "　" + (members[role[j]].length ? members[role[j]].join(", ") : "ー") + "</p>";
  }

  // Apps Script アプリケーションで出力する
  var ui = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(200);
  SpreadsheetApp.getUi().showSidebar(ui);
}


function fc_showCase(sheet, row) {
  // ===== 案件情報を出力する ===== 
  var startCol = 11
  var endCol = 18 //メンバーの数 (列数ではない)
  var headers = sheet.getRange(1, startCol, 1, endCol).getValues()[0]; // L1～AC1の見出し
  var rowValues = sheet.getRange(row, startCol, 1, endCol).getValues()[0]; // L～ACの対象行の値


  var sheetId = '1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II';
  var sheetIT = SpreadsheetApp.openById(sheetId).getSheetByName('IT基盤構築案件管理表(抜粋-計上管理)');

  if (sheet.getRange(row, 7).getValue() === "SV系"){
    var sheetA = SpreadsheetApp.openById(sheetId).getSheetByName('SV案件管理表');
    var outputMenu = [
      "■ステータス",
      "▽全体：",
      "▽事前作業：",
      "▽現地作業：",
      "▽事後作業：",
      "■Info",
      "▽計上予定月：",
      "▽担当プリ：",
      "▽事前作業場所",
      "▽現地作業場所",
    ]
    var outputValue = [
      "",
      sheet.getRange(row, 2).getValue(),
      sheetA.getRange(row, 15).getValue(),
      sheetA.getRange(row, 23).getValue(),
      "▽事後作業：",
      "■Info",
      "▽計上予定月：",
      "▽担当プリ：",
      "▽事前作業場所",
      "▽現地作業場所",
    ]
  }
  if (sheet.getRange(row, 7).getValue() === "PC系"){
    var sheetA = SpreadsheetApp.openById(sheetId).getSheetByName('CL案件管理表');
    var outputMenu = [
      "■ステータス",
      "▽全体",
      "▽事前作業",
      "▽現地作業",
      "▽事後作業",
      "■Info",
      "▽計上予定月：",
      "▽担当プリ：",
      "▽事前作業場所",
      "▽現地作業場所",
    ]
  }

  // 出力するHTMLフォーマットを作成
  var html = "<p><b>【" + sheet.getRange(row, 7).getValue() + "】" + "<br>" + "【" + sheet.getRange(row, 3).getValue() + "】" + sheet.getRange(row, 4).getValue() + "</b></p>"; //案件情報
    for (var j = 0; j < outputMenu.length; j++) {
      html += "<p>" + outputMenu[j] + "<br>" + "　" + outputValue[j] + "</p>";
    }

  // Apps Script アプリケーションで出力する
  var ui = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(200);
  SpreadsheetApp.getUi().showSidebar(ui);
}