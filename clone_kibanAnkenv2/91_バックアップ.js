function x91_backupSpreadsheet() {
  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // バックアップのコピーを作成
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  var backupName = spreadsheet.getName() + "_Backup_" + now;
  
  // バックアップをGoogleドライブに作成
  var file = DriveApp.getFileById(spreadsheet.getId());
  var folder = DriveApp.getFolderById('1nShNKpwcBDQkHd74r3gfu7FF6P-bFzZ9'); // フォルダIDを設定
  file.makeCopy(backupName, folder);

}
