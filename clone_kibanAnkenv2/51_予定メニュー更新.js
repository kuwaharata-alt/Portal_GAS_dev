/** =======================================================================
 * 
 */
function e51_updateMenu() {
  // スプレッドシートとアクティブシートを取得
  var sheetId = '1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II';
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('予定管理表menu');

  // =========【先月分を非表示に変更】=============================
  // 先月の1日と末日を計算
  var today = new Date();
  var lastMonthFirst = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  var lastMonthLast = new Date(today.getFullYear(), today.getMonth(), 0);

  // A列のデータ範囲を取得（1行目から最終行まで）
  var lastRow = sheet.getLastRow();
  var dateRange = sheet.getRange(1, 1, lastRow); // A列全体
  var dates = dateRange.getValues();


  var startRow = null;
  var endRow = null;

  // 各セルの日付をチェック
  for (var i = 0; i < dates.length; i++) {
    var date = new Date(dates[i][0]);

    // 先月の1日から末日までの日付か確認
    if (date >= lastMonthFirst && date <= lastMonthLast) {
      // 先月の日付に該当する行の範囲を取得（その行全体を対象にする場合、A列だけを対象にする場合に応じて調整）
      var rangeToPaste = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()); // その行全体の範囲を取得
      var values = rangeToPaste.getValues(); // 現在の値を取得

      // 値を貼り付け（関数を削除して値に置き換え）
      rangeToPaste.setValues(values);
      Logger.log(date);

      if (startRow === null) {
        startRow = i + 1; // グループ開始の行
      }
      endRow = i + 1; // グループ終了の行を更新
      } else if (startRow !== null && endRow !== null) {
        // グループ化して折りたたみ（非表示）
        sheet.getRange(startRow, 1, endRow - startRow + 1).activate(); 
        sheet.getActiveRange().shiftRowGroupDepth(1); // 行をグループ化
        sheet.hideRows(startRow, endRow - startRow + 1); // グループを非表示
        startRow = null;
        endRow = null;
      }
    }
    
  // 最後に該当する行をグループ化して非表示にする（もし残っていれば）
  if (startRow !== null && endRow !== null) {
    sheet.getRange(startRow, 1, endRow - startRow + 1).activate(); 
    sheet.getActiveRange().shiftRowGroupDepth(1); // 行をグループ化
    sheet.hideRows(startRow, endRow - startRow + 1); // グループを非表示
  }
  Logger.log("先月分の関数を削除しました");
 
  // =========【5か月後の分を作成】=============================

  var lastRow = sheet.getLastRow();   // A列の最終行を取得
  var lastDate = sheet.getRange(lastRow, 1).getValue();   // 最終行の日付を取得

  // 次の日の日付を取得
  var nextDate = new Date(lastDate);
  nextDate.setDate(nextDate.getDate() + 1);

  // 月の末尾を取得
  var endOfMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0); // 次の月の0日（末日）

  // 日数を取得
  var timeDifference = endOfMonth - nextDate + 1;
  var days = Math.ceil(timeDifference / (1000 * 60 * 60 * 24)); // ミリ秒を日数に変換

  // 数式をオートフィル
  var rangeToCopy = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  var fillRange = sheet.getRange(lastRow, 1, days + 1, sheet.getLastColumn());
  rangeToCopy.autoFill(fillRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  Logger.log("追記しました");

  // ★修正後のシートコピー
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // 現在のスプレッドシートを取得
  var sourceSheet = ss.getSheetByName('予定管理表menu'); // コピー元のシート名を指定
  
  for (var i = 2; i <= 19; i++) {
   
    // 1. シートをコピー
    var newSheet = sourceSheet.copyTo(ss);
    
    // 2. B1セルをiに変更
    newSheet.getRange('B1').setValue(i);
    
    // 3. A1セルの値を取得し、その値をシート名にする
    var sheetName = newSheet.getRange('A1').getValue();
    
    // 既存のシート名があるかチェックし、存在する場合は削除
    var existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet); // 既存のシートを削除
    }
    
    // 新しいシート名を設定
    newSheet.setName(sheetName);
    
    // 4. シートを非表示に変更
    newSheet.hideSheet();

    // 5. シートを保護（シート全体を保護）
    var protection = newSheet.protect();
    protection.setDescription('このシートは保護されています。');
    
    // 6. 保護からオーナー自身の編集権限を除外
    protection.removeEditors(protection.getEditors());
    
    // 7. 必要に応じて、特定のユーザーやグループに編集権限を付与（例: 自分の編集権限を追加）
    var me = Session.getEffectiveUser();
    protection.addEditor(me); // オーナー（実行者）に再度編集権限を付与

    }
}

