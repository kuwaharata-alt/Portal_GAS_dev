function mikomikonurse() {
  // 参照するスプレッドシートのIDとシート名
  var sourceSpreadsheetId = '1mDMUFShbF9Zmq6VC03_yNHpBUExNiVoqJ2nLZyt6gM8';
  var sourceSheetName = '見込管理';

  // 参照するスプレッドシートとシートを取得
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

  // 転記先のシートを指定して取得
  var destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var destinationSheetName = '見込高案件';
  var destinationSheet = destinationSpreadsheet.getSheetByName(destinationSheetName);

  // 参照するシートのデータを取得
  var sourceData = sourceSheet.getDataRange().getValues();

  // 転記先のシートの開始行を設定
  var destinationRowStart = 6;
  var destinationRow = destinationRowStart;

  // **転記先の既存データを消去**
  var lastRow = destinationSheet.getLastRow();
  if (lastRow >= destinationRowStart) {
    destinationSheet.getRange(destinationRowStart, 1, lastRow - destinationRowStart + 1, destinationSheet.getLastColumn()).clearContent();
  }

  // 転記するデータを格納する配列
  var transferData = [];

  // データを1行ずつ処理
  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var progress = row[8];   // I列（進行）
    var tantouPS = row[19]; // T列（担当PS）
    var サービス売上 = row[23]; // X列（サービス売上）

    // 条件に一致するか確認
    if (
      (progress === 'C:内示' || progress === 'D:見込高' ) &&
      (tantouPS === '田崎玲央' || tantouPS === '松井俊樹' || tantouPS === '宮崎茂雄' || tantouPS === '網谷祐亮' || tantouPS === '岩澤菜摘' || tantouPS === '浜里航也' || tantouPS === '青栁凜汰郎') &&
      サービス売上
    ) {
      // 条件に一致する場合、選択した列のデータを配列に追加
      transferData.push([
        row[1],  // B列
        row[2],  // C列
        row[3],  // D列
        row[5],  // F列
        row[6],  // G列
        row[7],  // H列
        row[8],  // I列
        row[9],  // J列
        row[17], // R列
        row[18], // S列
        row[19], // T列
        row[20], // U列
        row[21], // V列
        row[22], // W列
        row[23]  // X列
      ]);
    }
  }

  // 転記するデータがある場合のみ処理
  if (transferData.length > 0) {
    // 転記先のシートにデータを書き込む
    destinationSheet.getRange(destinationRowStart, 1, transferData.length, transferData[0].length).setValues(transferData);

    // 並び替えの範囲を指定
    var sortRange = destinationSheet.getRange(destinationRowStart, 1, transferData.length, transferData[0].length);

    // 並び替えの条件を設定
    sortRange.sort([
      { column: 6, ascending: true },  // F列を昇順
      { column: 7, ascending: true },  // G列を昇順
      { column: 1, ascending: true },  // A列を昇順
      { column: 8, ascending: true }   // H列を昇順
    ]);
  }
}