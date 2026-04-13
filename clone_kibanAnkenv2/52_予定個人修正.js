function e52_updateMemberMenu() {
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