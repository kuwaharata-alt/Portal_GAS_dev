/** 
 * =======================================================================
 * 
 * 【メニュー追加】
 * ・スプレッドシートにGAS実行ボタンを追加
 * 
 * ========================================================================
 **/ 

function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('SV手動実行')
    .addItem('ハイパーリンク挿入', 'SV_InsertLinksManual')
    .addItem('アサイン通知', 'SV_AssignAnnounce_manual')
    .addItem('フォルダ作成', 'SV_AcreateWorkFoldersManual')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('CL手動実行')
    .addItem('ハイパーリンク挿入', 'CL_InsertLinksManual')
    .addItem('フォルダ作成', 'CL_AcreateWorkFoldersManual')
    .addToUi();
}