/** 
 * =======================================================================
 * 
 * 【日時タスク】
 * ・更新日が翌日以降の値はクリアする 
 * ・HyperLinkを挿入（見積、作業フォルダ）
 * ・作業工数を追記
 * 
 * ======================================================================== 
 **/ 

function SV_dailyTask() {

  // ----- HyperLink挿入 --------------------------------------------------
  SV_InsertLinks()
  
  // ----- 作業フォルダ作成 -------------------------------------------------
  //SV_AcreateWorkFolders()

  // ----- 作業工数追記 ----------------------------------------------------
  SV_TotalWork()

  // ----- 作業報告書作成 ----------------------------------------------------
  //SV_createWorkReport()

  // ----- 更新日修正（明日以降の値はリセットする） ------------------------------
  tommorow.setHours(0, 0, 0, 0); // 時刻をリセットして日付のみ比較
  
  const rangeA = sheetSV.getRange(1, 1, lastRowSV, 1); // A列のすべてのデータを取得
  const values = rangeA.getValues(); // 取得したデータを配列として取得
  
  // 値が今日以降の日付であればそのセルの値をクリア
  for (let i = 0; i < values.length; i++) {
    if (SVCOL_UPDATE_DATE > tommorow) {
      sheetSV.getRange(i + 1, 1).clearContent(); // 行番号は1から始まるため、インデックスに1を足す
    }
  }
}

/** 
 * =======================================================================
 * 
 * 【クローズ時トリガー】
 * ・98.SE作業完了に変更時に作動する
 * ・ドライブに作業報告書が正しい名称で保存されているか確認する
 * ・フィードバックメールを通知する
 * 
 * ======================================================================== 
 **/ 

function SV_CloseTask(e){
  // ****** シンプルトリガー防止 *****
  if (e && e.authMode && e.authMode === ScriptApp.AuthMode.LIMITED) return;
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getSheetName() !== vSV) return;

  const row = range.getRow();
  const col = range.getColumn();

  if (row < 2) return;
  if (col !== 2) return; // ステータス欄

  const oldValue = (typeof e.oldValue !== 'undefined') ? e.oldValue : '';
  const newValue = String(e.value || range.getValue()).trim();

  if (newValue !== '98. SE対応完了') return;

  SpreadsheetApp.getActive().toast(
    'クローズ後の処理を開始します', // メッセージ
    '【98.SE対応完了】へステータスが変更されました',     // タイトル（任意）
    10                                            // 表示秒数
  );

  // ----- 作業報告書アップロードチェック -------------------------------------
  const ui = SpreadsheetApp.getUi();
  const pdfExists = SV_setSignedReportLink(row);
  if (!pdfExists) {
    // ステータスを元に戻す
    range.setValue(oldValue);

    ui.alert(
      '作業報告書 未アップロード',
      '「受注後_作業用資料 ＞ 08_作業報告書」フォルダに、\n' +
      '「作業報告書_署名済み」のPDFをアップロードしてください。',
      ui.ButtonSet.OK
    );
    return;
  }
  // ----- フィードバック通知発報 -----------------------------------------------
  CloseMail_SV(row)
}