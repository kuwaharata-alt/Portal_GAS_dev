/**
 * 【案件情報登録】1時～2時 実行
 * ・作業依頼をメールから取得
 * ・「IT基盤構築案件管理表-計上管理」から情報を転写
 * ・「案件情報」シートを更新
 * 
 * ・作業フォルダ、見積リンクを挿入
 * ・フォルダ作成
 * ・見積工数を転記
 */

function Auto_案件登録() {
  //* 案件情報登録  */
  writeLog_('Auto_作業依頼登録', '▼開始▼');
  Auto_作業依頼登録()
  SpreadsheetApp.flush();
  writeLog_('Auto_作業依頼登録', '▲終了▲');

  writeLog_('Auto_計上管理転写', '▼開始▼');
  Auto_計上管理転写();
  SpreadsheetApp.flush();
  writeLog_('Auto_計上管理転写', '▲終了▲');

  writeLog_('Auto_案件情報作成', '▼開始▼');
  Auto_案件情報作成();
  SpreadsheetApp.flush();
  writeLog_('Auto_案件情報作成', '▲終了▲');

  //* 案件設定 */
  writeLog_('Auto_リンク挿入', '▼開始▼');
  Auto_リンク挿入();
  SpreadsheetApp.flush();
  writeLog_('Auto_リンク挿入', '▲終了▲');

  writeLog_('Auto_見積工数取得', '▼開始▼');
  Auto_見積工数取得()
  SpreadsheetApp.flush();
  writeLog_('Auto_見積工数取得', '▲終了▲');

  //* 案件管理表更新 */
  writeLog_('Auto_案件情報To案件管理表_転記', '▼開始▼');
  Auto_案件情報To案件管理表_転記()
  writeLog_('Auto_案件情報To案件管理表_転記', '▲終了▲');

}

/**
 * 【案件情報登録】2時～3時 実行
 * ・フォルダ作成
 * ・見積工数を転記
 */
function Auto_案件情報更新(){
  writeLog_('Auto_受注後フォルダ作成', '▼開始▼');
  Auto_受注後フォルダ作成();
  SpreadsheetApp.flush();
  writeLog_('受注Auto_受注後フォルダ作成後サブフォルダ作成', '▲終了▲');

}
