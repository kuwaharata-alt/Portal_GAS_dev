/** 
 * =======================================================================
 * 
 * 【作業用フォルダの作成】
 * D列のリンク先フォルダ配下に
 * 「受注後_作業用資料」フォルダを探し、
 * その配下に 01,02,08,09 の4フォルダを作成
 *
 * 新規に作成したフォルダは
 * 「it-kiban@systena.co.jp」にオーナー権限を譲渡する
 * 
 * ========================================================================
 **/
// 作成するフォルダ名リスト
const SUB_CLFOLDERS = [
  '01_作業用資料',
  '02_チェックシート',
  '08_作業報告書',
  '09_納品物'
];

/** 
 * =======================================================================
 * 
 * 【フォルダ作成】
 * ・自動実行用
 * 
 * ========================================================================
 **/
function CL_AcreateWorkFolders(){
  if (startRowCL === null ) return;
  const startRow = startRowCL
  const lastRow = lastRowCL
  AcreateWorkFoldersCL(startRow,lastRow)
}

/** 
 * =======================================================================
 * 
 * 【フォルダ作成】
 * ・手動実行用
 * ・案件番号を入力する
 * 
 * ========================================================================
 **/
function CL_AcreateWorkFoldersManual(){
  const row = searchKanriTableByPopup(sheetCL);
  const startRow = row
  const lastRow = row
  AcreateWorkFoldersCL(startRow,lastRow)
}

/** 
 * =======================================================================
 * 
 * 【フォルダ作成】
 * ・フォルダ作成する関数
 * ・引数：開始行、終了行
 * 
 * ========================================================================
 **/
function AcreateWorkFoldersCL(startRow,lastRow) {
  Logger.log("〇フォルダ作成(CL):");

  for (let r = startRow; r <= lastRow; r++) {

    const linkCell = sheetCL.getRange(r, CLCOL_FOLDER); // D列：HYPERLINK

    const url = getUrlFromHyperlink(linkCell);
    if (!url) continue;

    const folderId = extractDriveIdFromUrl(url);
    if (!folderId) continue;

    try {
      const baseFolder = DriveApp.getFolderById(folderId);

      // 「受注後_作業用資料」フォルダを取得 or 作成(＋新規ならオーナー譲渡)
      const workFolder = getOrCreateFolderWithOwnership(baseFolder, '受注後_作業用資料');

      // 下位フォルダ 4つ作成（新規ならオーナー譲渡）
      SUB_CLFOLDERS.forEach(name => {
        getOrCreateFolderWithOwnership(workFolder, name);
      });

      Logger.log(`Row ${r}: フォルダ作成＆オーナー処理完了`);

    } catch (e) {
      Logger.log(`Row ${r}: エラー - ${e}`);
    }
  }
}


