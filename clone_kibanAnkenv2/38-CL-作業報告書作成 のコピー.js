/** 
 * =======================================================================
 *  
 * 【作業報告書 自動作成】CL版
 * ・アクティブ行の情報から作業報告書を1つ作成
 * ・テンプレートをコピーして、名前・本文を自動セット
 * 
 * ======================================================================== 
 **/

/** 
 * =======================================================================
 * 
 * 【作業報告書 自動作成】
 * ・自動実行用
 * 
 * ======================================================================== 
 **/
function CL_createWorkReport() {
  if (startRowCL === null ) return;
  createWorkReportFromRowCL(sheetCL, startRowCL, lastRowCL);
}

/** 
 * =======================================================================
 * 
 * 【作業報告書 自動作成】
 * ・手動実行用
 * 
 * ======================================================================== 
 **/
function CL_createWorkReportManual() {
  const row = searchKanriTableByPopup(sheetCL);
  createWorkReportFromRowCL(sheetCL, row, row);
}

/** 
 * =======================================================================
 *
 *  【作業報告書 自動作成】
 * 
 * ======================================================================== 
 **/
function createWorkReportFromRowCL(sheet, startRow, lastRow) {
  
  for (let row = startRow; row <= lastRow; row++) {

    const valC  = sheet.getRange(row, CLCOL_JUCHU_MITSUMORI).getValue(); 
    const valE  = sheet.getRange(row, CLCOL_CUSTOMER).getValue();        
    const valBD = sheet.getRange(row, CLCOL_担当営業).getValue();        
    const linkCellMitsumori = sheet.getRange(row, CLCOL_MITSUMORI);      
    const linkCellFolder    = sheet.getRange(row, CLCOL_FOLDER);         

    if (!valE) throw new Error(`E列（顧客名）が空です: 行 ${row}`);
    if (!valC) throw new Error(`C列（案件番号）が空です: 行 ${row}`);

    // --- 見積ExcelのファイルID ---
    const fileId = extractFileIdFromUrl(getUrlFromHyperlink(linkCellMitsumori));
    if (!fileId) throw new Error(`見積リンクからファイルID取得不可: 行 ${row}`);

    // --- 作業フォルダのフォルダID ---
    const folderId = extractDriveIdFromUrl(getUrlFromHyperlink(linkCellFolder));
    if (!folderId) throw new Error(`作業フォルダリンクからフォルダID取得不可: 行 ${row}`);

    const baseFolder = DriveApp.getFolderById(folderId);

    // ▼「受注後_作業用資料」フォルダ（必要なら作成＋権限移譲）
    const afterOrderFolder =
      getOrCreateChildFolderWithOwner(baseFolder, '受注後_作業用資料');

    // ▼「08_作業報告書」フォルダ（必要なら作成＋権限移譲）
    const reportFolder =
      getOrCreateChildFolderWithOwner(afterOrderFolder, '08_作業報告書');

    // --- 報告内容の読み取り ---
    const srcSS = openExcelAsSpreadsheet(fileId);
    const found = findSourceSheetForReport(srcSS);

    const reportLines = found
      ? getReportLinesFromSourceSheet(found.sheet, found.startRow)
      : [];

    // --- テンプレコピー ---
    const custName = toZenkakuKana(valE.toString());
    const newName = `【${custName}様】作業報告書※作成中`;

    const template = DriveApp.getFileById(TEMPLATE_FILE_ID);

    // ★ コピー先を「08_作業報告書」フォルダへ
    const newFile = template.makeCopy(newName, reportFolder);

    // ★ コピーしたファイルのオーナーを移譲
    transferFileOwner(newFile);

    const reportSS = SpreadsheetApp.openById(newFile.getId());
    const reportSheet = reportSS.getSheetByName(REPORT_SHEET_NAME) || reportSS.getSheets()[0];

    // --- 各種反映 ---
    const cPrefix = valC.toString().substring(0, 6);
    reportSheet.getRange('Y3').setValue(cPrefix);
    reportSheet.getRange('F5').setValue(custName);

    const f40Text = `ビジネスソリューション事業本部 SOL営業統括部 ${valBD}`;
    reportSheet.getRange('F42').setValue(f40Text);

    fillReportBody(reportSheet, reportLines);
  }
}
