/** 
 * =======================================================================
 * 
 * F列のリンク先フォルダ配下の
 * 「受注後_作業用資料 / 08_作業報告書」フォルダ内から
 * 「作業報告書_署名済み」を含むPDFを検索し、
 * 見つかれば AE列に HYPERLINK をセットする
 *  
 * ======================================================================== 
 **/ 

function SV_setSignedReportLink(row) {

    const folderCell = sheetSV.getRange(row, SVCOL_FOLDER); // F列リンク
    const folderUrl = getUrlFromHyperlink(folderCell);
    if (!folderUrl)  return false;

    const rootFolder = getFolderByUrl(folderUrl);
    if (!rootFolder)  return false;

    // ▼ 受注後_作業用資料 フォルダ検索
    const afterOrder = findSubFolder(rootFolder, "受注後_作業用資料");
    if (!afterOrder)  return false;

    // ▼ 08_作業報告書 フォルダ検索
    const reportFolder = findSubFolder(afterOrder, "08_作業報告書");
    if (!reportFolder)  return false;

    // ▼「作業報告書_署名済み」を含むPDFを探す
    const targetFile = findSignedReport(reportFolder);
    if (!targetFile)  return false;

    const url = reportFolder.getUrl();
    Logger.log(row)

    // ▼ AE列に HYPERLINK を設定
    const rng = sheetSV.getRange(row, 31);  // AE列
    const rich = SpreadsheetApp.newRichTextValue()
    .setText('リンク')
    .setLinkUrl(url)
    .build();

    rng.setRichTextValue(rich);

  return true;
  
}

/** 
 * =======================================================================
 * 
 * 「作業報告書_署名済み」を含む PDF を検索する関数
 *  
 * ======================================================================== 
 **/ 
function findSignedReport(folder) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();

    if (name.includes("作業報告書_署名済み") && name.endsWith(".pdf")) {
      Logger.log("PDFあり")
      return f; // 最初に見つかったものを返す
    }
  }
  return null;
}

/** 
 * =======================================================================
 * 
 * URLからDriveフォルダを取得
 *  
 * ======================================================================== 
 **/ 
function getFolderByUrl(url) {
  try {
    const id = url.match(/[-\w]{25,}/)[0];
    return DriveApp.getFolderById(id);
  } catch (e) {
    return null;
  }
}


/** 
 * =======================================================================
 * 
 * 完全一致でサブフォルダを探す
 *  
 * ======================================================================== 
 **/ 
function findSubFolder(parent, name) {
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName() === name) return f;
  }
  return null;
}
