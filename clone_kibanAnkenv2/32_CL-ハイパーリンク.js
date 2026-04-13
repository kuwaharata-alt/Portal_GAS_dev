/** 
 * =======================================================================
 * 
 * 【リンク挿入】
 * ・D列に見積りのリンクを入れる
 * ・F列に作業フォルダへのリンクを入れる
 * 
 * ======================================================================== 
 **/ 
 
function CL_InsertLinks() {    
  Logger.log("〇リンク追加");
  if (startRowCL === null ) return;
 
  let startRow = null;
  let endRow = null;
  
  const lastRowCL = getLastRowByColumn(sheetCL, CLCOL_MITSUMORI);

  // ----- D列に見積のリンクを入れる -----
  for (let i = startRowCL; i <= lastRowCL; i++) {
    const cellDate = sheetCL.getRange(i, 4).getValue();

    // 日付が昨日の日付と一致するか確認
    if (cellDate instanceof Date && 
        cellDate.getFullYear() === yesterday.getFullYear() &&
        cellDate.getMonth() === yesterday.getMonth() &&
        cellDate.getDate() === yesterday.getDate()) {
      
      if (startRow === null) {
        startRow = i; // 最初の行
      }
      endRow = i; // 最後の行を更新
    }
  }
  
  // ----- リンクを入れる -----
  if (startRow !== null && endRow !== null) {

    for (row = startRow; row <= endRow; row++) {
    //HyperLink(作業フォルダ)
      const searchTermFolder = sheetCL.getRange(row, CLCOL_MITSUMORI-1).getValue().toString().substring(0, 6);
      const folders = DriveApp.searchFolders(`title contains '${searchTermFolder}'`);
      const folder = folders.next();
      const folderUrl = folder.getUrl();
      
    // F列にハイパーリンクを挿入
      sheetCL.getRange(row, CLCOL_FOLDER).setFormula(`=HYPERLINK("${folderUrl}", "リンク")`);  // リンクを挿入
    
    //HyperLink(見積り)
     const searchTermFile = sheetCL.getRange(row, CLCOL_MITSUMORI-1).getValue().toString();
     const files = DriveApp.searchFiles(`title contains '${searchTermFile}'`);
     const file = files.next();
     const fileUrl = file.getUrl();

    //　D列にハイパーリンクを挿入
      sheetCL.getRange(row, CLCOL_MITSUMORI).setFormula(`=HYPERLINK("${fileUrl}", "見積り")`);  // リンクを挿入
    }
  } else {
    console.log("昨日の日付が入力されている行は見つかりませんでした。");
  }
}

/** 
 * =======================================================================
 * 
 * 【リンク挿入】
 * ・手動実行用
 * ・D列に見積りのリンクを入れる
 * ・F列に作業フォルダへのリンクを入れる
 * 
 * ======================================================================== 
 **/ 

function CL_InsertLinksManual() {   
  Logger.log("〇リンク追加");
  const rows = runWithPopup();
  if (!rows) {
    SpreadsheetApp.getUi().alert("処理をキャンセルしました。");
    return;
  }

  const { startRow, endRow } = rows;  // ← 分割代入で取得

  // ----- リンクを入れる -----
  if (startRow !== null && endRow !== null) {
    console.log("最初の行: " + startRow + ", 最後の行: " + endRow);

    for (row = startRow; row <= endRow; row++) {
    //HyperLink(作業フォルダ)
      const searchTermFolder = sheetCL.getRange(row, CLCOL_MITSUMORI-1).getValue().toString().substring(0, 6);
      const folders = DriveApp.searchFolders(`title contains '${searchTermFolder}'`);
      const folder = folders.next();
      const folderUrl = folder.getUrl();
      
    // F列にハイパーリンクを挿入
      sheetCL.getRange(row, CLCOL_FOLDER).setFormula(`=HYPERLINK("${folderUrl}", "リンク")`);  // リンクを挿入
    
    //HyperLink(見積り)
     const searchTermFile = sheetCL.getRange(row, CLCOL_MITSUMORI-1).getValue().toString();
     const files = DriveApp.searchFiles(`title contains '${searchTermFile}'`);
     const file = files.next();
     const fileUrl = file.getUrl();

    //　D列にハイパーリンクを挿入
      sheetCL.getRange(row, CLCOL_MITSUMORI).setFormula(`=HYPERLINK("${fileUrl}", "見積り")`);  // リンクを挿入
    }
  }
}