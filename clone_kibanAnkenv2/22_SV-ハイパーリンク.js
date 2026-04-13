/** 
 * =======================================================================
 * 
 * 【リンク挿入】
 * ・D列に見積りのリンクを入れる
 * ・F列に作業フォルダへのリンクを入れる
 * 
 * ======================================================================== 
 **/ 

function SV_InsertLinks() { 
  Logger.log("〇リンク追加");
  if (startRowSV === null ) return;
  
  let startRow = null;
  let endRow = null;
  const lastRowSV = getLastRowByColumn(sheetSV, SVCOL_MITSUMORI);

  // *** D列に見積のリンクを入れる -----
  for (let i = startRowSV; i <= lastRowSV; i++) {
    const cellDate = sheetSV.getRange(i, 4).getValue();

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
  
  // リンクを入れる
  if (startRow !== null && endRow !== null) {

    for (row = startRow; row <= endRow; row++) {
    //HyperLink(作業フォルダ)
      const searchTermFolder = sheetSV.getRange(row, SVCOL_MITSUMORI-1).getValue().toString().substring(0, 6);
      const folders = DriveApp.searchFolders(`title contains '${searchTermFolder}'`);
      const folder = folders.next();
      const folderUrl = folder.getUrl();
      
    // F列にハイパーリンクを挿入
      sheetSV.getRange(row, SVCOL_FOLDER).setFormula(`=HYPERLINK("${folderUrl}", "リンク")`);  // リンクを挿入
    
    //HyperLink(見積り)
     const searchTermFile = sheetSV.getRange(row, SVCOL_MITSUMORI-1).getValue().toString();
     const files = DriveApp.searchFiles(`title contains '${searchTermFile}'`);
     const file = files.next();
     const fileUrl = file.getUrl();

    //　D列にハイパーリンクを挿入
      sheetSV.getRange(row, SVCOL_MITSUMORI).setFormula(`=HYPERLINK("${fileUrl}", "見積り")`);  // リンクを挿入
    
    Logger.log("リンクを追加しました");
    }
  } else {
    Logger.log("昨日の日付が入力されている行は見つかりませんでした。");
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

function SV_InsertLinksManual() { 
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
    // ▼ 親フォルダのID（「IT基盤構築案件フォルダ」）
    const PARENT_ID = "1kjHPdpN_eL28CbIATueHmfEPSrrOZug0";


    //HyperLink(作業フォルダ)
    const searchTermFolder = sheetSV.getRange(row, SVCOL_MITSUMORI - 1)
                                  .getValue()
                                  .toString()
                                  .slice(0, -3)
                                  .substring(0, 6);

    const parentFolder = DriveApp.getFolderById(PARENT_ID);     // 親フォルダ
    const childFolders = parentFolder.getFolders();             // 配下フォルダ一覧

    let folder = null;

  // 親フォルダ配下のフォルダだけを検索
  while (childFolders.hasNext()) {
    const f = childFolders.next();
    if (f.getName().includes(searchTermFolder)) {
      folder = f;
      break;
    }
  }

if (!folder) {
  throw new Error(`フォルダが見つかりません: ${searchTermFolder}`);
}

const folderUrl = folder.getUrl();
    // F列にハイパーリンクを挿入
      sheetSV.getRange(row, SVCOL_FOLDER ).setFormula(`=HYPERLINK("${folderUrl}", "リンク")`);  // リンクを挿入
    
    //HyperLink(見積り)
     const searchTermFile = sheetSV.getRange(row, SVCOL_MITSUMORI-1).getValue().toString();
     const files = DriveApp.searchFiles(`title contains '${searchTermFile}'`);
     const file = files.next();
     const fileUrl = file.getUrl();

    //　D列にハイパーリンクを挿入
      sheetSV.getRange(row, SVCOL_MITSUMORI).setFormula(`=HYPERLINK("${fileUrl}", "見積り")`);  // リンクを挿入
    }
  }
}