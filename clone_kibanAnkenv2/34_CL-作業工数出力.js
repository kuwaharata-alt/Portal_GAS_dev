/** 
 * =======================================================================
 * 
 * 【作業工数登録】
 * 1.各行の D列のHYPERLINK から リンク先ブックを開く
 * 2.一番左のシートのV列から「（総工数）」の行を探す
 * 3.その行のU列の値(工数)をK列へ転記する
 * 
 * ======================================================================== 
 **/ 

function CL_TotalWork() {
  Logger.log("〇作業工数追加");
  const startRowCL =  340;
  if (startRowCL === null ) return;
  for (let row = startRowCL; row <= lastRowCL; row++) {
    try {
      const sheetName = sheetCL.getRange(row, 3).getDisplayValue(); // C列
      const linkCell  = sheetCL.getRange(row, 4);                   // D列：HYPERLINK

      // C列も D列も空ならスキップ
      if (!sheetName || !linkCell.getDisplayValue()) {
        continue;
      }

      // ▼ URL → fileId
      const url    = getUrlFromHyperlink(linkCell);
      const fileId = extractFileId(url);

      // ▼ スプシ or Excel 判定
      const file = DriveApp.getFileById(fileId);
      let targetSsId;

      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        targetSsId = fileId;  // すでにスプレッドシート
      } else {
        // Excel → スプシへ一時変換
        const blob = file.getBlob();
        const converted = Drive.Files.create(
          {
            name: 'tempConverted_' + row,
            mimeType: MimeType.GOOGLE_SHEETS
          },
          blob,
          { convert: true }
        );
        targetSsId = converted.id;
      }

      // ▼ リンク先シートを取得
      const targetSs = SpreadsheetApp.openById(targetSsId);
      const sh = targetSs.getSheets()[0]; 

      // ▼ V列から「（総工数）」を探す
      const targetLastRow = sh.getLastRow();
      const vValues = sh.getRange(1, 22, targetLastRow, 1).getValues(); // V列

      let foundRow = -1;

      for (let i = 0; i < targetLastRow; i++) {
        const cellText = String(vValues[i][0]).trim();
        if (cellText === '（総工数）' || cellText === '総工数') {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow === -1) {
        throw new Error("『（総工数）』が見つからない");
      }

      // ▼ 見つかった行の U列（21列目）を取得
      const uValue = sh.getRange(foundRow, 21).getValue();

      // ▼ 元シートの K列（11列）に書き込み
      sheetCL.getRange(row, 11).setValue(uValue);

      Logger.log(row + " 行目：総工数=" + uValue);

      // ▼ Excel 変換時は一時ファイル削除
      if (targetSsId !== fileId) {
        DriveApp.getFileById(targetSsId).setTrashed(true);
      }

    } catch (err) {
      Logger.log("行 " + row + " でエラー：" + err);
      // ※必ず処理は継続する
    }

  }

  Logger.log("完了！");
}