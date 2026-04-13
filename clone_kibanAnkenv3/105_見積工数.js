/**
 * =======================================================================
 * 【作業工数取得（統合版）】
 * - 対象: 案件情報（SV/CL統合）
 * - 条件: 作業依頼日 = 昨日（※不要ならコメントアウト）
 * - 見積(URL) からブックを開く（Google Sheets / Excel 両対応）
 * - 先頭シートの V列から「（総工数）」or「総工数」を探す
 * - その行の U列の値を、案件情報の「工数」列へ転記
 * =======================================================================
 */

function Auto_見積工数取得() {

  const sh = getSH_('案件情報');
  const shSV = getSH_('案件管理表_本社');
  const shCL = getSH_('案件管理表_倉庫');

  const HEADER_ROW = 1;

  const cols = getColsByHeaders_(sh, [
    '作業依頼日',
    '見積',
    '見積工数'
  ], HEADER_ROW);

  const firstRow = HEADER_ROW + 1;
  const lastRow  = getLastRowByColumn_(sh, cols['見積']);
  if (lastRow < firstRow) {
    writeLog_('Auto_見積工数取得', 'データ行がありません');
    Logger.log("データ行がありません");
    return;
  }

  const n = lastRow - firstRow + 1;

  const dateVals      = sh.getRange(firstRow, cols['作業依頼日'], n, 1).getValues();
  const mitsumoriRng  = sh.getRange(firstRow, cols['見積'], n, 1);
  const kosuVals      = sh.getRange(firstRow, cols['見積工数'], n, 1).getValues();

  let ok = 0, skip = 0, ng = 0;

  for (let i = 0; i < n; i++) {
    const r = firstRow + i;

    try {

      // ★ 工数が既に入っている場合はスキップ
      if (kosuVals[i][0] !== "" && kosuVals[i][0] !== null) {
        skip++;
        continue;
      }

      // ★ 昨日案件のみ対象
      if (!isYesterday_(dateVals[i][0], tz)) {
        skip++;
        continue;
      }

      const cell = mitsumoriRng.getCell(i + 1, 1);
      const url  = getUrlFromCell_(cell);

      if (!url) {
        skip++;
        continue;
      }

      const fileId = extractDriveIdFromUrl_(url);
      if (!fileId) throw new Error("fileId抽出失敗");

      const { ssId, tempCreated } = openAsSpreadsheetId_(fileId, r);

      const targetSs = SpreadsheetApp.openById(ssId);
      const targetSh = targetSs.getSheets()[0];

      const foundRow = findTotalKosuRow_(targetSh);
      if (foundRow === -1) {
        throw new Error("『総工数』が見つからない");
      }

      const uValue = targetSh.getRange(foundRow, 21).getValue();

      kosuVals[i][0] = uValue;
      ok++;

      writeLog_('Auto_見積工数取得', `Row ${r}: 総工数=${uValue}`);
      Logger.log(`Row ${r}: 総工数=${uValue}`);

      if (tempCreated) {
        DriveApp.getFileById(ssId).setTrashed(true);
      }

    } catch (e) {
      Logger.log(`Row ${r}: エラー - ${e}`);
      ng++;
    }
  }

  sh.getRange(firstRow, cols['見積工数'], n, 1).setValues(kosuVals);

  writeLog_('Auto_見積工数取得', `ok=${ok}, skip=${skip}, ng=${ng}`);
  Logger.log(`◎完了 ok=${ok}, skip=${skip}, ng=${ng}`);
}

function Manual_見積工数取得() {
  // ** ここを変更する *************************//

  const firstRow = 101;
  const lastRow = 200;

  // ****************************************//

  const sh = getSH_('案件情報');
  const HEADER_ROW = 1;

  const cols = getColsByHeaders_(sh, [
    '作業依頼日',
    '見積',
    '見積工数'
  ], HEADER_ROW);

  const n = lastRow - firstRow + 1;

  const mitsumoriRng  = sh.getRange(firstRow, cols['見積'], n, 1);
  const kosuVals      = sh.getRange(firstRow, cols['見積工数'], n, 1).getValues();

  let ok = 0, skip = 0, ng = 0;

  for (let i = 0; i < n; i++) {
    const r = firstRow + i;

    try {
      // ★ 工数が既に入っている場合はスキップ
      if (kosuVals[i][0] !== "" && kosuVals[i][0] !== null) {
        skip++;
        continue;
      }

      const cell = mitsumoriRng.getCell(i + 1, 1);
      const url  = getUrlFromCell_(cell);

      if (!url) {
        skip++;
        continue;
      }

      const fileId = extractDriveIdFromUrl_(url);
      if (!fileId) throw new Error("fileId抽出失敗");

      const { ssId, tempCreated } = openAsSpreadsheetId_(fileId, r);

      const targetSs = SpreadsheetApp.openById(ssId);
      const targetSh = targetSs.getSheets()[0];

      const foundRow = findTotalKosuRow_(targetSh);
      if (foundRow === -1) {
        throw new Error("『総工数』が見つからない");
      }

      const uValue = targetSh.getRange(foundRow, 21).getValue();

      kosuVals[i][0] = uValue;
      ok++;

      writeLog_('Manual_見積工数取得', `Row ${r}: 総工数=${uValue}`);
      Logger.log(`Row ${r}: 総工数=${uValue}`);

      if (tempCreated) {
        DriveApp.getFileById(ssId).setTrashed(true);
      }

    } catch (e) {
      Logger.log(`Row ${r}: エラー - ${e}`);
      ng++;
    }
  }

  sh.getRange(firstRow, cols['見積工数'], n, 1).setValues(kosuVals);

  writeLog_('Manual_見積工数取得', `ok=${ok}, skip=${skip}, ng=${ng}`);
  Logger.log(`◎完了 ok=${ok}, skip=${skip}, ng=${ng}`);
}

/** Excel/Sheets を Spreadsheet として開けるIDにする（Excelは一時変換） */
function openAsSpreadsheetId_(fileId, rowNoForName) {
  const file = DriveApp.getFileById(fileId);
  const mt = file.getMimeType();

  if (mt === MimeType.GOOGLE_SHEETS) {
    return { ssId: fileId, tempCreated: false };
  }

  // Excel等 → 一時変換（Drive API）
  const blob = file.getBlob();

  // Drive API(v2) の create は name/title 揺れるので両方入れておく
  const converted = Drive.Files.create(
    {
      title: `tempConverted_${rowNoForName}`,
      name:  `tempConverted_${rowNoForName}`,
      mimeType: MimeType.GOOGLE_SHEETS
    },
    blob,
    { convert: true }
  );

  return { ssId: converted.id, tempCreated: true };
}


/** 先頭シートの V列から「（総工数）」/「総工数」を探し、行番号を返す（無ければ -1） */
function findTotalKosuRow_(sheet) {
  const last = sheet.getLastRow();
  if (last <= 0) return -1;

  const V_COL = 22; // V
  const vals = sheet.getRange(1, V_COL, last, 1).getValues();

  for (let i = 0; i < vals.length; i++) {
    const t = String(vals[i][0] ?? '').trim();
    if (t === '（総工数）' || t === '総工数') return i + 1;
  }
  return -1;
}