/**
 * 作業依頼を起点に 案件情報 を更新
 * - 新規見積番号は追加
 * - 既存見積番号は空欄セルのみ補完
 * - 処理済みの作業依頼行は ScriptProperties で管理
 */
function Auto_案件情報作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shReq = ss.getSheetByName('作業依頼');
  const shKeijo = ss.getSheetByName('計上管理_転写');
  const shCases = ss.getSheetByName('案件情報');
  const TZ = Session.getScriptTimeZone() || 'Asia/Tokyo';

  if (!shReq) throw new Error('シート「作業依頼」が見つかりません');
  if (!shKeijo) throw new Error('シート「計上管理_転写」が見つかりません');
  if (!shCases) throw new Error('シート「案件情報」が見つかりません');

  // ===== 案件情報ヘッダー =====
  const casesHeaderMap = buildHeaderMapFromSheet_(shCases);
  const CASES_KEY = mustCol_(casesHeaderMap, '見積番号');

  const CASES_COL = {
    顧客名: casesHeaderMap['顧客名'] || null,
    作業区分: casesHeaderMap['作業区分'] || null,
    ｶﾃｺﾞﾘｰ: casesHeaderMap['ｶﾃｺﾞﾘｰ'] || null,
    担当営業: casesHeaderMap['担当営業'] || null,
    担当プリ: casesHeaderMap['担当プリ'] || null,
    案件概要: casesHeaderMap['案件概要'] || null,
    作業依頼日: casesHeaderMap['作業依頼日'] || null,
    検収予定: casesHeaderMap['検収予定'] || null,
  };

  // ===== 作業依頼 固定列 =====
  const REQ = {
    DATE: 1,        // A 日付
    ESTIMATE: 7,    // G 見積番号
    CUSTOMER: 8,    // H 顧客名
    SALES: 9,       // I 担当営業
    WORK_TYPE: 11,  // K 作業区分
  };

  // ===== 1) 作業依頼の差分取得 =====
  const props = PropertiesService.getScriptProperties();
  const propKey = 'WR_LAST_PROCESSED_ROW';
  let startRow = Number(props.getProperty(propKey) || '2');
  if (!startRow || startRow < 2) startRow = 2;

  const reqLastRow = shReq.getLastRow();
  if (reqLastRow < startRow) {
    Logger.log('作業依頼：新規行なし');
    return;
  }

  const numRows = reqLastRow - startRow + 1;
  const numCols = Math.max(REQ.DATE, REQ.ESTIMATE, REQ.CUSTOMER, REQ.SALES, REQ.WORK_TYPE);
  const reqValues = shReq.getRange(startRow, 1, numRows, numCols).getValues();

  const reqMap = new Map();
  const processed = new Set();
  let skippedBlankKey = 0;
  let skippedDup = 0;

  for (const row of reqValues) {
    const key = normalizeEstimateNo_(row[REQ.ESTIMATE - 1]);
    if (!key) {
      skippedBlankKey++;
      continue;
    }
    if (processed.has(key)) {
      skippedDup++;
      continue;
    }
    processed.add(key);

    reqMap.set(key, {
      顧客名: row[REQ.CUSTOMER - 1],
      担当営業: row[REQ.SALES - 1],
      作業依頼日: formatDateYMD_(row[REQ.DATE - 1], TZ),
      作業区分: row[REQ.WORK_TYPE - 1],
    });
  }

  if (reqMap.size === 0) {
    props.setProperty(propKey, String(reqLastRow + 1));
    Logger.log(`作業依頼：有効キー0件（空キー=${skippedBlankKey}, 重複=${skippedDup}）`);
    return;
  }

  // ===== 2) 計上管理_転写 補助Map =====
  const keijoHeaderMap = buildHeaderMapFromSheet_(shKeijo);
  const keijoKeyCol = findFirstCol_(keijoHeaderMap, ['管理番号']);
  if (!keijoKeyCol) {
    throw new Error('計上管理_転写 にキー列「管理番号」が見つかりません');
  }

  const KJ = {
    担当PS: keijoHeaderMap['担当PS'] || null,
    案件概要: keijoHeaderMap['案件概要'] || null,
    計上月: keijoHeaderMap['計上月'] || null,
    ｶﾃｺﾞﾘｰ: keijoHeaderMap['ｶﾃｺﾞﾘｰ'] || null,
  };

  const keijoMap = new Map();
  const keijoLastRow = shKeijo.getLastRow();

  if (keijoLastRow >= 2) {
    const keijoLastCol = shKeijo.getLastColumn();
    const values = shKeijo.getRange(2, 1, keijoLastRow - 1, keijoLastCol).getValues();

    for (const row of values) {
      const key = normalizeEstimateNo_(row[keijoKeyCol - 1]);
      if (!key) continue;

      keijoMap.set(key, {
        担当プリ: KJ.担当PS ? row[KJ.担当PS - 1] : '',
        案件概要: KJ.案件概要 ? row[KJ.案件概要 - 1] : '',
        検収予定: KJ.計上月 ? formatKeijoMonth_(row[KJ.計上月 - 1], TZ) : '',
        ｶﾃｺﾞﾘｰ: KJ.ｶﾃｺﾞﾘｰ ? row[KJ.ｶﾃｺﾞﾘｰ - 1] : '',
      });
    }
  }

  // ===== 3) 案件情報 既存キー -> 行番号 =====
  const casesLastRow = shCases.getLastRow();
  const casesKeyToRow = new Map();

  if (casesLastRow >= 2) {
    const keys = shCases.getRange(2, CASES_KEY, casesLastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      const key = normalizeEstimateNo_(keys[i][0]);
      if (key) casesKeyToRow.set(key, i + 2);
    }
  }

  // ===== 4) 案件情報へ反映 =====
  const appendRows = [];
  const width = shCases.getLastColumn();

  let appended = 0;
  let updated = 0;
  let noChange = 0;

  for (const [key, reqObj] of reqMap.entries()) {
    const addOn = keijoMap.get(key) || {};
    const src = { ...reqObj, ...addOn };
    const existingRow = casesKeyToRow.get(key);

    if (!existingRow) {
      const row = new Array(width).fill('');
      row[CASES_KEY - 1] = key;

      setIfColExists_(row, CASES_COL.顧客名, src.顧客名);
      setIfColExists_(row, CASES_COL.作業区分, src.作業区分);
      setIfColExists_(row, CASES_COL.ｶﾃｺﾞﾘｰ, src.ｶﾃｺﾞﾘｰ);
      setIfColExists_(row, CASES_COL.担当営業, src.担当営業);
      setIfColExists_(row, CASES_COL.担当プリ, src.担当プリ);
      setIfColExists_(row, CASES_COL.案件概要, src.案件概要);
      setIfColExists_(row, CASES_COL.作業依頼日, src.作業依頼日);
      setIfColExists_(row, CASES_COL.検収予定, src.検収予定);

      appendRows.push(row);
      appended++;
      continue;
    }

    const currentValues = shCases.getRange(existingRow, 1, 1, width).getValues()[0];
    const writes = [];

    pushIfBlankByValue_(writes, currentValues, CASES_KEY, key);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.顧客名, src.顧客名);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.作業区分, src.作業区分);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.ｶﾃｺﾞﾘｰ, src.ｶﾃｺﾞﾘｰ);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.担当営業, src.担当営業);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.担当プリ, src.担当プリ);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.案件概要, src.案件概要);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.作業依頼日, src.作業依頼日);
    pushIfBlankByValue_(writes, currentValues, CASES_COL.検収予定, src.検収予定);

    if (writes.length > 0) {
      for (const w of writes) {
        shCases.getRange(existingRow, w.col).setValue(w.value);
      }
      updated++;
    } else {
      noChange++;
    }
  }

  if (appendRows.length > 0) {
    const lastDataRow = getLastDataRowByCol_(shCases, CASES_KEY);
    shCases.getRange(lastDataRow + 1, 1, appendRows.length, width).setValues(appendRows);
  }

  props.setProperty(propKey, String(reqLastRow + 1));

  writeLog_(
    `案件情報更新`,
    `案件情報Upsert完了: 追加=${appended}, 更新=${updated}, 変更なし=${noChange}, ` +
    `作業依頼処理=${startRow}..${reqLastRow}, 空キー=${skippedBlankKey}, 重複=${skippedDup}`
  );
}

/**
 * 案件情報 から 案件管理表へ転記
 * - 作業区分 = 本社 → 案件管理表_本社-現地
 * - 作業区分 = 倉庫 → 案件管理表_倉庫
 * - 既に転記済みの見積番号はスキップ
 * - 最終行の次に追記のみ
 */
function Auto_案件情報To案件管理表_転記() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const srcSheet = ss.getSheetByName('案件情報');
  const headOfficeSheet = ss.getSheetByName('案件管理表_本社-現地');
  const warehouseSheet = ss.getSheetByName('案件管理表_倉庫');

  if (!srcSheet) throw new Error('「案件情報」シートが見つかりません。');
  if (!headOfficeSheet) throw new Error('「案件管理表_本社-現地」シートが見つかりません。');
  if (!warehouseSheet) throw new Error('「案件管理表_倉庫」シートが見つかりません。');

  const srcValues = srcSheet.getDataRange().getValues();
  if (srcValues.length < 2) {
    Logger.log('案件情報にデータがありません。');
    return;
  }

  const srcHeader = srcValues[0];
  const srcData = srcValues.slice(1);

  const srcCol = buildHeaderMapFromRowZeroBased_(srcHeader);
  const headOfficeCol = buildHeaderMapFromRowZeroBased_(
    headOfficeSheet.getRange(2, 1, 1, headOfficeSheet.getLastColumn()).getValues()[0]
  );
  const warehouseCol = buildHeaderMapFromRowZeroBased_(
    warehouseSheet.getRange(2, 1, 1, warehouseSheet.getLastColumn()).getValues()[0]
  );

  const transferHeaders = [
    '見積番号',
    '顧客名',
  ];

  const headOfficeKeyToRow = getEstimateRowMap_(headOfficeSheet, '見積番号');
  const warehouseKeyToRow = getEstimateRowMap_(warehouseSheet, '見積番号');

  let addHead = 0;
  let addWare = 0;
  let updHead = 0;
  let updWare = 0;

  srcData.forEach(row => {
    const estimateNo = normalizeEstimateNo_(row[srcCol['見積番号']]);
    const workType = String(row[srcCol['作業区分']] || '').trim();

    if (!estimateNo) return;

    if (workType === '本社') {
      const rowNo = headOfficeKeyToRow.get(estimateNo);
      if (rowNo) {
        if (updateDestinationBlankCells_(headOfficeSheet, rowNo, row, srcCol, headOfficeCol, transferHeaders)) {
          updHead++;
        }
      } else {
        const newRow = buildRowForDestination_(row, srcCol, headOfficeCol, transferHeaders);
        headOfficeSheet
          .getRange(headOfficeSheet.getLastRow() + 1, 1, 1, headOfficeSheet.getLastColumn())
          .setValues([newRow]);
        addHead++;
      }
    } else if (workType === '倉庫') {
      const rowNo = warehouseKeyToRow.get(estimateNo);
      if (rowNo) {
        if (updateDestinationBlankCells_(warehouseSheet, rowNo, row, srcCol, warehouseCol, transferHeaders)) {
          updWare++;
        }
      } else {
        const newRow = buildRowForDestination_(row, srcCol, warehouseCol, transferHeaders);
        warehouseSheet
          .getRange(warehouseSheet.getLastRow() + 1, 1, 1, warehouseSheet.getLastColumn())
          .setValues([newRow]);
        addWare++;
      }
    }
  });

  writeLog_('案件情報To案件管理表_転記', `本社-現地 追加=${addHead}, 更新=${updHead}`);
  writeLog_('案件情報To案件管理表_転記', `倉庫 追加=${addWare}, 更新=${updWare}`);
}

/**
 * 差分ポインタとデータをリセット
 * - WR_LAST_PROCESSED_ROW を削除
 * - 案件情報 / 案件管理表_本社-現地 / 案件管理表_倉庫 の2行目以降をクリア
 */
function Manual_Gen_Upsertリセット() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TARGET_SHEETS = [
    '案件情報',
    '案件管理表_本社-現地',
    '案件管理表_倉庫',
  ];

  PropertiesService.getScriptProperties().deleteProperty('WR_LAST_PROCESSED_ROW');

  TARGET_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) {
      Logger.log(`スキップ: シート「${name}」が見つかりません`);
      return;
    }
    clearSheetBodyValues_(sh);
    Logger.log(`シート「${name}」の2行目以降をクリアしました`);
  });

  Logger.log('Upsert関連のリセットが完了しました');
}

/* =========================================================
 * helper
 * ========================================================= */

/**
 * 1行目ヘッダーを {ヘッダー名: 1-based列番号} で返す
 */
function buildHeaderMapFromSheet_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return buildHeaderMapFromRowOneBased_(headerRow);
}

/**
 * ヘッダー配列を {ヘッダー名: 1-based列番号} で返す
 */
function buildHeaderMapFromRowOneBased_(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i] ?? '').trim();
    if (h) map[h] = i + 1;
  }
  return map;
}

/**
 * ヘッダー配列を {ヘッダー名: 0-based列番号} で返す
 */
function buildHeaderMapFromRowZeroBased_(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i] ?? '').trim();
    if (h) map[h] = i;
  }
  return map;
}

function mustCol_(headerMap, name) {
  const col = headerMap[name];
  if (!col) throw new Error(`ヘッダー「${name}」が見つかりません（1行目を確認）`);
  return col;
}

function findFirstCol_(headerMap, candidates) {
  for (const name of candidates) {
    if (headerMap[name]) return headerMap[name];
  }
  return null;
}

/**
 * 見積番号を正規化
 * - ダッシュ記号統一
 * - 末尾1桁数字は 0埋め
 */
function normalizeEstimateNo_(value) {
  let t = String(value ?? '').trim();
  if (!t) return '';

  t = t.replace(/[‐-‒–—―ｰ－]/g, '-').replace(/\s+/g, '');

  const m1 = t.match(/^(\d{6})-(\d)$/);
  if (m1) return `${m1[1]}-0${m1[2]}`;

  return t;
}

function isBlank_(v) {
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

function formatDateYMD_(value, tz) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, tz || 'Asia/Tokyo', 'yyyy/MM/dd');
}

function formatKeijoMonth_(value, tz) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz || 'Asia/Tokyo', 'yyyy/MM');
  }

  const s = String(value ?? '').trim();
  if (!s) return '';

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, tz || 'Asia/Tokyo', 'yyyy/MM');
  }

  return s;
}

function setIfColExists_(rowArr, col, value) {
  if (!col) return;
  if (isBlank_(value)) return;
  rowArr[col - 1] = value;
}

function pushIfBlankByValue_(writes, currentRowValues, col, value) {
  if (!col) return;
  if (isBlank_(value)) return;

  const currentValue = currentRowValues[col - 1];
  if (isBlank_(currentValue)) {
    writes.push({ col, value });
  }
}

function getLastDataRowByCol_(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const v = values[i][0];
    if (!isBlank_(v)) return i + 2;
  }
  return 1;
}

function normalizeHeader_(s) {
  return String(s ?? '')
    .replace(/[\r\n]/g, '')
    .replace(/\s+/g, '')
    .trim();
}

function buildHeaderMapFromRowZeroBased_(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    const raw = String(headerRow[i] ?? '').trim();
    const norm = normalizeHeader_(raw);
    if (norm) map[norm] = i;
  }
  return map;
}

function buildHeaderMapFromRowOneBased_(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    const raw = String(headerRow[i] ?? '').trim();
    const norm = normalizeHeader_(raw);
    if (norm) map[norm] = i + 1;
  }
  return map;
}

function getExistingEstimateNos_(sheet, headerName) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3) return new Set();

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const headerMap = buildHeaderMapFromRowZeroBased_(headers);

  const key = normalizeHeader_(headerName);
  if (headerMap[key] === undefined) {
    throw new Error(`シート「${sheet.getName()}」に「${headerName}」列がありません。`);
  }

  const colIndex = headerMap[key] + 1;
  const values = sheet.getRange(3, colIndex, lastRow - 2, 1).getValues().flat();

  return new Set(
    values.map(v => normalizeEstimateNo_(v)).filter(v => v !== '')
  );
}

function getEstimateRowMap_(sheet, headerName) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const map = new Map();
  if (lastRow < 3) return map;

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const headerMap = buildHeaderMapFromRowZeroBased_(headers);
  const keyCol = headerMap[headerName];
  if (keyCol === undefined) {
    throw new Error(`シート「${sheet.getName()}」に「${headerName}」列がありません。`);
  }

  const values = sheet.getRange(3, keyCol + 1, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const key = normalizeEstimateNo_(values[i][0]);
    if (key) map.set(key, i + 3);
  }
  return map;
}

function updateDestinationBlankCells_(sheet, destRow, srcRow, srcColMap, destColMap, transferHeaders) {
  const destWidth = sheet.getLastColumn();
  const current = sheet.getRange(destRow, 1, 1, destWidth).getValues()[0];
  let changed = false;

  transferHeaders.forEach(header => {
    if (srcColMap[header] === undefined || destColMap[header] === undefined) return;

    const srcVal = srcRow[srcColMap[header]];
    const destIdx = destColMap[header];

    if (isBlank_(current[destIdx]) && !isBlank_(srcVal)) {
      current[destIdx] = srcVal;
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(destRow, 1, 1, destWidth).setValues([current]);
  }

  return changed;
}

function buildRowForDestination_(srcRow, srcColMap, destColMap, transferHeaders) {
  const destWidth = Math.max(...Object.values(destColMap)) + 1;
  const output = new Array(destWidth).fill('');

  transferHeaders.forEach(header => {
    if (destColMap[header] !== undefined && srcColMap[header] !== undefined) {
      output[destColMap[header]] = srcRow[srcColMap[header]];
    }
  });

  return output; // ← 1行分の配列を返す
}

function clearSheetBodyValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return;

  sheet.getRange(3, 1, lastRow - 1, lastCol).clearContent();
}