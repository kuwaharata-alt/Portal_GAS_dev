/**
 * 列名から列数を取得する関数
 **/

function getLastRowByColumn_(sheet, colNumber) {
  const values = sheet.getRange(1, colNumber, sheet.getLastRow()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] != null) {
      return i + 1; // 行番号に戻す
    }
  }
  return 0;
}

/**
 * 開始列、行を入力する関数（手動実行用）
 **/

function runWithPopup() {
  const ui = SpreadsheetApp.getUi();

  // --- 開始行 ---
  const startResp = ui.prompt(
    "開始行の入力",
    "開始行を入力してください：",
    ui.ButtonSet.OK_CANCEL
  );
  if (startResp.getSelectedButton() !== ui.Button.OK) return null;

  const startRow = Number(startResp.getResponseText());
  if (isNaN(startRow) || startRow <= 0) {
    ui.alert("開始行が不正です");
    return null;
  }

  // --- 終了行 ---
  const endResp = ui.prompt(
    "終了行の入力",
    "終了行を入力してください：",
    ui.ButtonSet.OK_CANCEL
  );
  if (endResp.getSelectedButton() !== ui.Button.OK) return null;

  const endRow = Number(endResp.getResponseText());
  if (isNaN(endRow) || endRow < startRow) {
    ui.alert("終了行が不正です");
    return null;
  }

  // --- start / end をセットで返す ---
  return { startRow, endRow };
}


/** 日付セルが「昨日」なら true */
function isYesterday_(d, tz){
  if (!d) return false;
  const ymd = 'yyyy/MM/dd';

  // Date 以外（文字列等）が来ても事故らないように
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return false;

  const now = new Date();
  const today0 = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd') + ' 00:00:00');
  const yesterday0 = new Date(today0.getTime() - 24 * 60 * 60 * 1000);

  return Utilities.formatDate(dt, tz, ymd) === Utilities.formatDate(yesterday0, tz, ymd);
}

/** 指定列(2次元values)から「昨日」の最初の index を返す（無ければ -1） */
function findFirstYesterdayIndex_(dateVals, tz){
  for (let i = 0; i < dateVals.length; i++){
    if (isYesterday_(dateVals[i][0], tz)) return i;
  }
  return -1;
}


/** セルからURLを抜く（リッチテキスト / HYPERLINK関数 / URL直書き 対応） */
function getUrlFromCell_(range) {
  const rt = range.getRichTextValue();
  if (rt) {
    const u = rt.getLinkUrl();
    if (u) return u;

    const runs = rt.getRuns();
    for (const run of runs) {
      const u2 = run.getLinkUrl();
      if (u2) return u2;
    }
  }

  const f = range.getFormula();
  if (f && /^=HYPERLINK\(/i.test(f)) {
    const m = f.match(/^=HYPERLINK\("([^"]+)"/i);
    if (m && m[1]) return m[1];
  }

  const v = range.getValue();
  if (typeof v === 'string' && /^https?:\/\//i.test(v)) return v;

  return '';
}

/** Drive URL からID抽出（file/folder共通） */
function extractDriveIdFromUrl_(url) {
  if (!url) return '';
  const s = String(url);

  let m = s.match(/\/folders\/([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];

  m = s.match(/\/file\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];

  m = s.match(/([a-zA-Z0-9_-]{25,})/);
  return m ? m[1] : '';
}

/** ==========================================================
 * 共通ログ関数
 * シート「実行ログ」にログを書き込む
 * ========================================================== */

function writeLog_(process, message, level = "INFO") {

  const sh = getSH_('実行ログ');
  const now = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');

  sh.appendRow([
    now,        // 日時
    process,    // 処理名
    level,      // レベル
    message     // メッセージ
  ]);

}

/** ==========================================================
 * 必須ヘッダーをまとめて取得する
 * 見つからない場合はエラーにする
 * ========================================================== */

function getColsByHeaders_(sheet, headerNames, headerRow) {
  if (!sheet) throw new Error('getColsByHeaders_: sheet が未指定です');

  if (headerNames == null) {
    throw new Error('getColsByHeaders_: headerNames が未指定です（配列 or 文字列を渡してね）');
  }
  if (typeof headerNames === 'string') headerNames = [headerNames];
  if (!Array.isArray(headerNames)) {
    throw new Error('getColsByHeaders_: headerNames は配列(string[])か文字列(string)で指定してください');
  }

  headerRow = Number(headerRow || 1);

  const map = buildHeaderMap(sheet, headerRow);
  const result = {};

  headerNames.forEach(name => {
    const key = String(name || '').trim();
    if (!key) return;

    if (!map[key]) {
      throw new Error(`ヘッダーが見つかりません: ${key}`);
    }
    result[key] = map[key];
  });

  return result;
}

function buildHeaderMap(sheet, headerRow) {
  if (!sheet) throw new Error('sheet が未指定です');
  headerRow = headerRow || 1;

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};

  const headers = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getValues()[0];

  const map = {};

  headers.forEach((name, i) => {
    const key = String(name || '').trim();
    if (key) {
      map[key] = i + 1; // 1-based
    }
  });

  return map;
}

