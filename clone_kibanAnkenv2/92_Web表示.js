/**
 * SV案件管理表 Viewer (WebApp) - Paged + ConditionalFormat colors (effectiveFormat)
 * - ヘッダー: 1行目
 * - 行数: C列の最終行
 * - 色: Sheets API の effectiveFormat.backgroundColor（条件付き書式反映色）
 */

const CONFIG = {
  sheetId: (typeof sheetId !== "undefined" && sheetId) ? sheetId : "＜スプレッドシートIDを入れてください＞",
  sheetName: (typeof vSV !== "undefined" && vSV) ? vSV : "SV案件管理表",

  defaultLimit: 100,
  lastRowCacheSeconds: 120, // C列最終行キャッシュ（数値なので安全）
  headerCacheSeconds: 600,  // ヘッダー/最終列キャッシュ
};

function doGet(e) {
  return HtmlService.createTemplateFromFile("92_index")
    .evaluate()
    .setTitle("SV案件管理表 Viewer (Paged+Format)")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTML から呼ばれる
 * offset/limit: ページング
 */
function api_getPagedTable(offset, limit) {
  offset = toInt_(offset, 0);
  limit  = toInt_(limit, CONFIG.defaultLimit);

  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${CONFIG.sheetName}`);

  const headerInfo = getHeaderInfoCached_(); // headers + lastCol
  const lastRowByC = getLastRowByColumnCached_(sheet, 3); // C列基準
  const totalDataRows = Math.max(0, lastRowByC - 1); // 2行目〜データ

  const safeOffset = Math.max(0, Math.min(offset, totalDataRows));
  // ここより上で totalDataRows を計算している前提
  let safeLimit;
  if (limit === -1) {
    // ★全件：残り全部（offset考慮）
    safeLimit = totalDataRows; // または remaining を使ってもOK
  } else {
    // 通常は上限200などに制限
    safeLimit = Math.max(1, Math.min(limit, 200));
  }

  // 取得するシート上の開始行（2行目が offset=0）
  const startRow = 2 + safeOffset;
  const remaining = totalDataRows - safeOffset;
  const fetchCount = Math.max(0, Math.min(safeLimit, remaining));

  let rows = [];
  let bgs = [];
  let links = []; // ★追加：リンクURLの2次元配列

  if (fetchCount > 0) {
    const a1 = `${CONFIG.sheetName}!${a1Range_(1, startRow)}:${a1Range_(headerInfo.lastCol, startRow + fetchCount - 1)}`;

    const res = Sheets.Spreadsheets.get(CONFIG.sheetId, {
      ranges: [a1],
      includeGridData: true,
      fields: "sheets(data(rowData(values(formattedValue,hyperlink,textFormatRuns,effectiveFormat(backgroundColor)))))"
    });

    const rowData = (((res || {}).sheets || [])[0] || {}).data?.[0]?.rowData || [];

    for (let r = 0; r < fetchCount; r++) {
      const v = rowData[r]?.values || [];
      const oneRow   = new Array(headerInfo.lastCol).fill("");
      const oneBg    = new Array(headerInfo.lastCol).fill("");
      const oneLinks = new Array(headerInfo.lastCol).fill(""); // ★追加

      for (let c = 0; c < headerInfo.lastCol; c++) {
        const cell = v[c] || {};

        oneRow[c] = (cell.formattedValue ?? "") + "";
        oneBg[c]  = colorToHex_(cell.effectiveFormat?.backgroundColor);

        // ★リンク抽出（優先：cell.hyperlink → textFormatRuns内のlink）
        let url = cell.hyperlink || "";

        if (!url && cell.textFormatRuns && cell.textFormatRuns.length) {
          // textFormatRuns: [{startIndex, format:{link:{uri}}}, ...]
          for (const run of cell.textFormatRuns) {
            const uri = run?.format?.link?.uri;
            if (uri) { url = uri; break; }
          }
        }

        oneLinks[c] = url || "";
      }

      rows.push(oneRow);
      bgs.push(oneBg);
      links.push(oneLinks); // ★追加
    }
  }

  return {
    ok: true,
    sheetName: CONFIG.sheetName,
    updatedAt: new Date().toISOString(),
    headers: headerInfo.headers,
    rows,
    bgs,
    links,         // ★追加
    offset: safeOffset,
    limit: safeLimit,
    total: totalDataRows
  };
}

/** ヘッダー行（1行目）と最終列をキャッシュして返す */
function getHeaderInfoCached_() {
  const cache = CacheService.getScriptCache();
  const key = `SV_HDR_${CONFIG.sheetId}_${CONFIG.sheetName}`;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  // ヘッダーは値だけ取れればいいので values API で軽く読む
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  const maxCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, maxCol).getValues()[0];

  const lastCol = Math.max(1, getLastColByHeader_(headerRow));
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h ?? "").trim());

  const obj = { headers, lastCol };
  cache.put(key, JSON.stringify(obj), CONFIG.headerCacheSeconds);
  return obj;
}

/** C列の最終行（キャッシュ付き） */
function getLastRowByColumnCached_(sheet, col) {
  const cache = CacheService.getScriptCache();
  const key = `SV_LASTROW_${CONFIG.sheetId}_${CONFIG.sheetName}_C${col}`;
  const cached = cache.get(key);
  if (cached) return parseInt(cached, 10);

  const upper = sheet.getLastRow();
  if (upper < 1) return 0;

  const colVals = sheet.getRange(1, col, upper, 1).getValues();
  let last = 0;
  for (let i = colVals.length - 1; i >= 0; i--) {
    const v = colVals[i][0];
    if (v !== "" && v !== null) { last = i + 1; break; }
  }

  cache.put(key, String(last), CONFIG.lastRowCacheSeconds);
  return last;
}

/** ヘッダー行配列から、最後の非空列番号（1-based） */
function getLastColByHeader_(headerRow) {
  for (let i = headerRow.length - 1; i >= 0; i--) {
    const v = headerRow[i];
    if (v !== "" && v !== null) return i + 1;
  }
  return 0;
}

function toInt_(v, def) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : def;
}

/** (col,row) -> A1  */
function a1Range_(col, row) {
  return `${colToA1_(col)}${row}`;
}
function colToA1_(col) {
  let n = col;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Sheets API backgroundColor {red,green,blue} (0..1) -> "#RRGGBB" / "" */
function colorToHex_(bg) {
  if (!bg) return "";
  const r = clamp255_(Math.round((bg.red ?? 1) * 255));
  const g = clamp255_(Math.round((bg.green ?? 1) * 255));
  const b = clamp255_(Math.round((bg.blue ?? 1) * 255));
  // 「真っ白(#ffffff)」も返す（必要なら空扱いに変更可）
  return "#" + [r,g,b].map(x => x.toString(16).padStart(2, "0")).join("");
}
function clamp255_(n) {
  return Math.max(0, Math.min(255, n));
}

/** B列ステータスの候補一覧（ユニーク）を返す：プルダウン用 */
function api_getStatusOptions() {
  const cache = CacheService.getScriptCache();
  const key = `SV_STATUS_OPTS_${CONFIG.sheetId}_${CONFIG.sheetName}`;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${CONFIG.sheetName}`);

  const lastRowByC = getLastRowByColumnCached_(sheet, 3); // C列基準
  const totalDataRows = Math.max(0, lastRowByC - 1);
  if (totalDataRows <= 0) return { ok: true, options: [] };

  // B列（2列目）だけを一気に取る（値だけなので軽い）
  const vals = sheet.getRange(2, 2, totalDataRows, 1).getDisplayValues().flat();

  const set = new Set();
  for (const v of vals) {
    const s = String(v ?? "").trim();
    if (s) set.add(s);
  }

  // 並び：文字列ソート（必要なら自然順にもできる）
  const options = Array.from(set).sort((a,b) => a.localeCompare(b, "ja"));

  const res = { ok: true, options };
  cache.put(key, JSON.stringify(res), 600); // 10分キャッシュ（小さいので安全）
  return res;
}
