/** =======================================================================
 *
 * 【前日分の作業依頼をチェックし、参照先(C列)に値が出るまで待つ】
 *
 * 1) 作業依頼 A列に「前日」があるか確認（なければ return）
 * 2) 前日行のうち、G列が重複している値は対象外（全て除外）
 * 3) 残ったG列(ユニーク)について、K列で参照先を分岐
 *    - PC系 → CL案件管理表
 *    - SV系 → SV案件管理表
 * 4) 参照先の C列（C3のFILTERスピル含む）に G値が存在するか確認
 * 5) 無ければ、反映されるまで待機（flush + sleep でリトライ）
 *
 * -----------------------------------------------------------------------
 * 期待するシート:
 *  - 作業依頼
 *  - CL案件管理表
 *  - SV案件管理表
 *
 * 注意:
 *  - 「更新されるまで待つ」は最大 waitMaxMs まで
 *  - 参照先が FILTER で生成される想定のため、C3以降の表示値を重点的に確認
 *
 * ======================================================================= */
function CheckRequest() {
  const tz = 'Asia/Tokyo';

  // ---- 必要ならここをあなたの共通変数に差し替えOK ----
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetReq = ss.getSheetByName('作業依頼');
  const sheetCL  = ss.getSheetByName('CL案件管理表');
  const sheetSV  = ss.getSheetByName('SV案件管理表');
  // --------------------------------------------------

  if (!sheetReq || !sheetCL || !sheetSV) {
    throw new Error('対象シートが見つかりません（作業依頼 / CL案件管理表 / SV案件管理表）');
  }

  // 前日（00:00:00に丸め）
  const yesterday = new Date();
  yesterday.setHours(0, 0, 0, 0);
  yesterday.setDate(yesterday.getDate() - 1);
  const yStr = Utilities.formatDate(yesterday, tz, 'yyyy/MM/dd');

  // 作業依頼のデータ取得（A:K くらいあれば足りる想定）
  const lastRow = sheetReq.getLastRow();
  if (lastRow < 2) return;

  const values = sheetReq.getRange(2, 1, lastRow - 1, 11).getValues(); // A(1)～K(11)
  // A=0, G=6, K=10
  const rowsYesterday = [];

  for (let i = 0; i < values.length; i++) {
    const a = values[i][0]; // A列
    const g = values[i][6]; // G列
    const k = values[i][10]; // K列

    if (!a) continue;

    // 日付判定（A列がDateでも文字でもOKにする）
    const aStr = normalizeToYmd_(a, tz);
    if (aStr !== yStr) continue;

    if (!g) continue; // Gが空なら対象外（必要なら外してOK）

    rowsYesterday.push({
      rowNumber: i + 2, // 実際の行番号
      g: String(g).trim(),
      k: (k === null || k === undefined) ? '' : String(k).trim(),
    });
  }

  // 1) 前日の日付が無い → 何もしない
  if (rowsYesterday.length === 0) return;

  // 2) G列の重複チェック（前日行の中で重複している値は全て除外）
  const gCount = {};
  rowsYesterday.forEach(r => {
    gCount[r.g] = (gCount[r.g] || 0) + 1;
  });

  const candidates = rowsYesterday.filter(r => gCount[r.g] === 1);

  // 重複だらけで候補なし → 何もしない
  if (candidates.length === 0) return;

  // 3) K列で参照先シートを分岐 → 4) C列に出るまで待つ
  // 待機設定（必要なら調整）
  const waitMaxMs = 60 * 1000;  // 最大60秒待つ
  const intervalMs = 500;       // 0.5秒ごとに確認

  // 同じGが複数候補に入ることは無い設計（上で重複排除済み）
  for (const item of candidates) {
    const type = item.k; // "PC系" or "SV系"
    const targetSheet =
      type === 'PC系' ? sheetCL :
      type === 'SV系' ? sheetSV : null;

    if (!targetSheet) {
      // Kが想定外ならスキップ（必要なら throw に変更OK）
      Logger.log(`SKIP: 作業依頼 行${item.rowNumber} K列が想定外: ${type}`);
      continue;
    }

    const found = waitUntilValueAppearsInColumnC_(targetSheet, item.g, {
      waitMaxMs,
      intervalMs,
      startRow: 1,      // C列全体を確認したいなら1
      // FILTERの結果がC3から出るなら 3 にしてもOK
      // startRow: 3,
    });

    if (!found) {
      Logger.log(`TIMEOUT: ${targetSheet.getName()} のC列に "${item.g}" が出るのを待ちましたが、時間切れ`);
      // ここで return するか続行するかは好み
      // return;
    }
  }
}

/** A列などの「Date / 文字列」を yyyy/MM/dd に正規化 */
function normalizeToYmd_(v, tz) {
  if (v instanceof Date) {
    const d = new Date(v);
    d.setHours(0, 0, 0, 0);
    return Utilities.formatDate(d, tz, 'yyyy/MM/dd');
  }
  const s = String(v).trim();
  if (!s) return '';

  // 文字列日付のゆるい吸収（yyyy/MM/dd, yyyy-MM-dd, yyyy.MM.dd 等）
  const m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (m) {
    const yy = m[1];
    const mm = ('0' + m[2]).slice(-2);
    const dd = ('0' + m[3]).slice(-2);
    return `${yy}/${mm}/${dd}`;
  }

  // Dateにパースできるならそれを使う
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    parsed.setHours(0, 0, 0, 0);
    return Utilities.formatDate(parsed, tz, 'yyyy/MM/dd');
  }

  return s; // 最後の手段
}

/**
 * 参照先シートの「C列」に targetValue が現れるまで待つ
 * - FILTERのスピル結果が遅れて出るケース用に flush + sleep でポーリング
 * - まず表示値(getDisplayValues)で確認（FILTER結果を拾いやすい）
 */
function waitUntilValueAppearsInColumnC_(sheet, targetValue, opt = {}) {
  const waitMaxMs  = opt.waitMaxMs ?? 60000;
  const intervalMs = opt.intervalMs ?? 500;
  const startRow   = opt.startRow ?? 1;

  const target = String(targetValue).trim();
  const start = Date.now();

  while (Date.now() - start < waitMaxMs) {
    SpreadsheetApp.flush(); // 再計算/反映を促す

    const lastRow = sheet.getLastRow();
    if (lastRow >= startRow) {
      const numRows = lastRow - startRow + 1;

      // C列だけ取得（表示値でチェック）
      const colC = sheet.getRange(startRow, 3, numRows, 1).getDisplayValues(); // C=3
      if (containsIn2D_(colC, target)) return true;
    }

    Utilities.sleep(intervalMs);
  }
  return false;
}

/** 2次元配列に target が含まれるか（trim一致） */
function containsIn2D_(arr2d, target) {
  for (let r = 0; r < arr2d.length; r++) {
    const v = (arr2d[r][0] === null || arr2d[r][0] === undefined) ? '' : String(arr2d[r][0]).trim();
    if (v === target) return true;
  }
  return false;
}
