/**
 * フォームの回答（C列=案件番号6桁）をキーに、
 * 別ファイル「計上管理」（B列=管理番号 / ヘッダー3行目）から
 * 指定列だけ抽出し、「計上管理_写し」に同期する
 */
function Auto_計上管理転写() {

  // ===== 設定 =====
  const FORM_KEY_COL = 7;
  const FORM_HEADER_ROW = 1;

  const SOURCE_SS_ID = "1mDMUFShbF9Zmq6VC03_yNHpBUExNiVoqJ2nLZyt6gM8";
  const SOURCE_SHEET_NAME = "計上管理";
  const SOURCE_HEADER_ROW = 3;   // ★ ヘッダーは3行目
  const SOURCE_KEY_COL = 20;  // ★ 受注見積もり

  const DEST_SHEET_NAME = "計上管理_転写";

  // 転記する列（この順で出力）
  const TARGET_COLS = [
    "管理番号",
    "部署",
    "営業",
    "顧客名",
    "案件概要",
    "案件総額",
    "計上月",
    "作業カテゴリ",
    "担当PS",
    "担当SE",
    "ｶﾃｺﾞﾘｰ",
    "受注見積番号",
    "作業の補足情報・留意事項",
  ];

  // ===== 作業依頼から案件番号を取得 =====
  const shreq = getSH_('作業依頼');
  if (!shreq) throw new Error(`フォームシートが見つかりません: ${shreq}`);

  const formLastRow = shreq.getLastRow();
  if (formLastRow <= FORM_HEADER_ROW) return;

  const keySet = new Set(
    shreq
      .getRange(FORM_HEADER_ROW + 1, FORM_KEY_COL, formLastRow - FORM_HEADER_ROW, 1)
      .getValues()
      .flat()
      .map(v => String(v ?? "").trim())
  );

  if (keySet.size === 0) {
    writeLog_('Auto_計上管理転写', "フォームに6桁案件番号がありません。");
    return;
  }

  // ===== コピー元 =====
  const srcSs = SpreadsheetApp.openById(SOURCE_SS_ID);
  const srcSheet = srcSs.getSheetByName(SOURCE_SHEET_NAME);
  if (!srcSheet) throw new Error(`コピー元シートが見つかりません: ${SOURCE_SHEET_NAME}`);

  const srcLastRow = srcSheet.getLastRow();
  const srcLastCol = srcSheet.getLastColumn();
  if (srcLastRow <= SOURCE_HEADER_ROW) return;

  // ヘッダー（3行目）
  const headers = srcSheet
    .getRange(SOURCE_HEADER_ROW, 1, 1, srcLastCol)
    .getValues()[0]
    .map(h => String(h ?? "").trim());

  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });

  // 改行入りヘッダー対策（発注\n主体）
  if (idx["発注\n主体"] !== undefined && idx["発注主体"] === undefined) {
    idx["発注主体"] = idx["発注\n主体"];
  }

  // 管理番号列
  const keyColIdx = SOURCE_KEY_COL - 1;

  // ===== データ抽出 =====
  const out = [];
  out.push(TARGET_COLS);

  const data = srcSheet
    .getRange(SOURCE_HEADER_ROW + 1, 1, srcLastRow - SOURCE_HEADER_ROW, srcLastCol)
    .getValues();

  let hit = 0;

  for (const row of data) {
    const cell = String(row[keyColIdx] ?? '').trim(); // SOURCE_KEY_COLのセル

    if (!cell) continue;

    // まずセル内のキーを全部抜く（例: "123456-01 123456-02" → ["123456-01","123456-02"]）
    const keysInCell = extractKeys_(cell);
    if (!keysInCell.length) continue;

    // フォーム側に存在するキーだけに絞る
    const matchedKeys = keysInCell.filter(k => keySet.has(k));

    // もし extractKeys_ が拾えない表記がある場合に備えて保険（包含チェック）
    // 例: "案件:123456-01/123456-02" みたいな変則でも拾う
    if (!matchedKeys.length) {
      // keySetはSetなので直ループでOK（サイズが大きいなら後で最適化可能）
      let found = null;
      for (const k of keySet) {
        if (cellHasKey_(cell, k)) { found = k; break; }
      }
      if (!found) continue;
      matchedKeys.push(found);
    }

    // ★ 行を「ヒットしたキー数分」出す（= 123456-01 と 123456-02 が同セルなら2行出す）
    for (const mKey of matchedKeys) {
      out.push(
        TARGET_COLS.map(col => {
          if (col === '管理番号') return mKey; // ★ここが重要：ヒットしたキーを管理番号として出す
          const c = idx[col];
          return c === undefined ? '' : row[c];
        })
      );
      hit++;
    }
  }

  // ===== 書き込み =====
  let destSh = getSH_('計上管理_転写');
  if (!destSh) destSh = ss.insertSheet(DEST_SHEET_NAME);
  destSh.clear();

  destSh
    .getRange(1, 1, out.length, TARGET_COLS.length)
    .setValues(out);

  destSh.setFrozenRows(1);

  writeLog_('Auto_計上管理転写',`計上管理_写し 同期完了：${hit}件（フォーム案件 ${keySet.size}件）`);
}

// セル文字列から 123456-01 形式のキーを全部抜く（スペース/改行/カンマ区切りでもOK）
function extractKeys_(s) {
  s = String(s ?? '');
  const m = s.match(/\d{6}-\d{2}/g);
  return m ? Array.from(new Set(m)) : [];
}

// セル内に「そのキー」が含まれるか（誤ヒット防止：前後が英数でないこと）
function cellHasKey_(cellText, key) {
  const t = String(cellText ?? '');
  const k = String(key ?? '');
  if (!t || !k) return false;
  const re = new RegExp(`(^|[^0-9A-Za-z])${escapeRegExp_(k)}([^0-9A-Za-z]|$)`);
  return re.test(t);
}

function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}