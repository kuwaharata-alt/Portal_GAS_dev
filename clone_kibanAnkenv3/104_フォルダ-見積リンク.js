/** 
 * 【リンク挿入】
 * ・見積りのリンクを入れる
 * ・作業フォルダへのリンクを入れる
 * 
 **/

function Auto_リンク挿入() {
  const sh = getSH_('案件情報');
  const cols = getColsByHeaders_(sh, ['見積番号', '作業依頼日', '見積', '作業フォルダ', 'open/close'], 1);

  const firstRow = 2;
  const lastRow = getLastRowByColumn_(sh, cols['作業依頼日']);
  if (lastRow < firstRow) {
    writeLog_('Auto_リンク挿入', 'データ行がありません');
    return;
  }

  const mitsumorinumCol = cols['見積番号'];
  const reqdayCol       = cols['作業依頼日'];
  const mitsumoriCol    = cols['見積'];
  const workfolCol      = cols['作業フォルダ'];
  const closeCol        = cols['open/close'];

  const n = lastRow - firstRow + 1;

  // --- まとめて取得 ---
  const numVals   = sh.getRange(firstRow, mitsumorinumCol, n, 1).getValues();
  const dateVals  = sh.getRange(firstRow, reqdayCol,       n, 1).getValues();

  // 既存値を保持（昨日以外を上書きしない）
  const folderLink = sh.getRange(firstRow, workfolCol, n, 1).getValues();
  const fileLink   = sh.getRange(firstRow, mitsumoriCol, n, 1).getValues();
  const closeVals  = sh.getRange(firstRow, closeCol, n, 1).getValues();

  // --- 昨日の対象行 index を抽出＆ユニークキー収集 ---
  const targetIdx = [];
  const uniqKey6 = new Set();
  const uniqNo   = new Set();

  for (let i = 0; i < n; i++) {
    const d = dateVals[i][0];
    if (!isYesterday_(d, tz)) continue;

    const no = String(numVals[i][0] ?? '').trim();
    if (!no) continue;

    targetIdx.push(i);

    const key6 = no.substring(0, 6);
    if (key6) uniqKey6.add(key6);

    uniqNo.add(no);
  }

  if (targetIdx.length === 0) {
    writeLog_('Auto_リンク挿入', '昨日の日付が入力されている行は見つかりませんでした。');
    return;
  }

  // --- Drive検索（ユニークキー単位で1回ずつ） ---
  const esc = s => String(s).replace(/'/g, "\\'");
  const folderUrlByKey6 = new Map();
  const fileUrlByNo     = new Map();

  // フォルダ：先頭6文字
  for (const key6 of uniqKey6) {
    let url = "";
    const it = DriveApp.searchFolders(`title contains '${esc(key6)}'`);
    if (it.hasNext()) url = it.next().getUrl();
    folderUrlByKey6.set(key6, url);
  }

  // ファイル：見積番号
  for (const no of uniqNo) {
    let url = "";
    const itF = DriveApp.searchFiles(`title contains '${esc(no)}'`);
    if (itF.hasNext()) url = itF.next().getUrl();
    fileUrlByNo.set(no, url);
  }

  // --- 対象行にだけ反映（配列更新） ---
  let hit = 0;
  for (const i of targetIdx) {
    const no = String(numVals[i][0] ?? '').trim();
    const key6 = no.substring(0, 6);

    const folderUrl = folderUrlByKey6.get(key6) || "";
    const fileUrl   = fileUrlByNo.get(no) || "";

    if (folderUrl) folderLink[i][0] = folderUrl;
    if (fileUrl)   fileLink[i][0]   = fileUrl;

    closeVals[i][0] = "open"; // ← 昨日対象だけオープン

    hit++;
  }

  // --- 一括反映 ---
  sh.getRange(firstRow, workfolCol, n, 1).setValues(folderLink);
  sh.getRange(firstRow, mitsumoriCol, n, 1).setValues(fileLink);
  sh.getRange(firstRow, closeCol,     n, 1).setValues(closeVals);

  writeLog_('Auto_リンク挿入', `リンク追加完了（対象: ${hit}行）`);
  Logger.log(`◎リンク追加完了（対象: ${hit}行）`);
}


/** 
 * 【リンク挿入(手動実行用)】
 * ・見積りのリンクを入れる
 * ・作業フォルダへのリンクを入れる
 * 
 **/
/**
 * 高速版：昨日の行だけを対象に、Drive検索をユニークキー単位に削減
 */
function Manual_リンク挿入() {

  // ** ここを変更する *************************//

  const firstRow = 2;
  const lastRow = 475;

  // ****************************************//

  const sh = getSH_('案件情報');
  const cols = getColsByHeaders_(sh, ['見積番号', '作業依頼日', '見積', '作業フォルダ', 'open/close'], 1);

  if (lastRow < firstRow) {
    writeLog_('Auto_リンク挿入', 'データ行がありません');
    return;
  }

  const mitsumorinumCol = cols['見積番号'];
  const mitsumoriCol    = cols['見積'];
  const workfolCol      = cols['作業フォルダ'];
  const closeCol        = cols['open/close'];

  const n = lastRow - firstRow + 1;

  // --- まとめて取得 ---
  const numVals   = sh.getRange(firstRow, mitsumorinumCol, n, 1).getValues();

  // 既存値を保持（昨日以外を上書きしない）
  const folderLink = sh.getRange(firstRow, workfolCol, n, 1).getValues();
  const fileLink   = sh.getRange(firstRow, mitsumoriCol, n, 1).getValues();
  const closeVals  = sh.getRange(firstRow, closeCol, n, 1).getValues();

  // --- 昨日の対象行 index を抽出＆ユニークキー収集 ---
  const targetIdx = [];
  const uniqKey6 = new Set();
  const uniqNo   = new Set();

  for (let i = 0; i < n; i++) {

    const no = String(numVals[i][0] ?? '').trim();
    if (!no) continue;

    targetIdx.push(i);

    const key6 = no.substring(0, 6);
    if (key6) uniqKey6.add(key6);

    uniqNo.add(no);
  }

  // --- Drive検索（ユニークキー単位で1回ずつ） ---
  const esc = s => String(s).replace(/'/g, "\\'");
  const folderUrlByKey6 = new Map();
  const fileUrlByNo     = new Map();

  // フォルダ：先頭6文字
  for (const key6 of uniqKey6) {
    let url = "";
    const it = DriveApp.searchFolders(`title contains '${esc(key6)}'`);
    if (it.hasNext()) url = it.next().getUrl();
    folderUrlByKey6.set(key6, url);
  }

  // ファイル：見積番号
  for (const no of uniqNo) {
    let url = "";
    const itF = DriveApp.searchFiles(`title contains '${esc(no)}'`);
    if (itF.hasNext()) url = itF.next().getUrl();
    fileUrlByNo.set(no, url);
  }

  // --- 対象行にだけ反映（配列更新） ---
  let hit = 0;
  for (const i of targetIdx) {
    const no = String(numVals[i][0] ?? '').trim();
    const key6 = no.substring(0, 6);

    const folderUrl = folderUrlByKey6.get(key6) || "";
    const fileUrl   = fileUrlByNo.get(no) || "";

    if (folderUrl) folderLink[i][0] = folderUrl;
    if (fileUrl)   fileLink[i][0]   = fileUrl;

    closeVals[i][0] = "open"; // ← 昨日対象だけオープン

    hit++;
  }

  // --- 一括反映 ---
  sh.getRange(firstRow, workfolCol, n, 1).setValues(folderLink);
  sh.getRange(firstRow, mitsumoriCol, n, 1).setValues(fileLink);
  sh.getRange(firstRow, closeCol,     n, 1).setValues(closeVals);

  writeLog_('Auto_リンク挿入', `リンク追加完了（対象: ${hit}行）`);
  Logger.log(`◎リンク追加完了（対象: ${hit}行）`);
}
