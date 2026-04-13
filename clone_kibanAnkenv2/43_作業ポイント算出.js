/** 
 * =======================================================================
 * 
 * 【役割ごとの「作業ポイント」を自動計算】
 * 
 * ======================================================================== 
 **/ 
function SC_CalcWorkPoints() {

  const startRowSC = 2; // データ開始行

  // 基本重み
  const W_SUP   = 0.1;
  const W_MAIN  = 0.4;
  const W_SUP1  = 0.25;
  const W_SUP2  = 0.25;

  if (!sheetSC) {
    throw new Error('シートが見つかりません: SV案件管理表');
  }

  // ▼ 最終行取得（必要に応じて「作業工数」列基準にしてもOK）
  const lastRowSC = sheetSC.getLastRow();
  if (lastRowSC < startRowSC) return;

  const numRows = lastRowSC - startRowSC + 1;

  // 全列まとめて取得
  const allValues = sheetSC
    .getRange(startRowSC, 1, numRows, sheetSC.getLastColumn())
    .getValues();

  // 出力配列（AO〜ARの4列分）
  const outValues = Array.from({ length: numRows }, () => ["", "", "", ""]);

  for (let i = 0; i < numRows; i++) {
    const row = allValues[i];

    const sup   = row[SCCOL_監督者          - 1];
    const main  = row[SCCOL_メイン_管理者   - 1];
    const sup1  = row[SCCOL_サポート1       - 1];
    const sup2  = row[SCCOL_サポート2       - 1];
    const work  = row[SCCOL_作業工数_月単位 - 1];

    // 作業工数が空 or 数値でない行はスキップ
    if (work === "" || work === null || isNaN(work)) continue;

    const hasSup  = !isEmptyRole_(sup);
    const hasMain = !isEmptyRole_(main);
    const hasSup1 = !isEmptyRole_(sup1);
    const hasSup2 = !isEmptyRole_(sup2);

    let sumW = 0;
    if (hasSup)  sumW += W_SUP;
    if (hasMain) sumW += W_MAIN;
    if (hasSup1) sumW += W_SUP1;
    if (hasSup2) sumW += W_SUP2;

    if (sumW === 0) continue; // 誰もいない

    outValues[i][0] = hasSup  ? work * (W_SUP  / sumW) : ""; // AO：監督
    outValues[i][1] = hasMain ? work * (W_MAIN / sumW) : ""; // AP：メイン
    outValues[i][2] = hasSup1 ? work * (W_SUP1 / sumW) : ""; // AQ：サポ①
    outValues[i][3] = hasSup2 ? work * (W_SUP2 / sumW) : ""; // AR：サポ②
  }

  // 作業工数列の1つ右（ANの右=AO）から4列に出力
  const outStartCol = SCCOL_作業工数_月単位 + 1;
  sheetSC
    .getRange(startRowSC, outStartCol, numRows, 4)
    .setValues(outValues);
}
/** =======================================================================
 * 
 * 
 * ======================================================================== 
 **/ 
function isEmptyRole_(value) {
  if (value === null) return true;
  const s = String(value).trim();
  return (s === "" || s === "ー" || s === "未アサイン");
}
