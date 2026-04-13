/** =======================================================================
 * 
 * 【ヒートマップ更新】
 * 
 * 
 * ======================================================================== 
 **/ 
function SC_Update_Heatmap() {

  const ss   = SpreadsheetApp.openById(sheetId);
  const shSC = sheetSC || ss.getSheetByName(vSC);     // スケジュール
  const shHM = ss.getSheetByName('ヒートマップ');     // ヒートマップ

  if (!shSC) throw new Error('スケジュールシートなし');
  if (!shHM) throw new Error('ヒートマップシートなし');

  // ===== ヒートマップ側の設定 =====
  const startColHM = 11;   // O列
  const endColHM   = 16;   // X列
  const monthCount = endColHM - startColHM + 1;

  const lastRowHM = shHM.getLastRow();
  if (lastRowHM < 2) return;    // 担当者がいない

  const memberCount = lastRowHM - 1;

  // 担当者名（B列）
  const hmMembers = shHM.getRange(2, 1, memberCount, 1).getValues()
    .map(r => String(r[0] || '').trim());

  // O1〜X1 の月
  const hmMonthValues = shHM.getRange(1, startColHM, 1, monthCount).getValues()[0];

  // ===== スケジュール側：基準月 & 列情報 =====
  const baseMonthColSC = 7; // G列
  const baseMonthVal   = shSC.getRange(1, baseMonthColSC).getValue();
  const baseMonthDate  = toMonthDate_(baseMonthVal);

  if (!baseMonthDate) {
    throw new Error('スケジュール!G1 の月が判別できません');
  }

  const lastColSC = shSC.getLastColumn();

  // ヒートマップO〜X列それぞれに対応する「スケジュールの列番号」を作る
  const scMonthCols = hmMonthValues.map(val => {
    const d = toMonthDate_(val);
    if (!d) return null;

    const diff = monthDiff_(baseMonthDate, d); // G1からの月差
    const col  = baseMonthColSC + diff;

    if (diff < 0 || col > lastColSC) {
      return null; // スケジュール側に存在しない
    }
    return col;
  });

  const validMonthCols = scMonthCols.filter(c => c);
  if (validMonthCols.length === 0) {
    throw new Error('O1〜X1 の月に対応するスケジュールの月列が見つかりません（G1基準のEDATEになっているか確認してください）');
  }

  // ===== スケジュール側：担当者列 & ボリューム列 =====

  // ヘッダー（上5行）から担当者名列とボリューム列を探す
  const headerRows  = 1;
  const headers2D   = shSC.getRange(1, 1, headerRows, lastColSC).getValues();

  const scMemberMap = {};   // 担当者名 → col
  let   volumeCol   = null; // 作業ボリューム列

  for (let r = 0; r < headerRows; r++) {
    const row = headers2D[r];
    row.forEach((val, i) => {
      const col  = i + 1;
      const name = String(val || '').trim();
      if (!name) return;

      // 担当者名
      scMemberMap[name] = col;

      // 作業ボリューム（AA列）
      if (!volumeCol && /ボリューム/.test(name)) {
        volumeCol = col;
      }
    });
  }

  if (!volumeCol) {
    throw new Error('「作業ボリューム」と書かれた列がスケジュール見つかりません');
  }

  // ヒートマップB列の担当者名 → スケジュールの列番号
  const memberCols = hmMembers.map(name => scMemberMap[name] || null);
  const validMemberCols = memberCols.filter(c => c);

  if (validMemberCols.length === 0) {
    throw new Error('ヒートマップB列の担当者名と一致する列がスケジュールにありません');
  }

  const memberMinCol = Math.min(...validMemberCols);
  const memberMaxCol = Math.max(...validMemberCols);

  // ===== スケジュール本体データ =====
  const lastRowSC = shSC.getLastRow();
  const dataRows  = lastRowSC - 1;
  if (dataRows <= 0) return;

  // 月チャート（G列〜右側の必要範囲）
  const monthMinCol = Math.min(...validMonthCols);
  const monthMaxCol = Math.max(...validMonthCols);

  const monthMatrix = shSC
    .getRange(2, monthMinCol, dataRows, monthMaxCol - monthMinCol + 1)
    .getValues();

  // 担当者ポイント（AI〜）
  const memberMatrix = shSC
    .getRange(2, memberMinCol, dataRows, memberMaxCol - memberMinCol + 1)
    .getValues();

  // 作業ボリューム（AA）
  const volumeValues = shSC
    .getRange(2, volumeCol, dataRows, 1)
    .getValues()
    .map(r => Number(r[0]) || 0);

  // ===== 出力バッファ（担当者 × 月） =====
  const result = Array.from({ length: memberCount }, () =>
    Array(monthCount).fill(0)
  );

  // ===== 集計ロジック =====
  for (let r = 0; r < dataRows; r++) {

    const volume = volumeValues[r];      // 5 / 3 / 1 / 0
    if (!volume) continue;              // 0 や空欄の案件は無視

    for (let j = 0; j < monthCount; j++) {

      const colSC = scMonthCols[j];
      if (!colSC) continue;            // スケジュール側に列がない月

      const relMonth = colSC - monthMinCol;
      const monthCell = monthMatrix[r][relMonth];

      // その案件がその月に稼働していなければスキップ
      if (monthCell === '' || monthCell === null) continue;

      for (let i = 0; i < memberCount; i++) {

        const mColSC = memberCols[i];
        if (!mColSC) continue;

        const relMember = mColSC - memberMinCol;
        const v = memberMatrix[r][relMember];   // 作業ポイント（1か月分）

        if (v === '' || v === null) continue;

        const point = Number(v);
        if (!isNaN(point)) {
          result[i][j] += point * volume;      // 作業ポイント × 作業ボリューム
        }
      }
    }
  }

  // ===== ヒートマップへ出力（O2〜X） =====
  const outRange = shHM.getRange(2, startColHM, memberCount, monthCount);
  outRange.clearContent();
  outRange.setValues(result);
}

/** =======================================================================
 * 
 * 【任意の値から「その月の1日」の Date を作る】
 * 例：2025/12, 2025-12-01, 2025年12月 → Date(2025,11,1)
 * 
 * ======================================================================== 
 **/ 
function toMonthDate_(val) {
  if (val == null || val === "") return null;

  // Date 型
  if (Object.prototype.toString.call(val) === "[object Date]") {
    return new Date(val.getFullYear(), val.getMonth(), 1);
  }

  const s = String(val).trim();
  if (!s) return null;

  // 年・月をざっくり抜き出す（2025/12, 2025-12, 2025年12月, 2025/12/1 等）
  const m = s.match(/(\d{4})\D+(\d{1,2})/);
  if (!m) return null;

  const y = Number(m[1]);
  const mm = Number(m[2]) - 1; // JSの月は0始まり

  if (isNaN(y) || isNaN(mm)) return null;

  return new Date(y, mm, 1);
}

/** =======================================================================
 * 
 * 【2つの月の差（base → target が何か月離れているか）を返す】
 * 
 * ======================================================================== 
 **/ 
function monthDiff_(baseDate, targetDate) {
  return (targetDate.getFullYear() - baseDate.getFullYear()) * 12 +
         (targetDate.getMonth() - baseDate.getMonth());
}
