/** 
 * =======================================================================
 * 
 * 【更新日の登録】
 * ・SV案件管理表、CL案件管理表のA列に更新日を入力する
 * 
 * ========================================================================
 **/ 

function Updatedate(e) { 
  if (!e || !e.range) return;

  const sh   = e.range.getSheet();
  const name = sh.getName();
  const row0 = e.range.getRow();
  const nrow = e.range.getNumRows();
  const col0 = e.range.getColumn();
  const ncol = e.range.getNumColumns();

  if (name !== vSV && name !== vCL) return;

  // ヘッダー行は対象外
  if (row0 < 2) return;

  // 対象列レンジ
  const ranges = {
    [vSV]: [2, 30],
    [vCL]: [2, 54],
  };
  const [cMin, cMax] = ranges[name];

  // 編集範囲が対象列と重なっていなければ終了
  const editedLeft  = col0;
  const editedRight = col0 + ncol - 1;
  const overlaps = !(editedRight < cMin || editedLeft > cMax);
  if (!overlaps) return;

  // 更新日をA列に記録
  const now = new Date();
  sh.getRange(row0, 1, nrow, 1).setValue(now);
}