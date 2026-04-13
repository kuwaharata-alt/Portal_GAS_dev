/** 
 * =======================================================================
 * 
 * 【スケジュールチャート登録】
 * 
 * ======================================================================== 
 **/ 

function SC_Update_ScheduleChart() {
  try {
    if (!sheetSC) throw new Error('sheetSC が定義されていません');
    if (!sheetPD) throw new Error('sheetPD が定義されていません');

    const sheet = sheetSC;
    const CHART_START_COL = 7; // G列

    // ===== G1 → 13ヶ月分の年月を 1 行目に出力 =====
    let firstMonth = sheet.getRange(1, CHART_START_COL).getValue();

    // 文字列だった場合 → 日付に変換
    if (!(firstMonth instanceof Date)) {
      const s = String(firstMonth).trim(); // "2025/12"
      const parts = s.split('/');
      if (parts.length === 2) {
        const y = Number(parts[0]);
        const m = Number(parts[1]) - 1; // 月は 0 開始
        firstMonth = new Date(y, m, 1);
      }
    }

    if (!(firstMonth instanceof Date) || isNaN(firstMonth.getTime())) {
      throw new Error('G1 を日付に変換できませんでした： ' + firstMonth);
    }


    for (let i = 0; i < 13; i++) {
      const d = new Date(firstMonth.getFullYear(), firstMonth.getMonth() + i, 1);
      sheet.getRange(1, CHART_START_COL + i).setValue(d);
    }

    const today = new Date();

    // ===== ヘッダから必要な列番号を取得 =====
    const headerMap = getColumnIndexMap(sheet.getSheetName());

    const colStart    = headerMap['作業依頼日'];
    const colEnd      = headerMap['完了予定'];
    const colKantoku  = headerMap['監督者'];
    const colMain     = headerMap['メイン(管理者)'];
    const colSup1     = headerMap['サポート①'];
    const colSup2     = headerMap['サポート②'];
    const colKey      = headerMap['案件番号'] || 2;  // なければ B列などに変更

    if (!colStart || !colEnd || !colKantoku || !colMain || !colSup1 || !colSup2) {
      throw new Error('ヘッダー名（作業依頼日/完了予定/監督者/メイン(管理者)/サポート①/サポート②）を確認してください');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('データ行がありません');
      return;
    }

    const lastCol = sheet.getLastColumn();
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // ===== 月ごとの開始/終了日を事前に配列で作っておく =====
    const monthInfoList = [];
    for (let i = 0; i < 13; i++) {
      const c = CHART_START_COL + i;
      const monthCell = sheet.getRange(1, c).getValue();
      if (!(monthCell instanceof Date)) {
        monthInfoList.push(null);
        continue;
      }
      const ms = new Date(monthCell.getFullYear(), monthCell.getMonth(), 1);
      const me = new Date(monthCell.getFullYear(), monthCell.getMonth() + 1, 0);
      const isCurrent =
        ms.getFullYear() === today.getFullYear() &&
        ms.getMonth() === today.getMonth();

      monthInfoList.push({
        col: c,
        start: ms,
        end: me,
        isCurrentMonth: isCurrent,
      });
    }

    // ===== PDmenu から「フルネーム → 略称」マップ作成 =====
    const nameMap = getPdNameMap();
    Logger.log('PDmenu nameMap size = ' + Object.keys(nameMap).length);

    // ===== 各行のチャートを作成 =====
    for (let r = 2; r <= lastRow; r++) {
      const rowIndex = r - 2;
      const row = values[rowIndex];

      const keyVal = row[colKey - 1];
      if (!keyVal) continue;  // 案件番号などが空ならスキップ

      const startRaw = row[colStart - 1];
      const endRaw   = row[colEnd   - 1];

      if (!(startRaw instanceof Date)) {
        // 開始日が日付でない行はスキップ
        continue;
      }

      const startDate = normalizeDate(startRaw);
      const endDate = normalizeDate(
        endRaw instanceof Date ? endRaw : startRaw
      );

      // ===== 役割別に略称を取得 =====
      const kantoku = sanitizeMember(row[colKantoku - 1], nameMap);
      const main    = sanitizeMember(row[colMain     - 1], nameMap);
      const sup1    = sanitizeMember(row[colSup1     - 1], nameMap);
      const sup2    = sanitizeMember(row[colSup2     - 1], nameMap);

      // 表示対象メンバー配列を作成（表示順：監督 → メイン → サポ1 → サポ2）
      const members = [];
      if (kantoku) members.push({ text: kantoku, role: 'kantoku' });
      if (main)    members.push({ text: main,    role: 'main'    });
      if (sup1)    members.push({ text: sup1,    role: 'support' });
      if (sup2)    members.push({ text: sup2,    role: 'support' });

      // １人もいなければチャートは書かないが、一応セルはクリアしておく
      for (let i = 0; i < 13; i++) {
        const info = monthInfoList[i];
        if (!info) continue;
        sheet.getRange(r, info.col).setValue('');
      }
      if (members.length === 0) continue;

      // ===== 月ごとに、期間がかぶるセルへ書き込み =====
      for (let i = 0; i < 13; i++) {
        const info = monthInfoList[i];
        if (!info) continue;

        if (!isOverlap(startDate, endDate, info.start, info.end)) {
          continue;
        }

        const rich = buildMemberRichText(members, info.isCurrentMonth);
        sheet.getRange(r, info.col).setRichTextValue(rich);
      }
    }

  } catch (e) {
    Logger.log('SC_ScheduleChart error: ' + e.message);
    Logger.log(e.stack);
  }
}

/** 
 * =======================================================================
 * 
 * 【フルネームから略称へマップ化】
 * 
 * ======================================================================== 
 **/ 
 function getPdNameMap() {
  const sheet = sheetPD;
  const last = sheet.getLastRow();
  const map = {};

  if (last < 2) return map;

  // T列(20)〜W列(23) を取得（2行目から）
  const rng = sheet.getRange(2, 20, last - 1, 4).getValues();
  rng.forEach(row => {
    const full = String(row[0] || '').trim(); // T列
    const short = String(row[3] || '').trim(); // W列
    if (full && short) {
      map[full] = short;
    }
  });
  return map;
}

/** 
 * =======================================================================
 * 
 * 【メンバー名の正規化 + 略称化】
 *　※ ー / - / 未アサイン / 空欄 は null 扱い
 * 
 * ======================================================================== 
 **/ 
function sanitizeMember(v, nameMap) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s) return null;
  if (s === '-' || s === 'ー' || s === '未アサイン') return null;

  // 略称が定義されていれば略称、なければフルネームそのまま
  return nameMap[s] || s;
}

/** 
 * =======================================================================
 * 
 * 【メンバー配列からリッチテキストを作成】
 * 
 * ======================================================================== 
 **/ 
function buildMemberRichText(members, isCurrentMonth) {
  const text = members.map(m => m.text).join('/');
  const builder = SpreadsheetApp.newRichTextValue().setText(text);

  let pos = 0;
  members.forEach((m, idx) => {
    const len = m.text.length;
    const color = getRoleColor(m.role, isCurrentMonth);

    const style = SpreadsheetApp.newTextStyle()
      .setForegroundColor(color)
      .build();

    // 範囲が文字列長を超えないようにガード
    const start = pos;
    const end = Math.min(pos + len, text.length);
    builder.setTextStyle(start, end, style);

    pos += len;
    if (idx < members.length - 1) {
      // "/" 1文字ぶん
      pos += 1;
    }
  });

  return builder.build();
}

/** 
 * =======================================================================
 * 
 * 【当月 / 翌月以降で色分け】
 * 
 * ======================================================================== 
 **/ 
function getRoleColor(role, isCurrentMonth) {
  if (isCurrentMonth) {
    switch (role) {
      case 'kantoku': return '#000080'; // ネイビー
      case 'main':    return '#FF0000'; // 赤
      case 'support': return '#800080'; // 紫
      default:        return '#000000';
    }
  } else {
    switch (role) {
      case 'kantoku': return '#7F7FBF'; // 薄ネイビー
      case 'main':    return '#FF9999'; // 薄赤
      case 'support': return '#CC99FF'; // 薄紫
      default:        return '#808080';
    }
  }
}

/** 
 * =======================================================================
 * 
 * 【期間がかぶっているか判定】
 * 
 * ======================================================================== 
 **/ 
function isOverlap(start1, end1, start2, end2) {
  const s1 = start1.getTime();
  const e1 = end1.getTime();
  const s2 = start2.getTime();
  const e2 = end2.getTime();
  return s1 <= e2 && s2 <= e1;
}

/** 
 * =======================================================================
 * 
 * 【日付の時間部分を 00:00 に正規化】
 * 
 * ======================================================================== 
 **/ 
function normalizeDate(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
