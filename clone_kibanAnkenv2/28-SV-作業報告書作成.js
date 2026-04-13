/** 
 * =======================================================================
 *  
 * 【作業報告書 自動作成】SV版
 * ・アクティブ行の情報から作業報告書を1つ作成
 * ・テンプレートをコピーして、名前・本文を自動セット
 * 
 * ======================================================================== 
 **/

/** 
 * =======================================================================
 * 
 * 【作業報告書 自動作成】
 * ・自動実行用
 * 
 * ======================================================================== 
 **/
function SV_createWorkReport() {
  if (startRowSV === null ) return;
  createWorkReportFromRow(sheetSV, startRowSV, lastRowSV);
}

/** 
 * =======================================================================
 * 
 * 【作業報告書 自動作成】
 * ・手動実行用
 * 
 * ======================================================================== 
 **/
function SV_createWorkReportManual() {
  const row = searchKanriTableByPopup(sheetSV);
  createWorkReportFromRow(sheetSV, row, row);
}

/** 
 * =======================================================================
 *
 *  【作業報告書 自動作成】
 * 
 * ======================================================================== 
 **/
function createWorkReportFromRow(sheet, startRow, lastRow) {
  
  for (let row = startRow; row <= lastRow; row++) {

    const valC  = sheet.getRange(row, SVCOL_JUCHU_MITSUMORI).getValue(); 
    const valE  = sheet.getRange(row, SVCOL_CUSTOMER).getValue();        
    const valBD = sheet.getRange(row, SVCOL_担当営業).getValue();        
    const linkCellMitsumori = sheet.getRange(row, SVCOL_MITSUMORI);      
    const linkCellFolder    = sheet.getRange(row, SVCOL_FOLDER);         

    if (!valE) throw new Error(`E列（顧客名）が空です: 行 ${row}`);
    if (!valC) throw new Error(`C列（案件番号）が空です: 行 ${row}`);

    // --- 見積ExcelのファイルID ---
    const fileId = extractFileIdFromUrl(getUrlFromHyperlink(linkCellMitsumori));
    if (!fileId) throw new Error(`見積リンクからファイルID取得不可: 行 ${row}`);

    // --- 作業フォルダのフォルダID ---
    const folderId = extractDriveIdFromUrl(getUrlFromHyperlink(linkCellFolder));
    if (!folderId) throw new Error(`作業フォルダリンクからフォルダID取得不可: 行 ${row}`);

    const baseFolder = DriveApp.getFolderById(folderId);

    // ▼「受注後_作業用資料」フォルダ（必要なら作成＋権限移譲）
    const afterOrderFolder =
      getOrCreateChildFolderWithOwner(baseFolder, '受注後_作業用資料');

    // ▼「08_作業報告書」フォルダ（必要なら作成＋権限移譲）
    const reportFolder =
      getOrCreateChildFolderWithOwner(afterOrderFolder, '08_作業報告書');

    // --- 報告内容の読み取り ---
    const srcSS = openExcelAsSpreadsheet(fileId);
    const found = findSourceSheetForReport(srcSS);

    const reportLines = found
      ? getReportLinesFromSourceSheet(found.sheet, found.startRow)
      : [];

    // --- テンプレコピー ---
    const custName = toZenkakuKana(valE.toString());
    const newName = `【${custName}様】作業報告書※作成中`;

    const template = DriveApp.getFileById(TEMPLATE_FILE_ID);

    // ★ コピー先を「08_作業報告書」フォルダへ
    const newFile = template.makeCopy(newName, reportFolder);

    // ★ コピーしたファイルのオーナーを移譲
    transferFileOwner(newFile);

    const reportSS = SpreadsheetApp.openById(newFile.getId());
    const reportSheet = reportSS.getSheetByName(REPORT_SHEET_NAME) || reportSS.getSheets()[0];

    // --- 各種反映 ---
    const cPrefix = valC.toString().substring(0, 6);
    reportSheet.getRange('Y3').setValue(cPrefix);
    reportSheet.getRange('F5').setValue(custName);

    const f40Text = `ビジネスソリューション事業本部 SOL営業統括部 ${valBD}`;
    reportSheet.getRange('F42').setValue(f40Text);

    fillReportBody(reportSheet, reportLines);
  }
}


/** 
 * =======================================================================
 * 
 * 【Google DriveのURLから fileId を抽出】
 * 
 * ======================================================================== 
 **/
function extractFileIdFromUrl(url) {
  if (!url) return '';
  let m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];

  return '';
}

/** 
 * =======================================================================
 * 
 * 【Excelファイルを一時的にスプレッドシートに変換して開く】
 * 　※元がスプレッドシートならそのまま開く
 * 　※Drive APIを有効化しておく
 * 
 * ======================================================================== 
 **/
function openExcelAsSpreadsheet(fileId) {
  const file = Drive.Files.get(fileId);

  if (file.mimeType === MimeType.GOOGLE_SHEETS) {
    // すでにスプレッドシート
    return SpreadsheetApp.openById(fileId);
  }

  // Excel → スプレッドシートにコピー
  const copied = Drive.Files.copy(
    {
      title: file.title + '_forReport',
      mimeType: MimeType.GOOGLE_SHEETS
    },
    fileId
  );

  return SpreadsheetApp.openById(copied.id);
}
/** 
 * =======================================================================
 * 
 * 【報告書本文用の行配列を生成】
 * 
 * ======================================================================== 
 **/
function getReportLinesFromSourceSheet(srcSheet, startRow) {
  const lastRow = srcSheet.getLastRow();
  if (!startRow || startRow > lastRow) return [];

  // ★「完成基準」の1行下は項目名行なのでスキップ
  const dataStartRow = startRow + 1; // 完成基準の2行下から実データ
  if (dataStartRow > lastRow) return [];

  // C列全体を一度取っておく
  const colCValues = srcSheet
    .getRange(1, 3, lastRow, 1)
    .getValues()
    .map(r => String(r[0] || '').trim());

  // ★C列の「留意事項」行を探す（あればその1行前までを対象）
  let dataEndRow = lastRow;
  for (let r = dataStartRow; r <= lastRow; r++) {
    const v = colCValues[r - 1]; // インデックスは0開始
    if (v === '留意事項' || v.indexOf('留意事項') !== -1) {
      dataEndRow = r - 1; // 「留意事項」行の1つ前まで
      break;
    }
  }

  if (dataEndRow < dataStartRow) {
    Logger.log('getReportLinesFromSourceSheet: 留意事項の位置により対象行なし');
    return [];
  }

  const numRows = dataEndRow - dataStartRow + 1;
  if (numRows <= 0) return [];

  Logger.log(`getReportLinesFromSourceSheet: dataStartRow=${dataStartRow}, dataEndRow=${dataEndRow}, numRows=${numRows}`);

  // C〜K列（C=3, K=11 → 9列分）を一括取得
  const data = srcSheet.getRange(dataStartRow, 3, numRows, 9).getValues();

  const lines = [];

  data.forEach((row, idx) => {
    const item    = String(row[0] || '').trim(); // C列：項目
    const content = String(row[8] || '').trim(); // K列：作業内容

    if (!item && !content) return;

    const bulletLines = [];

    if (content) {
      content.split(/\r\n|\n/).forEach(line => {
        let t = String(line).trim();
        if (!t) return;
        if (t.startsWith('※')) return; // ※※印などの行は除外
        bulletLines.push(t);
      });
    }

    if (!item && bulletLines.length === 0) return;

    // 〇項目
    if (item) lines.push('〇' + item);

    // ・作業範囲（1行ずつ）
    bulletLines.forEach(b => {
      lines.push('　・' + b); // 全角スペース＋中黒
    });
  });

  Logger.log(`getReportLinesFromSourceSheet: totalLines=${lines.length}`);
  return lines;
}

/** 
 * =======================================================================
 * 
 * 【作業報告書の本文出力】
 * 
 * ======================================================================== 
 **/
function fillReportBody(sheet, lines) {

  const ITEM_COL1    = 3;  // C列（項目）
  const CONTENT_COL1 = 4;  // D列（内容）
  const ITEM_COL2    = 18; // R列（項目）
  const CONTENT_COL2 = 19; // S列（内容）

  const START_ROW = 12;
  let capacityPerCol = 20;     // 1ブロックあたりの基本行数（11〜30）
  let maxCapacity    = capacityPerCol * 2; // 2ブロック合計

  Logger.log(`fillReportBody: raw lines.length=${lines.length}`);

  // ───────── 〇／・ごとに「グループ」化 ─────────
  // 例：["〇項目1", "　・内容1", "　・内容2", "〇項目2", ...]
  //   → [ ["〇項目1","　・内容1","　・内容2"], ["〇項目2", ...], ... ]
  const groups = [];
  let currentGroup = [];

  lines.forEach(rawLine => {
    let text = String(rawLine || '').trim();

    // 〇で始まる行は「項目」扱い
    if (text.startsWith('〇')) {
      // 〇項目の中に改行が入っている場合、改行以降を削除して1行化
      if (text.includes('\n')) {
        text = text.split(/\r?\n/)[0];
      }

      // 新しい項目開始
      if (currentGroup.length > 0) {
        groups.push(currentGroup);
      }
      currentGroup = [text];
      return;
    }

    // 〇以外（内容行など）
    if (currentGroup.length === 0) {
      // 念のため、先頭が〇でない例外ケース
      currentGroup = [text];
    } else {
      currentGroup.push(text);
    }
  });

  if (currentGroup.length > 0) groups.push(currentGroup);

  if (groups.length === 0) {
    Logger.log('fillReportBody: groups is empty, nothing to output');
    return;
  }

  // ───────── 全体の必要行数を概算 ─────────
  // グループを縦に並べるとして、
  // 各グループの行数の合計 ＋ グループ間の空行（グループ数-1）
  function calcTotalRowsForGroups(gs) {
    let sum = 0;
    gs.forEach(g => sum += g.length);
    if (gs.length > 1) sum += (gs.length - 1); // グループ間の空行
    return sum;
  }

  let totalRows = calcTotalRowsForGroups(groups);
  Logger.log(`fillReportBody: groups=${groups.length}, estimated totalRows=${totalRows}, maxCapacity=${maxCapacity}`);

  // ───────── 行数オーバー時の調整 ─────────
  if (totalRows > maxCapacity) {
    const extra = totalRows - maxCapacity;

    if (extra <= 5) {
      // 5行程度なら行を追加
      sheet.insertRowsAfter(START_ROW + capacityPerCol - 1, extra); // 30行目の下に挿入
      capacityPerCol += extra;
      maxCapacity = capacityPerCol * 2;
      Logger.log(`行追加: extra=${extra}, capacityPerCol=${capacityPerCol}, maxCapacity=${maxCapacity}`);
    } else {
      // それ以上溢れる場合は「〇項目」だけ残す（配下の・は削る）
      Logger.log(`多すぎるので「〇項目のみ」出力に切り替え`);
      for (let i = 0; i < groups.length; i++) {
        const g = groups[i];
        groups[i] = g[0] ? [g[0]] : [];
      }
      totalRows = calcTotalRowsForGroups(groups);
      Logger.log(`項目のみ totalRows=${totalRows}, maxCapacity=${maxCapacity}`);
    }
  }

  const lastRowToClear = START_ROW + capacityPerCol - 1;

  // ───────── 既存内容クリア ─────────
  for (let r = START_ROW; r <= lastRowToClear; r++) {
    sheet.getRange(r, ITEM_COL1).clearContent();
    sheet.getRange(r, CONTENT_COL1).clearContent();
    sheet.getRange(r, ITEM_COL2).clearContent();
    sheet.getRange(r, CONTENT_COL2).clearContent();
  }

  // ───────── 列ごとに「項目列・内容列」の配列を持つ ─────────
  // colIdx 0 → C/D, colIdx 1 → R/S
  const itemColLines    = [[], []];  // 項目用（C, R）
  const contentColLines = [[], []];  // 内容用（D, S）
  const usedRows        = [0, 0];    // 各ブロックで使用している行数
  let colIdx = 0;

  groups.forEach((grp, gIndex) => {
    if (colIdx > 1) return; // 3ブロック目は使わない

    const groupLength = grp.length;

    // このグループを現在列に置く際の必要行数
    // 既に何か書いてある列なら 1 行分の空行を追加
    let needRows = groupLength + (usedRows[colIdx] === 0 ? 0 : 1);

    // 入りきらなければ次ブロックへ
    if (usedRows[colIdx] + needRows > capacityPerCol) {
      colIdx++;
      if (colIdx > 1) {
        Logger.log(`colIdx>1 となったため、以降のグループは出力されません (gIndex=${gIndex})`);
        return;
      }
      // 新ブロック先頭グループは空行不要
      needRows = groupLength;
    }

    // 同じブロックで2個目以降のグループなら行間に空行を入れる
    if (usedRows[colIdx] > 0) {
      itemColLines[colIdx].push('');
      contentColLines[colIdx].push('');
      usedRows[colIdx]++;
    }

    // グループ内行を追加
    grp.forEach(line => {
      const text = String(line || '');

      if (text.startsWith('〇')) {
        // 項目行 → 項目列に出力、内容列は空白
        itemColLines[colIdx].push(text);
        contentColLines[colIdx].push('');
      } else if (text) {
        // 内容行（・など）→ 内容列に出力、項目列は空白
        itemColLines[colIdx].push('');
        contentColLines[colIdx].push(text);
      } else {
        // 念のための空行（通常はここには来ない想定）
        itemColLines[colIdx].push('');
        contentColLines[colIdx].push('');
      }

      usedRows[colIdx]++;
    });
  });

  Logger.log(
    `C/D rows=${itemColLines[0].length}, R/S rows=${itemColLines[1].length}, ` +
    `usedRows1=${usedRows[0]}, usedRows2=${usedRows[1]}`
  );

  // ───────── シートへ書き込み（1セルずつ）─────────
  // 左側 C/D
  itemColLines[0].forEach((val, i) => {
    sheet.getRange(START_ROW + i, ITEM_COL1).setValue(val);
  });
  contentColLines[0].forEach((val, i) => {
    sheet.getRange(START_ROW + i, CONTENT_COL1).setValue(val);
  });

  // 右側 R/S
  itemColLines[1].forEach((val, i) => {
    sheet.getRange(START_ROW + i, ITEM_COL2).setValue(val);
  });
  contentColLines[1].forEach((val, i) => {
    sheet.getRange(START_ROW + i, CONTENT_COL2).setValue(val);
  });

  SpreadsheetApp.flush();
}

/** 
 * =======================================================================
 * 
 *【文字列ユーティリティ】
 * ・半角カタカナを含む文字を全角に正規化
 * ※NFKC 正規化：半角ｶﾅ→全角カナ 他も全角化される点は注意
 * 
 * ======================================================================== 
 **/
function toZenkakuKana(str) {
  if (!str) return '';
  return str.normalize('NFKC');
}

/** 
 * =======================================================================
 * 
 *　【??】
 * 
 * ======================================================================== 
 **/
function findSourceSheetForReport(ss) {
  const sheets = ss.getSheets();

  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    const colCValues = sh
      .getRange(1, 3, lastRow, 1)
      .getValues()
      .map(r => String(r[0] || '').trim());

    for (let i = 0; i < lastRow; i++) {
      const v = colCValues[i];

      // 「完成基準」「完成基準○○」どちらもOK
      if (v === '完成基準' || v.indexOf('完成基準') !== -1) {
        const startRow = i + 2; // 「完成基準」の1行下から本文
        Logger.log(`完成基準 found: sheet=${sh.getName()}, row=${i + 1}, startRow=${startRow}`);
        return { sheet: sh, startRow: startRow };
      }
    }
  }

  return null;
}

/** 
 * =======================================================================
 * 
 * 【オーナー権限】
 * 
 * ======================================================================== 
 **/
function getOrCreateChildFolderWithOwner(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next(); // 既に存在

  // 新規作成
  const newFolder = parentFolder.createFolder(name);

  try {
    newFolder.setOwner(NEW_OWNER_EMAIL);
  } catch (err) {
    Logger.log(`★注意: フォルダ「${name}」のオーナー譲渡に失敗: ${err}`);
  }

  return newFolder;
}

/** 
 * =======================================================================
 * 
 * 【ファイルのオーナー譲渡】
 * 
 * ======================================================================== 
 **/
function transferFileOwner(file) {
  try {
    file.setOwner(NEW_OWNER_EMAIL);
  } catch (err) {
    Logger.log(`★注意: ファイル「${file.getName()}」のオーナー譲渡に失敗: ${err}`);
  }
}