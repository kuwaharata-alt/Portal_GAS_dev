/** 
 * =======================================================================
 * 
 * 【作業用フォルダの作成】(SV版)
 * ・D列のリンク先フォルダ配下に「受注後_作業用資料」フォルダを探し、4種類のフォルダを作成
 * ・新規に作成したフォルダは「it-kiban@systena.co.jp」にオーナー権限を譲渡する
 * ・01_スケジュール配下タイムスケジュールのテンプレートをコピーする
 * 
 * ======================================================================== 
 **/

// 作成するフォルダ名リスト
const SUB_FOLDERS = [
  '01_スケジュール',
  '02_ヒアリングシート',
  '08_作業報告書',
  '09_納品物'
];

/** 
 * =======================================================================
 * 
 * 【フォルダ作成】
 * ・自動実行用
 * 
 * ======================================================================== 
 **/

function SV_AcreateWorkFolders(){
  const startRow = startRowSV
  const lastRow = lastRowSV
  
  if (startRow === null ) return;
  AcreateWorkFoldersSV(startRow,lastRow)
}

/** 
 * =======================================================================
 * 
 * 【フォルダ作成】
 * ・手動実行用
 * ・案件番号を入力する
 * 
 * ======================================================================== 
 **/

function SV_AcreateWorkFoldersManual(){
  const row = searchKanriTableByPopup(sheetSV);
  const startRow = row
  const lastRow = row
  AcreateWorkFoldersSV(startRow,lastRow)
}

/** 
 * =======================================================================
 * 
 * 【フォルダ作成】
 * ・フォルダ作成する関数
 * ・引数：開始行、終了行
 * 
 * ======================================================================== 
 **/

function AcreateWorkFoldersSV(startRow,lastRow) {
  Logger.log("〇フォルダ作成(SV):S:"+startRow+"L:"+lastRow);

  for (let r = startRow; r <= lastRow; r++) {   

    const linkCell = sheetSV.getRange(r, SVCOL_FOLDER); // D列：HYPERLINK

    const url = getUrlFromHyperlink(linkCell);
    if (!url) continue;

    const folderId = extractDriveIdFromUrl(url);
    if (!folderId) continue;

    try {
      const baseFolder = DriveApp.getFolderById(folderId);

      // 「受注後_作業用資料」フォルダを取得 or 作成(＋新規ならオーナー譲渡)
      const workFolder = getOrCreateFolderWithOwnership(baseFolder, '受注後_作業用資料');

      // 下位フォルダ 4つ作成（新規ならオーナー譲渡）
      // ついでに 01_スケジュールフォルダを控えておく
      let scheduleFolder = null;
      SUB_FOLDERS.forEach(name => {
        const f = getOrCreateFolderWithOwnership(workFolder, name);
        if (name === '01_スケジュール') {
          scheduleFolder = f;
        }
      });

      // 01_スケジュールがあればテンプレコピー＆B2/G2セット
      if (scheduleFolder) {
        const url = copyScheduleTemplateAndFill_SV(scheduleFolder, r);

        // ▼ AE列に HYPERLINK を設定
        const rng = sheetSV.getRange(r, SVCOL_TIMESCHEDULELINK);  // AE列
        const rich = SpreadsheetApp.newRichTextValue()
        .setText('リンク')
        .setLinkUrl(url)
        .build();
  
        rng.setRichTextValue(rich);

        Logger.log(`Row ${r}: フォルダ作成＆オーナー処理完了`);
      }
    } catch (e) {
      Logger.log(`Row ${r}: エラー - ${e}`);
    } 
  }
}


/** 
 * =======================================================================
 * 
 * 【スケジュール自動作成】
 * ・引数：保存先フォルダ、行数
 * 
 * ======================================================================== 
 **/

function copyScheduleTemplateAndFill_SV(workFolder, row) {

  // ▼ まず E列の値を取得（顧客名など）
  const nameVal = sheetSV.getRange(row, 5).getDisplayValue(); // E列
  const baseName = nameVal && nameVal !== '' ? nameVal : 'Temp';

  // ▼ ファイル名：【E列の値様】タイムスケジュール
  const newFileName = `【${baseName}様】タイムスケジュール`;

  // ① テンプレファイルをコピー（ファイル名を上記に変更）
  const templateFile = DriveApp.getFileById(TEMPLATE_SCHEDULE_FILE_ID);
  const copiedFile   = templateFile.makeCopy(newFileName, workFolder);

  // ② オーナー譲渡（ファイル）
  try {
    transferOwnership(copiedFile.getId(), NEW_OWNER_EMAIL);
    Logger.log(`タイムスケジュールの所有権を ${NEW_OWNER_EMAIL} に譲渡しました`);
  } catch (e) {
    Logger.log(`タイムスケジュール所有権の譲渡に失敗しました: ${e}`);
  }

  // ③ コピーしたスプレッドシートを開く
  const ss    = SpreadsheetApp.openById(copiedFile.getId());
  const sheet =
    ss.getSheetByName(TEMPLATE_SCHEDULE_SHEET_NAME) || ss.getSheets()[0];

  // ④ C列の値（案件番号など）を取得（G2用）
  const noVal = sheetSV.getRange(row, 3).getDisplayValue(); // C列
  const g2Text = noVal ? String(noVal).slice(0, 6) : '';

  // ▼ B2 もファイル名と同じ書式にする
  //    「【E列の値様】タイムスケジュール」
  const b2Text = `【${baseName}様】タイムスケジュール`;

  sheet.getRange('B2').setValue(b2Text);
  sheet.getRange('G2').setValue(g2Text);

  // ⑤ コピーしたファイルの URL を返す
  return `https://docs.google.com/spreadsheets/d/${copiedFile.getId()}`;
}

/** 
 * =======================================================================
 * 
 * 【新規フォルダ作成&オーナー権限譲渡】
 * ・親フォルダ配下に指定名のフォルダが存在すればそれを返す
 * ・なければ作成して「NEW_OWNER_EMAIL」にオーナー権限を譲渡
 * 
 * ======================================================================== 
 **/

function getOrCreateFolderWithOwnership(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    // 既存の場合はオーナー変更しない（必要ならここで変更も可能）
    return folders.next();
  }

  // 新規作成したフォルダ
  const newFolder = parent.createFolder(name);

  // オーナー譲渡
  try {
    transferOwnership(newFolder.getId(), NEW_OWNER_EMAIL);
    Logger.log(`所有権を ${NEW_OWNER_EMAIL} に譲渡しました: ${name}`);
  } catch (e) {
    Logger.log(`所有権譲渡に失敗しました (${name}): ${e}`);
  }

  return newFolder;
}

/** 
 * =======================================================================
 * 
 * 【所有権を譲渡する】
 * ※Drive APIの有効化が必要
 * 
 * ======================================================================== 
 **/

function transferOwnership(folderId, email) {

  // まず writer 権限で共有を追加
  const permissionResource = {
    type: 'user',
    role: 'writer',
    emailAddress: email
  };

  const perm = Drive.Permissions.create(
    permissionResource,
    folderId,
    { sendNotificationEmail: false } // 通知メールが要るなら true
  );

  // 追加した権限を owner に昇格させつつ所有権を譲渡
  Drive.Permissions.update(
    {
      role: 'owner'
    },
    folderId,
    perm.id,
    {
      transferOwnership: true
    }
  );
}

/** 
 * =======================================================================
 * 
 * 【HYPERLINKセルから URL を取り出す】
 * 
 * ======================================================================== 
 **/
function getUrlFromHyperlink(range) {
  const formula = range.getFormula();
  if (!formula) return null;

  const match = formula.match(/HYPERLINK\("(.+?)"/i);
  return match ? match[1] : null;
}

/** 
 * =======================================================================
 * 
 * 【HYPERLINK関数内URLからドライブIDを抜き出す】
 * 
 * ======================================================================== 
 **/

function extractDriveIdFromUrl(url) {
  if (!url) return null;

  // open?id=ID パターン
  const openMatch = url.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (openMatch) return openMatch[1];

  // /folders/ID や /d/ID/ のパターン
  const idMatch = url.match(/[-\w]{25,}/);
  return idMatch ? idMatch[0] : null;
}
