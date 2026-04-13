function Auto_受注後フォルダ作成() {
  const NEW_OWNER_EMAIL = 'it-kiban@systena.co.jp'; 
  
  const sh = getSH_('案件情報');
  const HEADER_ROW = 1;

  const cols = getColsByHeaders_(sh, [
    '分類',
    '作業依頼日',
    '作業フォルダ',
    '見積',
  ], HEADER_ROW);

  const firstRow = HEADER_ROW + 1;
  const lastRow  = getLastRowByColumn_(sh, cols['作業依頼日']);
  if (lastRow < firstRow) { Logger.log("データ行がありません"); return; }

  const n = lastRow - firstRow + 1;

  const catVals        = sh.getRange(firstRow, cols['分類'], n, 1).getValues();
  const dateVals       = sh.getRange(firstRow, cols['作業依頼日'], n, 1).getValues();
  const folderRange    = sh.getRange(firstRow, cols['作業フォルダ'], n, 1);
  const mitsumoriRange = sh.getRange(firstRow, cols['見積'], n, 1);

  const SUB_FOLDERS_SV = [
    '01_見積り',
    '02_スケジュール',
    '03_ヒアリングシート',
    '04_作業用資料',
    '08_作業報告書',
    '09_納品物'
  ];

  const SUB_FOLDERS_CL = [
    '01_見積り',
    '02_スケジュール',
    '03_チェックシート',
    '04_作業用資料',
    '08_作業報告書',
    '09_納品物'
  ];

  const mapFolders = { SV: SUB_FOLDERS_SV, PC: SUB_FOLDERS_CL };

  // 対象行（昨日）
  const targetIdx = [];
  for (let i = 0; i < n; i++) {
    if (isYesterday_(dateVals[i][0], tz)) targetIdx.push(i);
  }
  if (targetIdx.length === 0) { Logger.log("◎昨日の行がありませんでした。"); return; }

  let ok = 0, skip = 0, ng = 0;

  for (const i of targetIdx) {
    const r = firstRow + i;

    try {
      const catVal = String(catVals[i][0] ?? '').trim().toUpperCase();
      const list = mapFolders[catVal];
      if (!list) {
        writeLog_('Auto_受注後フォルダ作成', `Row ${r}: 分類が不正（${catVal}）のためスキップ`);
        Logger.log(`Row ${r}: 分類が不正（${catVal}）のためスキップ`);
        skip++;
        continue;
      }

      // 作業フォルダ（base）
      const baseUrl = getUrlFromCell_(folderRange.getCell(i + 1, 1));
      if (!baseUrl) { skip++; continue; }

      const baseFolderId = extractDriveIdFromUrl_(baseUrl);
      if (!baseFolderId) { skip++; continue; }

      const baseFolder = DriveApp.getFolderById(baseFolderId);
      const workFolder = getOrCreateFolderWithOwnership_(baseFolder, '受注後_作業用資料', NEW_OWNER_EMAIL);

      // 下位フォルダ作成
      let mitsumoriFolder = null;
      list.forEach(name => {
        const f = getOrCreateFolderWithOwnership_(workFolder, name, NEW_OWNER_EMAIL);
        if (name === '01_見積り') mitsumoriFolder = f;
      });

      // 見積（Excel）コピー→スプレッドシート変換→見積セルURL差し替え
      if (mitsumoriFolder) {
        const cell = mitsumoriRange.getCell(i + 1, 1);
        const mitsumoriUrl = getUrlFromCell_(cell);
        if (mitsumoriUrl) {
          const fileId = extractDriveIdFromUrl_(mitsumoriUrl);
          if (fileId) {
            const copied = copyExcelAsGoogleSheetToFolder_(fileId, mitsumoriFolder.getId());
            const newUrl = DriveApp.getFileById(copied.id).getUrl(); // ←確実
            cell.setValue(newUrl);

            if (NEW_OWNER_EMAIL) transferOwnership_(copied.id, NEW_OWNER_EMAIL);
          }
        }
      }

      ok++;
    } catch (e) {
      writeLog_('Auto_受注後フォルダ作成', `Row ${r}: エラー - ${e}`);
      Logger.log(`Row ${r}: エラー - ${e}`);
      ng++;
    }
  }

  writeLog_('Auto_受注後フォルダ作成', `ok=${ok}, skip=${skip}, ng=${ng}`);
  Logger.log(`◎完了 ok=${ok}, skip=${skip}, ng=${ng}`);
}

/** フォルダ取得 or 作成（新規の時だけ譲渡） */
function getOrCreateFolderWithOwnership_(parent, name, newOwnerEmail) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();

  const newFolder = parent.createFolder(name);

  // 新規作成時だけ譲渡（メールが設定されている場合）
  if (newOwnerEmail) {
    try {
      transferOwnership_(newFolder.getId(), newOwnerEmail);
      Logger.log(`所有権を ${newOwnerEmail} に譲渡: ${name}`);
    } catch (e) {
      Logger.log(`所有権譲渡に失敗 (${name}): ${e}`);
    }
  }
  return newFolder;
}

/**
 * ExcelをGoogleスプレッドシートに変換しつつコピーし、
 * 必ず dstFolderId 配下に入るよう「親の付け替え」まで行う
 *
 * ※ Advanced Drive Service (Drive) 必須
 */
function copyExcelAsGoogleSheetToFolder_(srcFileId, dstFolderId) {
  // 元ファイル
  const src = Drive.Files.get(srcFileId, { supportsAllDrives: true });

  // ① まずコピー（convert）
  const resource = {
    title: src.title,
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents: [{ id: dstFolderId }]
  };

  const copied = Drive.Files.copy(resource, srcFileId, {
    convert: true,
    supportsAllDrives: true
  });

  // ② 親が効かないケースがあるので、必ず「親付け替え」をする
  const after = Drive.Files.get(copied.id, { supportsAllDrives: true });

  const curParents = (after.parents || []).map(p => p.id).filter(Boolean);
  const removeParents = curParents.filter(id => id !== dstFolderId).join(',');

  // ★ここがポイント：patch(resource, fileId, optionalArgs) を使う（mediaData不要）
  Drive.Files.update(
    {},
    copied.id,
    null, // ← mediaData を null にする
    {
      addParents: dstFolderId,
      removeParents: removeParents || undefined,
      supportsAllDrives: true
    }
  );

  const fixed = Drive.Files.get(copied.id, { supportsAllDrives: true });

  return {
    id: fixed.id,
    url: DriveApp.getFileById(fixed.id).getUrl(), // ★確実
    name: fixed.title,
    parents: (fixed.parents || []).map(p => p.id)
  };
}

/** Drive API(v3)で所有者譲渡 */
function transferOwnership_(fileOrFolderId, email) {
  if (!fileOrFolderId || !email) return;

  const perm = Drive.Permissions.create(
    { type: 'user', role: 'writer', emailAddress: email },
    fileOrFolderId,
    { sendNotificationEmail: false }
  );

  Drive.Permissions.update(
    { role: 'owner' },
    fileOrFolderId,
    perm.id,
    { transferOwnership: true }
  );
}

/** ファイルを指定フォルダにコピー */
function copyFileToFolder_(srcFileId, dstFolder) {
  const src = DriveApp.getFileById(srcFileId);
  return src.makeCopy(src.getName(), dstFolder);
}
