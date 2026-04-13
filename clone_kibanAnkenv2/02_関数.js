/** =======================================================================
 * 
 *【Chat へメッセージを送る関数】
 *   
 * ======================================================================== */ 
function fc_postToChat(webhookUrl, message) {
  const payload = JSON.stringify({ text: message });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
  };

  UrlFetchApp.fetch(webhookUrl, options);
}

/** ======================================================================= 
 * 
 *【ヘッダー行から “列名配列” を取得する関数】
 * 引数：シート名
 *  
 * ======================================================================== */ 

function getHeaderList(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が存在しません`);

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  return header.map(v =>
    (v === null || v === undefined)
      ? ""
      : String(v).replace(/\r?\n/g, "").trim()
  ).filter(name => name !== ""); // 空欄は除外
}

/** ======================================================================= 
 * 
 *【列マップを作成する関数（共通）】
 * 引数：シート名
 *  
 * ======================================================================== */ 
function createColumnMap(sheetName) {
  const headerList = getHeaderList(sheetName);  // ←ここで 1行目から直接取得
  const headerMap  = getColumnIndexMap(sheetName);
  const colMap     = {};

  headerList.forEach(name => {
    if (headerMap[name] !== undefined) {
      colMap[name] = headerMap[name];
    }
  });

  return colMap;
}

/** ======================================================================= 
 * 
 *【日付が有効か確認する関数】
 * 引数：日付
 *  
 * ======================================================================== */ 
function fc_isValidDate(d) {
  return d instanceof Date && !isNaN(d);
}


/** ======================================================================= 
 * 
 *【前営業日を取得する関数】
 * 引数：タイムゾーン
 *  
 * ======================================================================== */ 
function fc_getPrevBusinessDayStr(tz) {
  const d = new Date();                         // 実行日時（ローカル）
  // まず前日に戻す
  d.setDate(d.getDate() - 1);
  // 土:6 / 日:0 はさらに戻す
  while (d.getDay() === 0 || d.getDay() === 6) {
    d.setDate(d.getDate() - 1);
  }
  return Utilities.formatDate(d, tz, 'yyyy/MM/dd');
}

/** ======================================================================= 
 * 
 *【前営業日を取得する関数】
 * 引数：タイムゾーン
 *  
 * ======================================================================== */ 
function getLastRowByColumn(sheet, colNumber) {
  const values = sheet.getRange(1, colNumber, sheet.getLastRow()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] != null) {
      return i + 1; // 行番号に戻す
    }
  }
  return 0;
}

/** ======================================================================= 
 *  
 *【月列の開始列を自動検出する関数】
 * 条件：ヘッダが Date 形式で、日付が「1日」のセル（例：2024/10/1）
 * ・見つかれば列番号（1始まり）を返す
 * ・見つからなければ null を返す
 *  
 * ======================================================================== */ 
function findMonthStartColumn(headerRow) {
  for (let i = 0; i < headerRow.length; i++) {
    const v = headerRow[i];

    // ★ シートの値が Date オブジェクトか？
    if (v instanceof Date) {

      // ★ 月の開始日（1日）なら「月ヘッダ」と判断
      if (v.getDate() === 1) {

        // 月列スタート（列番号1始まり）
        return i + 1;
      }
    }
  }

  return null;
}


/** ======================================================================= 
 * 
 *【開始・終了行をポップアップ表示する関数】
 *  
 * ======================================================================== */ 
function runWithPopup() {
  const ui = SpreadsheetApp.getUi();

  // --- 開始行 ---
  const startResp = ui.prompt(
    "開始行の入力",
    "開始行を入力してください：",
    ui.ButtonSet.OK_CANCEL
  );
  if (startResp.getSelectedButton() !== ui.Button.OK) return null;

  const startRow = Number(startResp.getResponseText());
  if (isNaN(startRow) || startRow <= 0) {
    ui.alert("開始行が不正です");
    return null;
  }

  // --- 終了行 ---
  const endResp = ui.prompt(
    "終了行の入力",
    "終了行を入力してください：",
    ui.ButtonSet.OK_CANCEL
  );
  if (endResp.getSelectedButton() !== ui.Button.OK) return null;

  const endRow = Number(endResp.getResponseText());
  if (isNaN(endRow) || endRow < startRow) {
    ui.alert("終了行が不正です");
    return null;
  }

  // --- start / end をセットで返す ---
  return { startRow, endRow };
}

/** =======================================================================
 * 
 * 【案件番号を入力して行番号を取得する関数】
 * 引数：シート
 * 
 * ======================================================================== */ 
function searchKanriTableByPopup(sheet) {

  // ▼ ① ポップアップを表示して値を取得
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "検索",
    "案件番号を入力してください（C列を検索します）",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("キャンセルされました");
    return;
  }

  const keyword = response.getResponseText().trim();
  if (keyword === "") {
    ui.alert("値が入力されていません");
    return;
  }

  // ▼ ② 案件管理表シートを取得
  if (!sheet) {
    ui.alert("対象のシートが見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("データがありません");
    return;
  }

  // ▼ ③ C列の値をまとめて取得（高速）
  const colCValues = sheet.getRange(1, 3, lastRow).getValues().flat();

  // ▼ ④ 入力値と一致する行を検索
  const rowIndex = colCValues.findIndex(val => String(val) === keyword);

  if (rowIndex === -1) {
    ui.alert(`C列に「${keyword}」は見つかりませんでした`);
    return -1;          // 見つからない場合
  }

  const hitRow = rowIndex + 1; // シート上の行番号に変換（0 → 1行目）
  return hitRow;
}

/**
 * =======================================================================
 * 
 * 【HYPERLINK / リンク付きセル から URL を取得する関数】
 *  
 * =======================================================================
 **/
function getUrlFromHyperlink(range) {
  const rt = range.getRichTextValue();
  if (rt && rt.getLinkUrl()) return rt.getLinkUrl();

  const formula = range.getFormula();
  if (formula && formula.includes('HYPERLINK')) {
    const m = formula.match(/HYPERLINK\(\s*"([^"]+)"/i);
    if (m) return m[1];
  }

  const v = range.getDisplayValue();
  if (typeof v === 'string') {
    let m = v.match(/HYPERLINK\(\s*"([^"]+)"/i);
    if (m) return m[1];
    if (v.match(/^https?:\/\//)) return v;
  }

  throw new Error('URL が取得できません: ' + range.getA1Notation());
}

/**
 * =======================================================================
 * 
 * 【開始行を取得する関数】
 * 
 * =======================================================================
 **/

function getStartRowByYesterday(sheet) {

  // 前日を文字列化（G列が文字列の場合の比較用）
  const ymd = Utilities.formatDate(yesterday, tz, "yyyy/MM/dd");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const gValues = sheet.getRange(1, 7, lastRow).getValues();  // G列(7列目)

  let startRow = null;

  for (let r = 1; r <= lastRow; r++) {
    const v = gValues[r - 1][0];

    if (!v) continue;

    // 日付型かどうかで比較を分岐
    if (v instanceof Date) {
      const vStr = Utilities.formatDate(v, tz, "yyyy/MM/dd");
      if (vStr === ymd) {
        startRow = r;
        break;
      }
    } else {
      // 文字列の場合
      if (String(v).trim() === ymd) {
        startRow = r;
        break;
      }
    }
  }

  return startRow; // 見つからなければ null
}


/** 
 * =======================================================================
 * 
 * HYPERLINKセルからURLのみを取得
 *  
 * ======================================================================== 
 **/ 
function getUrlFromHyperlink(cell) {
  const formula = cell.getFormula();
  if (formula && formula.includes("HYPERLINK")) {
    const match = formula.match(/HYPERLINK\("([^"]+)"/);
    return match ? match[1] : null;
  }
  const value = cell.getValue();
  return (typeof value === "string" && value.startsWith("http")) ? value : null;
}