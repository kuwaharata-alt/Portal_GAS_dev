/**
 * =======================================================================
 * 
 *【ヘッダーを見て列番号を辞書化する関数】
 * ・引数：シート名
 * ※改行は無視して処理する
 *  
 * ========================================================================
 **/

function getColumnIndexMap(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が存在しません`);

  const lastCol = sheet.getLastColumn();

  // 1行目を 1次元配列で取得
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const map = {};

  header.forEach((value, index) => {
    // ★ null / undefined を "" に、その他はすべて String に変換
    let name = (value === null || value === undefined)
      ? ""
      : String(value);

    // 前後スペース削除 & 改行削除
    name = name.replace(/\r?\n/g, "").trim();

    // 空欄ヘッダはスキップ
    if (!name) return;

    map[name] = index + 1;  // 1-based の列番号
  });

  return map;
}

/**
 * =======================================================================
 * 
 *【列変数】
 *  
 * =======================================================================
 **/

// ----- SV案件管理表 ------------------------------------------------------

const SVCOL = getColumnIndexMap(vSV);

const SVCOL_UPDATE_DATE        = SVCOL["更新日"];
const SVCOL_STATUS             = SVCOL["ステータス"];
const SVCOL_JUCHU_MITSUMORI    = SVCOL["受注見積(自動)"];
const SVCOL_MITSUMORI          = SVCOL["見積(自動)"];
const SVCOL_CUSTOMER           = SVCOL["顧客名(自動)"];
const SVCOL_FOLDER             = SVCOL["フォルダ(自動)"];
const SVCOL_SAGYOU_IRAI        = SVCOL["作業依頼日(自動)"];
const SVCOL_KENSHU_YOTEI       = SVCOL["検収予定(自動)"];
const SVCOL_KANRYOU_YOTEI      = SVCOL["完了予定"];
const SVCOL_SAGYOU_KEITAI      = SVCOL["作業形態"];
const SVCOL_KANTOKU            = SVCOL["監督者"];
const SVCOL_MAIN               = SVCOL["メイン(管理者)"];
const SVCOL_SUPPORT1           = SVCOL["サポート1"];
const SVCOL_SUPPORT2           = SVCOL["サポート2"];
const SVCOL_GENCHI_PLACE       = SVCOL["現地作業場所"];
const SVCOL_BIKOU              = SVCOL["備考"];

const SVCOL_SN_STATUS          = SVCOL["【社内作業】ステータス"];
const SVCOL_SN_TASK            = SVCOL["【社内作業】タスク"];

const SVCOL_GENCHI_STATUS      = SVCOL["【現地作業】ステータス"];
const SVCOL_GENCHI_TASK        = SVCOL["【現地作業】タスク"];

const SVCOL_KICKOFF            = SVCOL["【案件管理】キックオフ"];
const SVCOL_HEARING            = SVCOL["【案件管理】ヒアリングシート"];
const SVCOL_TIMESCHEDULE       = SVCOL["【案件管理】タイムスケジュール"];
const SVCOL_TIMESCHEDULELINK   = SVCOL["【案件管理】タイムスケジュールリンク"];
const SVCOL_WBS                = SVCOL["【案件管理】WBS"];
const SVCOL_DELIVERABLES       = SVCOL["【案件管理】納品物"];
const SVCOL_REPORT             = SVCOL["【案件管理】作業報告書"];
const SVCOL_ITKARTE            = SVCOL["【案件管理】ITカルテ"];

const SVCOL_担当営業              = SVCOL["担当営業"];
const SVCOL_PRE_SALES          = SVCOL["担当プリセールス"];


// ----- CL案件管理表 ------------------------------------------------------

const CLCOL = getColumnIndexMap(vCL);

const CLCOL_UPDATE_DATE        = CLCOL["更新日"];
const CLCOL_STATUS             = CLCOL["ステータス"];
const CLCOL_JUCHU_MITSUMORI    = CLCOL["受注見積(自動)"];
const CLCOL_MITSUMORI          = CLCOL["見積(自動)"];
const CLCOL_CUSTOMER           = CLCOL["顧客名(自動)"];
const CLCOL_FOLDER             = CLCOL["フォルダ(自動)"];
const CLCOL_SAGYOU_IRAI        = CLCOL["作業依頼日(自動)"];
const CLCOL_KENSHU_MONTH       = CLCOL["検収予定月(自動)"];
const CLCOL_KANRYOU_YOTEI      = CLCOL["完了予定"];
const CLCOL_SAGYOU_KEITAI      = CLCOL["作業形態"];
const CLCOL_TANTO_PRE          = CLCOL["担当プリ"];
const CLCOL_MAIN1              = CLCOL["メイン1(管理者)"];
const CLCOL_MAIN2              = CLCOL["メイン2"];
const CLCOL_SUPPORT            = CLCOL["サポート※入力注意※"];

const CLCOL_HANADA             = CLCOL["花田涼"];
const CLCOL_OOSAWA             = CLCOL["大澤義和"];
const CLCOL_TANIGUCHI          = CLCOL["谷口勝之"];
const CLCOL_UCHIDA             = CLCOL["内田進一"];
const CLCOL_HAGIWARA           = CLCOL["萩原永士"];
const CLCOL_SATO_HIROSHI       = CLCOL["佐藤浩"];
const CLCOL_SATO_TETSURO       = CLCOL["佐藤哲郎"];
const CLCOL_KUWAHARA           = CLCOL["桑原拓也"];
const CLCOL_ITO_KAITO          = CLCOL["伊藤開斗"];
const CLCOL_HAGIYA             = CLCOL["萩谷駿介"];
const CLCOL_ASHIZAWA           = CLCOL["芦澤隆希"];
const CLCOL_SUGINO             = CLCOL["杉野勇人"];
const CLCOL_KITAGAWA           = CLCOL["北川十吾"];
const CLCOL_ENDO               = CLCOL["遠藤岳"];
const CLCOL_NAKANO             = CLCOL["中野丈章"];
const CLCOL_ISHIWARI           = CLCOL["石割智貴"];
const CLCOL_OTAMA              = CLCOL["大玉拳嗣"];
const CLCOL_IZUMI              = CLCOL["泉龍之介"];
const CLCOL_SUZUKI             = CLCOL["鈴木琉沙"];
const CLCOL_YOSHIDA            = CLCOL["吉田翔"];
const CLCOL_KONISHI            = CLCOL["小西惇斗"];
const CLCOL_KONNO              = CLCOL["今野雅之"];
const CLCOL_SANPEI             = CLCOL["三瓶翔真"];
const CLCOL_KUMAGAI            = CLCOL["熊谷将紀"];
const CLCOL_KUSANAGI           = CLCOL["草薙龍"];
const CLCOL_TAZAKI             = CLCOL["田崎玲央"];
const CLCOL_HAMASATO           = CLCOL["浜里航也"];
const CLCOL_AOYANAGI           = CLCOL["青栁凜汰郎"];
const CLCOL_BP                 = CLCOL["BP"];
const CLCOL_BIKOU              = CLCOL["備考"];

const CLCOL_SN_STATUS          = CLCOL["【社内作業】ステータス"];
const CLCOL_SN_PLACE           = CLCOL["【社内作業】場所"];
const CLCOL_SN_BIKOU           = CLCOL["【社内作業】備考"];

const CLCOL_GENCHI_PROGRESS    = CLCOL["【現地作業】進捗"];
const CLCOL_GENCHI_BIKOU       = CLCOL["【現地作業】備考"];

const CLCOL_KICKOFF            = CLCOL["【案件管理】キックオフ"];
const CLCOL_WBS                = CLCOL["【案件管理】WBS"];
const CLCOL_CHECKSHEET         = CLCOL["【案件管理】チェックシート"];
const CLCOL_DOCUMENT           = CLCOL["【案件管理】ドキュメント"];
const CLCOL_REPORT             = CLCOL["【案件管理】作業報告書"];
const CLCOL_担当営業             = CLCOL["担当営業"];


// ----- スケジュール ------------------------------------------------------

const SCCOL = getColumnIndexMap(vSC);

const SCCOL_ステータス           = SCCOL["ステータス"];
const SCCOL_案件番号             = SCCOL["案件番号"];
const SCCOL_顧客名_自動        = SCCOL["顧客名(自動)"];
const SCCOL_作業依頼日         = SCCOL["作業依頼日"];
const SCCOL_完了予定           = SCCOL["完了予定"];
const SCCOL_作業ボリューム     = SCCOL["作業ボリューム"];
const SCCOL_ボリューム重み     = SCCOL["ボリューム重み"];
const SCCOL_監督者             = SCCOL["監督者"];
const SCCOL_メイン_管理者      = SCCOL["メイン(管理者)"];
const SCCOL_サポート1          = SCCOL["サポート①"];
const SCCOL_サポート2          = SCCOL["サポート②"];

const SCCOL_作業工数_月単位   = SCCOL["作業工数(月単位)"];
const SCCOL_作業ポイント_監督 = SCCOL["作業ポイント監督"];   
const SCCOL_作業ポイント_メイン = SCCOL["作業ポイントメイン"];
const SCCOL_作業ポイント_サポ1 = SCCOL["作業ポイントサポート①"];
const SCCOL_作業ポイント_サポ2 = SCCOL["作業ポイントサポート②"];

const SCCOL_HANADA             = SCCOL["花田涼"];
const SCCOL_OOSAWA             = SCCOL["大澤義和"];
const SCCOL_TANIGUCHI          = SCCOL["谷口勝之"];
const SCCOL_UCHIDA             = SCCOL["内田進一"];
const SCCOL_HAGIWARA           = SCCOL["萩原永士"];
const SCCOL_SATO_HIROSHI       = SCCOL["佐藤浩"];
const SCCOL_SATO_TETSURO       = SCCOL["佐藤哲郎"];
const SCCOL_KUWAHARA           = SCCOL["桑原拓也"];
const SCCOL_ITO_KAITO          = SCCOL["伊藤開斗"];
const SCCOL_HAGIYA             = SCCOL["萩谷駿介"];
const SCCOL_ASHIZAWA           = SCCOL["芦澤隆希"];
const SCCOL_SUGINO             = SCCOL["杉野勇人"];
const SCCOL_KITAGAWA           = SCCOL["北川十吾"];
const SCCOL_ENDO               = SCCOL["遠藤岳"];
const SCCOL_NAKANO             = SCCOL["中野丈章"];
const SCCOL_ISHIWARI           = SCCOL["石割智貴"];
const SCCOL_OTAMA              = SCCOL["大玉拳嗣"];
const SCCOL_IZUMI              = SCCOL["泉龍之介"];
const SCCOL_SUZUKI             = SCCOL["鈴木琉沙"];
const SCCOL_YOSHIDA            = SCCOL["吉田翔"];
const SCCOL_KONISHI            = SCCOL["小西惇斗"];
const SCCOL_KONNO              = SCCOL["今野雅之"];
const SCCOL_SANPEI             = SCCOL["三瓶翔真"];
const SCCOL_KUMAGAI            = SCCOL["熊谷将紀"];
const SCCOL_KUSANAGI           = SCCOL["草薙龍"];
const SCCOL_TAZAKI             = SCCOL["田崎玲央"];
const SCCOL_HAMASATO           = SCCOL["浜里航也"];
const SCCOL_AOYANAGI           = SCCOL["青栁凜汰郎"];


// ----- 開始行 -----------------------------------------------------------
const startRowSV = getStartRowByYesterday(sheetSV);
const startRowCL = getStartRowByYesterday(sheetCL);
const startRowSC = 300


// ----- 終了行 -----------------------------------------------------------
const lastRowSV = getLastRowByColumn(sheetSV, SVCOL_JUCHU_MITSUMORI);
const lastRowCL = getLastRowByColumn(sheetCL, CLCOL_JUCHU_MITSUMORI);
const lastRowSC = getLastRowByColumn(sheetSC, SCCOL_案件番号);

