/** 
 * =======================================================================
 * 
 * 使用する変数を設定
 * 
 * ========================================================================
 **/ 

// ----- シート変数 ---------------------------------------------------------
const vSV = 'SV案件管理表'
const vCL = 'CL案件管理表'
const vSC = 'スケジュール'
var sheetId = '1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II';
var sheetReq = SpreadsheetApp.openById(sheetId).getSheetByName('作業依頼');
var sheetSV = SpreadsheetApp.openById(sheetId).getSheetByName(vSV);
var sheetCL = SpreadsheetApp.openById(sheetId).getSheetByName(vCL);
var sheetSC = SpreadsheetApp.openById(sheetId).getSheetByName(vSC);
var sheetPD = SpreadsheetApp.openById(sheetId).getSheetByName('PDMenu');
var sheetHM = SpreadsheetApp.openById(sheetId).getSheetByName('ヒートマップ');
var sheetBSOL = SpreadsheetApp.openById(sheetId).getSheetByName('BSOL情報');


// ----- 日付変数 ------------------------------------------------------------
const tz = 'Asia/Tokyo';

const today = new Date(); // 現在の日付を取得
const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

const yesterday = new Date(today);
yesterday.setDate(today.getDate()-1); // 日付を1日前に設定

const tommorow = new Date(today);
tommorow.setDate(today.getDate()+1); // 日付を1日後に設定


// ----- WebhokURL ---------------------------------------------------------
// *44基盤チームbot
 const webhookUrl_G = "https://chat.googleapis.com/v1/spaces/AAQA1Ou2jsg/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=5HzVdEkAYglMDoYRV-DKBNPRWwDC3oR8dWaRNTYihFg"; 

// *案件確認
 const webhookUrl_C = "https://chat.googleapis.com/v1/spaces/AAAAZkbJFSA/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=kSZ3RZi6gb2Z2ZGUSIE2rUwj4iDEhwKvwKHporZTW9U";

// *案件完了確認（案件確認スペース）
 const webhookUrl_F = "https://chat.googleapis.com/v1/spaces/AAAAZkbJFSA/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=q1anpCPk9ghUhANyIDwriP-ONX9Km6UNqC7mjXJsm1Y"

// テスト用
 const webhookUrl_T = "https://chat.googleapis.com/v1/spaces/AAQAuHpWqBA/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=uw8WMfiJjh1BP8W-p_pWaenRQj01VW0Fzx1Dh_xgYwQ"; 

// 確認用
 const webhookUrl_A = "https://chat.googleapis.com/v1/spaces/AAQApzvWPvQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=4QR3t3UsN5SFkF9T6qyQpOykOrWo-R3l38XJzas1TAg";


// ----- テンプレートファイル情報 ---------------------------------------------
// *作業報告書テンプレートのファイルID
const TEMPLATE_FILE_ID = '1aAM9kJXJHaqAaQxQCRIE9ubGiH77lxcjT1Ncpchti5k';

// *作業報告書シート名（テンプレ側）
const REPORT_SHEET_NAME = '作業報告書';


// *スケジュールテンプレのファイルID
const TEMPLATE_SCHEDULE_FILE_ID = '1n_vk_Fyz4xLPMmIh9hKIlAJLFH5D6hWrwMSfXt75Ti0';
// *テンプレ内のシート名（違っていればここを書き換え）
const TEMPLATE_SCHEDULE_SHEET_NAME = 'TS-Day1';


// ----- その他 -------------------------------------------------------------
// *オーナーを譲渡したい相手
const NEW_OWNER_EMAIL = 'it-kiban@systena.co.jp';
 