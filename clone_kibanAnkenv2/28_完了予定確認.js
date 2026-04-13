/** =========================
 * トリガーから呼ぶ入口（3本）
 * ========================= */

// ① 毎日：期限超過
function trigger_SV_DueOverdue_Daily() {
  if (!isWeekday_()) return; // ★土日スキップ
  postSVCompletionDueAlerts_withMention_ByType_('overdue');
}

// ② 特定の日付（15、26）、ただし土日なら直前金曜：当月
function trigger_SV_DueThisMonth_OnSchedule() {
  if (!shouldSendThisMonthOnBusinessRule_()) return;
  postSVCompletionDueAlerts_withMention_ByType_('thisMonth');
}

// ③ 毎日：未入力
function trigger_SV_DueMissing_Daily() {
  if (!isWeekday_()) return; // ★土日スキップ
  postSVCompletionDueAlerts_withMention_ByType_('missing');
}


/** =========================
 * 種類別に送信する本体
 * type: 'overdue' | 'thisMonth' | 'missing'
 * ========================= */
function postSVCompletionDueAlerts_withMention_ByType_(type) {
  const TZ = typeof tz !== 'undefined' ? tz : 'Asia/Tokyo';
  const SHEET_ID = typeof sheetId !== 'undefined' ? sheetId : '';
  const SV_NAME = typeof vSV !== 'undefined' ? vSV : 'SV案件管理表';
  const WEBHOOK = typeof webhookUrl_F !== 'undefined' ? webhookUrl_F : '';

  if (!SHEET_ID) throw new Error('sheetId が未設定です');
  if (!WEBHOOK) throw new Error('webhookUrl が未設定です');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sv = ss.getSheetByName(SV_NAME);
  if (!sv) throw new Error(`シートが見つかりません: ${SV_NAME}`);

  const mentionMap = buildMentionMapFromBSOL_();

  const EXCLUDE_STATUS = new Set([
    "01.【案件管理】\nアサイン未",
    "81.【作業対象外】\nPSSEチーム対応",
    "82.【作業対象外】\nCL案件",
    "83.【作業対象外】\nキャンセル",
    "84.【作業対象外】\n保守",
    "85.【作業対象外】\nプリ/BP対応",
    "98. SE対応完了",
    "99. クローズ",
    "A1.【定常案件】\n対応中",
    "A2.【定常案件】\nクローズ",
    "B1. 検収後対応"
  ]);

  const startRow = 2;
  const lastRow = sv.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const values = sv.getRange(startRow, 1, numRows, 15).getValues();

  const now = new Date();
  const thisMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const nextMonthStart = new Date(now.getFullYear(), now.getMonth() + 1, 1);

  // typeに応じて1つだけ作る
  const groupedMap = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const status = row[1]; // B
    const e = row[4];      // E
    const due = row[8];    // I
    const o = row[14];     // O

    if (EXCLUDE_STATUS.has(String(status).trim())) continue;

    const rawMember = (o === null || o === undefined) ? '' : String(o).trim();
    const memberLabel = toMentionLabel_(rawMember, mentionMap);

    // --- 未入力 ---
    if (!(due instanceof Date)) {
      if (type !== 'missing') continue;
      const line = formatMissingLine_(e);
      if (line) pushMap_(groupedMap, memberLabel, `・${line}`);
      continue;
    }

    // --- 期限あり ---
    const isOverdue = due < thisMonthStart;
    const isThisMonth = due >= thisMonthStart && due < nextMonthStart;

    if (type === 'overdue' && isOverdue) {
      pushMap_(groupedMap, memberLabel, formatDueLine_(e, due, TZ));
    } else if (type === 'thisMonth' && isThisMonth) {
      pushMap_(groupedMap, memberLabel, formatDueLine_(e, due, TZ));
    }
  }

  // 案件が0件なら送らない
  if (groupedMap.size === 0) return;

  // ヘッダー文面
  const headerLines = getHeaderLinesByType_(type);

  // modeは表示形式（missingだけ箇条書きが違うなら分岐）
  const mode = (type === 'missing') ? 'missing' : 'due';

  const text = buildMessage_({ headerLines, groupedMap, mode });
  postToChat_(WEBHOOK, text);
}

function getHeaderLinesByType_(type) {
  if (type === 'overdue') {
    return [
      '【🔴作業完了予定 超過 通知🔴】',
      '下記案件の完了予定が超過しています。',
      '完了予定が延長になる場合は、スレッドへ理由を投稿した上で、案件管理表を更新してください。',
      'ステータス更新後、【済】アクションを行ってください。'
    ];
  }
  if (type === 'thisMonth') {
    return [
      '【🟡作業完了予定 確認 通知🟡】',
      '完了予定が【当月】です。',
      '当月完了予定の場合は、【OK】アクションを行ってください。',
      '完了予定が延長になる場合は、スレッドへ理由を投稿した上で、案件管理表を更新してください。'
    ];
  }
  // missing
  return [
    '【❗作業完了予定 登録依頼 通知❗】',
    '完了予定が未登録です。',
    '完了予定を登録してください。',
    'ステータス更新後、【済】アクションを行ってください。'
  ];
}


/** BSOL情報から (氏名/メール) → Chat USER_ID のMapを作る */
function buildMentionMapFromBSOL_() {
  // 既に共通変数で sheetBSOL がある前提（なければ openById）
  const sh = (typeof sheetBSOL !== 'undefined' && sheetBSOL)
    ? sheetBSOL
    : SpreadsheetApp.openById(sheetId).getSheetByName('BSOL情報');

  if (!sh) throw new Error('BSOL情報 シートが見つかりません');

  // あなたの運用：5〜31行にデータ
  const START_ROW = 5;
  const END_ROW = 31;
  const numRows = END_ROW - START_ROW + 1;

  // 想定：E=氏名 / F=メール / H=USER_ID
  const NAME_COL = 7;  // G
  const EMAIL_COL = 6; // F
  const ID_COL = 8;    // H

  const values = sh.getRange(START_ROW, 1, numRows, ID_COL).getValues();

  const map = new Map();

  values.forEach(r => {
    const name = (r[NAME_COL - 1] || '').toString().trim();
    const email = (r[EMAIL_COL - 1] || '').toString().trim();
    const id = (r[ID_COL - 1] || '').toString().trim();

    if (!id) return;

    if (email) map.set(email.toLowerCase(), id);
    if (name) map.set(name, id);
  });

  return map;
}

/** 担当者文字列をメンションラベルへ（できなければ元の文字列） */
function toMentionLabel_(rawMember, mentionMap) {
  const s = (rawMember || '').trim();
  if (!s) return '（担当者未設定）';

  // まずメールとして探す（O列がメール or 名前でもBSOLにメールがあるケースに備える）
  const emailLike = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
  if (emailLike) {
    const id = mentionMap.get(s.toLowerCase());
    return id ? `<users/${id}>` : s;
  }

  // 次に氏名キーで探す（O列が氏名の場合）
  const idByName = mentionMap.get(s);
  return idByName ? `<users/${idByName}>` : s;
}

/** Mapに追加（キーが無ければ配列作成） */
function pushMap_(map, key, val) {
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(val);
}

/** 期限あり：案件名 + 全角スペース2個 + 完了予定 */
function formatDueLine_(eVal, dueDate, tz) {
  const e = (eVal === null || eVal === undefined) ? '' : String(eVal).trim();
  const dueStr = Utilities.formatDate(dueDate, tz, 'yyyy/MM/dd');
  return `・完了予定：${dueStr}　　${e}`; // 全角スペース2個
}

/** 未入力用：案件名だけ（空なら出さない） */
function formatMissingLine_(eVal) {
  const e = (eVal === null || eVal === undefined) ? '' : String(eVal).trim();
  return e ? e : '';
}

/**
 * 本文生成
 * - 区切り線は「最初の担当者の上」と「一番下」の2か所だけ
 * - 2人目以降の担当者の上に空行を1行入れる
 */
function buildMessage_({ headerLines, groupedMap, mode }) {
  const sep = '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~';

  // キーはメンション文字列の可能性があるので、ソートは見た目優先でそのまま
  const keys = Array.from(groupedMap.keys());

  const blocks = [];
  blocks.push(...headerLines);
  blocks.push(sep);

  keys.forEach((k, idx) => {
    if (idx > 0) blocks.push(''); // 2人目以降の前に1行空ける

      blocks.push(`${k}`);
      (groupedMap.get(k) || []).forEach(line => blocks.push(`　　${line}`));
  });

  blocks.push(sep);
  return blocks.join('\n');
}

/** Google Chat Incoming Webhookへ投稿 */
function postToChat_(webhookUrl, text) {
  const res = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify({ text }),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Chat投稿に失敗: HTTP ${code} / ${res.getContentText()}`);
  }
}

/**
 * 10/18/25日に送信
 * ただし、その日が土日の場合「直前の金曜」に送信
 *
 * 例：
 * - 10日が土曜 → 8日(金)に送る
 * - 18日が日曜 → 17日(金)に送る
 */
function shouldSendThisMonthOnBusinessRule_() {
  const TZ = typeof tz !== 'undefined' ? tz : 'Asia/Tokyo';
  const now = new Date();
  const y = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  const m = Number(Utilities.formatDate(now, TZ, 'M'));
  const d = Number(Utilities.formatDate(now, TZ, 'd'));

  const targets = [15, 26];

  // 今日が対象日なら送る
  if (targets.includes(d)) return true;

  // 今日が金曜なら「明日/明後日が対象日(土日)」のケースを救済
  // 金曜=5（0:日,1:月,2:火,3:水,4:木,5:金,6:土）
  const dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1(月)〜7(日)
  const isFriday = (dow === 5);

  if (!isFriday) return false;

  const date = new Date(y, m - 1, d); // JS Date
  const sat = new Date(date); sat.setDate(date.getDate() + 1);
  const sun = new Date(date); sun.setDate(date.getDate() + 2);

  const satDay = sat.getDate();
  const sunDay = sun.getDate();

  // もし「土曜/日曜」が 10/18/25 なら、金曜の今日に送る
  if (targets.includes(satDay) || targets.includes(sunDay)) {
    // ただし月跨ぎ対策：翌日/翌々日が同じ月の時だけ
    const satMonth = sat.getMonth();
    const sunMonth = sun.getMonth();
    const thisMonth = date.getMonth();
    if (satMonth === thisMonth || sunMonth === thisMonth) return true;
  }

  return false;
}

function isWeekday_() {
  const TZ = typeof tz !== 'undefined' ? tz : 'Asia/Tokyo';
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1(月)〜7(日)
  return dow >= 1 && dow <= 5;
}


function testChatMentionByEmail() {
  // ★あなたのIncoming Webhook URLに置き換え

  const text = `<users/108572564371992059927>\nテスト`;

  const res = UrlFetchApp.fetch(webhookUrl_T, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify({ text }),
    muteHttpExceptions: true,
  });

  Logger.log(`HTTP ${res.getResponseCode()}`);
  Logger.log(res.getContentText());
}

/**
 * PeopleID取得GAS
 * BSOL情報
 * F列：メールアドレス
 * H列：Google Chat USER_ID を出力
 * 対象行：5〜31行
 */
function setChatUserIdToBSOL() {
  const SHEET_NAME = 'BSOL情報';
  const START_ROW = 4;
  const END_ROW = 31;

  const COL_EMAIL = 6; // F列
  const COL_USERID = 8; // H列

  const sh = SpreadsheetApp
    .openById(sheetId)
    .getSheetByName(SHEET_NAME);

  if (!sh) throw new Error('BSOL情報 シートが見つかりません');

  const numRows = END_ROW - START_ROW + 1;

  // F列（メール）取得
  const emails = sh
    .getRange(START_ROW, COL_EMAIL, numRows, 1)
    .getValues();

  const out = [];

  for (let i = 0; i < emails.length; i++) {
    const email = (emails[i][0] || '').toString().trim();

    if (!email) {
      out.push(['']);
      continue;
    }

    try {
      const user = AdminDirectory.Users.get(email, {
        viewType: 'domain_public',
      });
      out.push([user.id || '']);
    } catch (e) {
      // 取得できなかった場合は空欄（必要ならログ）
      console.warn(`USER_ID取得失敗: ${email}`, e.message);
      out.push(['']);
    }
  }

  // H列へ一括書き込み
  sh.getRange(START_ROW, COL_USERID, numRows, 1).setValues(out);
}
