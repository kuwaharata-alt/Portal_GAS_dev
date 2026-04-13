// ----- 日付変数 ------------------------------------------------------------
const tz = 'Asia/Tokyo';

const today = new Date(); // 現在の日付を取得
const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

const yesterday = new Date(today);
yesterday.setDate(today.getDate() - 1); // 日付を1日前に設定

const tommorow = new Date(today);
tommorow.setDate(today.getDate() + 1); // 日付を1日後に設定

// ----- オーナー譲渡先 ------------------------------------------------------------
const NEW_OWNER_EMAIL = 'it-kiban@systena.co.jp';

// ----- シート変数 ---------------------------------------------------------

// DB用スプレッドシート
function getSH_(key) {
  const k = String(key ?? '').trim();   // ← 空白・null対策

  var sheetId = '1czr0sPRCiAi9y0Yi-Dya4K5g-VLDlZFAEuZH63KWPKU';
  const ss = SpreadsheetApp.openById(sheetId);

  const map = {
    作業依頼: ss.getSheetByName('作業依頼'),
    計上管理_転写: ss.getSheetByName('計上管理_転写'),
    案件情報: ss.getSheetByName('案件情報'),
    案件管理表_本社: ss.getSheetByName('案件管理表_本社-現地'),
    案件管理表_倉庫: ss.getSheetByName('案件管理表_倉庫'),
    実行ログ: ss.getSheetByName('実行ログ'),
  }

  const id = map[k];
  if (!id) {
    throw new Error(`getSH_: unknown key "${key}"`);
  }
  return id;
}

// ----- WebhokURL ---------------------------------------------------------
function getWHU_(key) {
  const k = String(key ?? '').trim();   // ← 空白・null対策

  const map = {
    // *44基盤チームbot
    TeamBot: "https://chat.googleapis.com/v1/spaces/AAQA1Ou2jsg/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=5HzVdEkAYglMDoYRV-DKBNPRWwDC3oR8dWaRNTYihFg",
    // *案件確認
    ProjectCheck: "https://chat.googleapis.com/v1/spaces/AAAAZkbJFSA/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=kSZ3RZi6gb2Z2ZGUSIE2rUwj4iDEhwKvwKHporZTW9U",
    // *案件完了確認（案件確認スペース）
    ProjectComp: "https://chat.googleapis.com/v1/spaces/AAAAZkbJFSA/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=q1anpCPk9ghUhANyIDwriP-ONX9Km6UNqC7mjXJsm1Y",
    // テスト用
    Test: "https://chat.googleapis.com/v1/spaces/AAQAuHpWqBA/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=uw8WMfiJjh1BP8W-p_pWaenRQj01VW0Fzx1Dh_xgYwQ",
    // 確認用
    TestCheck: "https://chat.googleapis.com/v1/spaces/AAQApzvWPvQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=4QR3t3UsN5SFkF9T6qyQpOykOrWo-R3l38XJzas1TAg",
  };

  const id = map[k];
  if (!id) {
    throw new Error(`getWHU_: unknown key "${key}"`);
  }
  return id;
}

