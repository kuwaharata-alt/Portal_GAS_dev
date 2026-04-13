/** =======================================================================
 * 
 * SV案件管理表：ステータス「98. SE対応完了」時に
 * ・OK/CANCEL ポップアップ
 * ・OK → メール送信（完了メッセージ付き）
 *
 * メール宛先:
 *   To = hanadary@sysntea.co.jp（固定）
 *        + L列の氏名 → BSOL情報!G列一致 → F列メール
 *
 *   Cc = M列・N列の氏名 → BSOL情報!G列一致 → F列メール
 *        ※ AC列も将来的に追加可能（コメントアウト）
 *
 * 対象外ワード：「要確認」「-」「」(空欄)
 *
 * 件名 = 回答期限（今日+7日）
 * 
 * ======================================================================== 
 **/ 

function CloseMail_SV(row) {
  
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'SE対応完了',
    'SE対応完了しました。フィードバックメールを送信しますか？',
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.CANCEL) {
    const oldValue = (typeof e.oldValue !== 'undefined') ? e.oldValue : '';
    range.setValue(oldValue);
    return;
  }

  const mitsumori = String(sheetSV.getRange(row, 3).getValue()).trim();  // C列
  const customerName = String(sheetSV.getRange(row, 5).getValue()).trim(); // E列

  // --- 氏名（L/M/N/AC） ---
  const nameKAN = cleanName(sheetSV.getRange(row, SVCOL_KANTOKU).getValue()); // L列
  const nameMAIN = cleanName(sheetSV.getRange(row, SVCOL_MAIN).getValue()); // M列
  const nameSUP1 = cleanName(sheetSV.getRange(row, SVCOL_SUPPORT1).getValue()); // N列
  const nameSUP2 = cleanName(sheetSV.getRange(row, SVCOL_SUPPORT2).getValue()); // N列
  const namePR = cleanName(sheetSV.getRange(row, SVCOL_PRE_SALES).getValue()); // AC列（コメントアウト予定）

  // NGワードは除外
  const ignoreWords = ['要確認', '-', '', null];

  const isValidName = (name) => {
    return (
      name &&
      !ignoreWords.includes(String(name).trim())
    );
  };

  // BSOL情報 読み込み
  let mailKAN = '';
  let mailMAIN = '';
  let mailSUP1 = '';
  let mailSUP2 = '';
  let mailPR = '';  // コメントアウト対象

  if (sheetBSOL) {
    const lastRowBS = sheetBSOL.getLastRow();
    if (lastRowBS > 1) {
      const bsolValues = sheetBSOL.getRange(2, 6, lastRowBS - 1, 2).getValues();
      const findMail = (target) => {
        if (!isValidName(target)) return '';
        for (let i = 0; i < bsolValues.length; i++) {
          const nameG = String(bsolValues[i][1]).trim();
          if (nameG === target) return String(bsolValues[i][0]).trim();
        }
        return '';
      };

      mailKAN = findMail(nameKAN);
      mailMAIN = findMail(nameMAIN);
      mailSUP1 = findMail(nameSUP1);
      mailSUP2 = findMail(nameSUP2);
      // mailPR = findMail(namePR);   // ← ★今はコメントアウト
    }
  }
  
  //------ 回答期限（今日+7日）-----------------------
  const today = new Date();
  const pad = (n) => ('0' + n).slice(-2);
  const limit = new Date(today);
  limit.setDate(limit.getDate() + 7);
  const limitStr =
    limit.getFullYear() + '/' + pad(limit.getMonth() + 1) + '/' + pad(limit.getDate());
  

  // ----------- メール宛先 ------------
  const fixedTo = 'hanadary@systena.co.jp';

  const toList = [fixedTo];
  if (mailKAN && !toList.includes(mailKAN)) toList.push(mailKAN);
  const to = toList.join(',');

  const ccList = [];
  if (mailMAIN) ccList.push(mailMAIN);
  if (mailSUP1) ccList.push(mailSUP1);
  if (mailSUP2) ccList.push(mailSUP2);
  // if (mailPR) ccList.push(mailPR);  // ← コメントアウトで AC はまだ追加しない
  const cc = ccList.join(',');

  const subject =
    `【作業FB】【${mitsumori}】${customerName}（回答期限：${limitStr}）`;

  const body =
    '案件の完了が報告されました。\n\n' +
    '案件対応メンバー全員とMTGを実施し以下のGoogleフォームに回答してください。\n' +
    'https://forms.gle/A33LUWraVd88oCbA6\n';

  const options = {};
  if (cc) options.cc = cc;

  GmailApp.sendEmail(to, subject, body, options);

  // 完了メッセージ
  let msg = "フィードバック依頼メールを送信しました。\n\nTo:\n" + to;
  if (cc) msg += "\n\nCc:\n" + cc;
  ui.alert(msg);
}

/** 
 * =======================================================================
 * 
 * 【余計な空白などを除去する共通関数】
 * 
 * ======================================================================== 
 **/
function cleanName(value) {
  if (value == null) return '';
  return String(value).trim();
}