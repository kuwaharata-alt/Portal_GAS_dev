/** 
 * =======================================================================
 * 
 *【更新通知3-1】
 * ・案件が登録された日に案件アサイン依頼をチャットへ通知する
 *  
 * ======================================================================== 
 **/ 

function SV_AssignAnnounce() {
  // 月〜金(1〜5)のみ実行
  const day = today.getDay(); // 0:日曜, 1:月曜, ..., 6:土曜

  if (day === 0 || day === 6) {
    return; // 土日なら何もせず終了
  }
    
  const targetStr = fc_getPrevBusinessDayStr(tz); // ★ 1営業日前(yyyy/MM/dd)
  
  const lastCol = Math.max(sheetSV.getLastColumn(), 7); // 少なくともGまで
  if (lastRowSV < 2) return;

  const rangeAll = sheetSV.getRange(1, 1, lastRowSV, lastCol);
  const values   = rangeAll.getValues();
  const displays = rangeAll.getDisplayValues();

  const formulasD = sheetSV.getRange(1, 4, lastRowSV, 1).getFormulas();
  const richD     = sheetSV.getRange(1, 4, lastRowSV, 1).getRichTextValues();
  const formulasF = sheetSV.getRange(1, 6, lastRowSV, 1).getFormulas();
  const richF     = sheetSV.getRange(1, 6, lastRowSV, 1).getRichTextValues();

  for (let r = 2; r <= lastRowSV; r++) {
    const i = r - 1;

    // ★ G列（実行日）＝ 1営業日前
    const gStr = normalizeDateToYmd(values[i][6], displays[i][6], tz);
    if (gStr !== targetStr) continue;

    // C/D/E/F
    const cTxt = safeStr(displays[i][2]); // C
    const dTxt = safeStr(displays[i][3]); // D 表示
    const eTxt = safeStr(displays[i][4]); // E
    const fTxt = safeStr(displays[i][5]); // F 表示

    const dLink = extractLinkForCell(dTxt, formulasD[i][0], richD[i][0]);
    const fLink = extractLinkForCell(fTxt, formulasF[i][0], richF[i][0]);

    // --- カード構築（URLはボタンにのみ設定し本文には出さない） ---
    const widgets = [
      { keyValue: { content: cTxt || '-' } },
      { keyValue: { content: eTxt || '-' } },
    ];

    const btns = [];
    if (dLink.url) { btns.push({ textButton: { text: '見積り', onClick: { openLink: { url: dLink.url } } } }); }
    else           { widgets.push({ keyValue: { content: dTxt || '-' } }); }

    if (fLink.url) { btns.push({ textButton: { text: '作業フォルダ', onClick: { openLink: { url: fLink.url } } } }); }
    else           { widgets.push({ keyValue: { content: fTxt || '-' } }); }

    if (btns.length) widgets.push({ buttons: btns });

    // ★ サブタイトルも対象日を表示
    const cardMsg = {
      text:"【" + cTxt +"】"+ eTxt, 
      cards: [{
        header: { 
          title: '【案件アサイン依頼】', 
          subtitle: '依頼日：' + targetStr 
        },
        sections: [{ widgets }]
      }]
    };

    fc_postToChat03(cardMsg);   // ← ラッパー経由で送信
    Utilities.sleep(300);
  }
}

/** 
 * =======================================================================
 * 
 *【更新通知3-2】
 * ・手動実行
 * ・案件アサイン依頼をチャットへ通知する
 *  
 * ======================================================================== 
 **/ 

function SV_AssignAnnounce_manual() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('案件の見積もり番号（C列）を入力してください', '', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = (res.getResponseText() || '').trim();
  if (!key) { ui.alert('番号が未入力です'); return; }
  a03_AssignAnnounce_ByNumber(key);
}

// === 本体：引数の番号とC列が一致する行を1件ずつチャット ===
function a03_AssignAnnounce_ByNumber(key) {

  const lastCol = Math.max(sheetSV.getLastColumn(), 7);

  const rangeAll = sheetSV.getRange(1, 1, lastRowSV, lastCol);
  const displays = rangeAll.getDisplayValues();

  // D / F の式・リッチテキスト（リンク抽出用）
  const formulasD = sheetSV.getRange(1, 4, lastRowSV, 1).getFormulas();
  const richD     = sheetSV.getRange(1, 4, lastRowSV, 1).getRichTextValues();
  const formulasF = sheetSV.getRange(1, 6, lastRowSV, 1).getFormulas();
  const richF     = sheetSV.getRange(1, 6, lastRowSV, 1).getRichTextValues();

  let hit = 0;
  for (let r = 2; r <= lastRowSV; r++) {
    const i = r - 1;

    // ★ C列の一致判定（表示値でトリム比較）
    const cKey = safeStr(displays[i][2]).trim(); // C列
    if (cKey !== key) continue;
    hit++;

    // C/D/E/F 値
    const cTxt = cKey;
    const dTxt = safeStr(displays[i][3]);
    const eTxt = safeStr(displays[i][4]);
    const fTxt = safeStr(displays[i][5]);

    const dLink = extractLinkForCell(dTxt, formulasD[i][0], richD[i][0]); // 既存のヘルパーを利用
    const fLink = extractLinkForCell(fTxt, formulasF[i][0], richF[i][0]);

    // --- classic cards（本文にURLは出さず、ボタンにだけ設定） ---
    const widgets = [
      { keyValue: { content: `${cTxt || '-'}` } },
      { keyValue: { content: eTxt || '-' } },
    ];

    const btns = [];
    if (dLink.url) { btns.push({ textButton: { text: '見積り',    onClick: { openLink: { url: dLink.url } } } }); }
    else           { widgets.push({ keyValue: { content: dTxt || '-' } }); }

    if (fLink.url) { btns.push({ textButton: { text: '作業フォルダ', onClick: { openLink: { url: fLink.url } } } }); }
    else           { widgets.push({ keyValue: { content: fTxt || '-' } }); }

    if (btns.length) widgets.push({ buttons: btns });

    const cardMsg = {
      text:"【" + cTxt +"】"+ eTxt, 
      cards: [{
        header: { title: '【案件アサイン依頼】', subtitle: `${todayStr}` },
        sections: [{ widgets }]
      }]
    };

    fc_postToChat03(cardMsg);   // 1件ずつ投下
    Utilities.sleep(300);
  }
}
/** 
 * =======================================================================
 *  
 *【更新通知4】
 * ・未アサイン一覧をチャットへ投稿
 *  
 * ======================================================================= 
 **/

function SV_PostUnassignedAnnounce() {
  // 月・水・金(1〜5)のみ実行
  const today = new Date();
  const day = today.getDay(); // 0:日曜, 1:月曜, ..., 6:土曜

  if (day === 0 || day === 2 || day === 4 || day === 6) {
    return;
  }

  const lastCol = Math.max(sheetSV.getLastColumn(), 7);
  if (lastRowSV < 2) return;

  const values = sheetSV.getRange(2, 1, lastRowSV - 1, lastCol).getValues();
  let bodyLines = [];

  // 条件に一致する行を抽出
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const b = row[1]; // B列
    const c = row[2]; // C列
    const e = row[4]; // D列
    const g = row[6]; // G列

    if (b && b.toString().includes("アサイン未")) {
      let line = `【${c || ''}】　${e || ''}`;

      if (g instanceof Date) {
        // 経過日数を計算
        const diffMs = today.getTime() - g.getTime();
        const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

        const gStr = Utilities.formatDate(g, tz, 'yyyy/MM/dd');
        line += `\n依頼日:${gStr}（${diffDays}日経過）`;
      }

      bodyLines.push(line, ""); // 空行
    }
  }

  // 投稿する内容がなければ終了
  if (bodyLines.length === 0) return; 

  // 投稿メッセージを作成
  const message = [
    `${todayStr}`,
    `以下の案件が未アサインです。`,
    bodyLines.join("\n"),
  ].join("\n");

  fc_postToChat03(message);   // 1件ずつ投下
}


/** 
 * =======================================================================
 *  
 *【Chat へメッセージを送る関数】
 * ※作業依頼用
 * 引数：記載するカードメッセージ
 *  
 * ======================================================================== 
 **/ 

function fc_postToChat03(message) {

  const payloadObj = (typeof message === 'string') ? { text: message } : message;

  // 送信（失敗時は classic cards にフォールバック）
  try {
    const res = UrlFetchApp.fetch(webhookUrl_C, {
      method: "post",
      contentType: "application/json; charset=utf-8",
      payload: JSON.stringify(payloadObj),
    });
    return;
  } catch (e) {
    const body = e?.response?.getContentText?.() || String(e);

    // cardsV2非対応などで弾かれた場合は classic cards に変換して再送
    if (payloadObj.cardsV2) {
      const classic = fc_convertCardsV2ToClassic(payloadObj);
      const res2 = UrlFetchApp.fetch(webhookUrl_C, {
        method: "post",
        contentType: "application/json; charset=utf-8",
        payload: JSON.stringify(classic),
      });
    } else {
      throw e;
    }
  }
}

/** 
 * =======================================================================
 * 
 * cardsV2 → classic cards への簡易変換
 *  
 * ======================================================================== 
 **/ 

function fc_convertCardsV2ToClassic(msg) {
  const cards = (msg.cardsV2 || []).map(c => {
    const cc = c.card || {};
    return {
      header: cc.header ? { title: cc.header.title, subtitle: cc.header.subtitle } : undefined,
      sections: cc.sections || [],
    };
  });
  const out = {};
  if (msg.text) out.text = msg.text;
  out.cards = cards;
  return out;
}

/** 
 * =======================================================================
 * 
 * ヘルパー
 *  
 * ======================================================================== 
 **/ 

function normalizeDateToYmd(value, display, tz) {
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'yyyy/MM/dd');
  if (typeof display === 'string' && display) {
    const s = display.trim().replace(/-/g, '/');
    const m = s.match(/^(\d{4})[\/](\d{1,2})[\/](\d{1,2})/);
    if (m) return `${m[1]}/${('0'+m[2]).slice(-2)}/${('0'+m[3]).slice(-2)}`;
  }
  return '';
}

function safeStr(v){ return (v===null||v===undefined) ? '' : String(v); }

/** 
 * =======================================================================
 * 
 * 【参照型HYPERLINKにも対応するリンク抽出】
 *  
 * ======================================================================== 
 **/ 

function extractLinkForCell(displayText, formula, richTextValue) {
  // 1) HYPERLINK("url","text")
  if (formula && /^=HYPERLINK\(/i.test(formula)) {
    // 文字リテラル2引数
    let m = formula.match(/^=HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)/i);
    if (m) return { url: m[1], text: m[2] || displayText };

    // =HYPERLINK(B2,"text") 等：第1引数がセル参照なら、その表示値をURL候補に
    const ref = getFirstArgCellRef(formula);
    if (ref) {
      try {
        const rng = SpreadsheetSVApp.getActivesheetSV().getRange(ref);
        const candidate = rng.getDisplayValue();
        if (candidate && /^https?:\/\//i.test(candidate)) {
          const t = (formula.match(/^=HYPERLINK\(\s*[^,]+,\s*"([^"]*)"\s*\)/i) || [])[1] || displayText;
          return { url: candidate, text: t };
        }
      } catch (e) {
        // 参照解決できない場合はスルー
      }
    }
  }

  // 2) リッチテキストリンク（セル全体 or 部分）
  if (richTextValue) {
    const whole = richTextValue.getLinkUrl && richTextValue.getLinkUrl();
    if (whole) return { url: whole, text: richTextValue.getText() };
    if (richTextValue.getRuns) {
      const runs = richTextValue.getRuns();
      for (const run of runs) {
        const u = run.getLinkUrl && run.getLinkUrl();
        if (u) return { url: u, text: run.getText() };
      }
    }
  }

  // 3) 表示値がURL
  if (/^https?:\/\//i.test(displayText)) return { url: displayText, text: displayText };

  return { url: null, text: displayText };
}

/** =======================================================================
 * 
 * 【HYPERLINK第1引数のセル参照を取り出す】
 *  
 * ======================================================================== 
 **/ 
function getFirstArgCellRef(formula) {
  const m = formula.match(/^=HYPERLINK\(\s*([^,]+)\s*,/i);
  if (!m) return null;
  const arg = m[1].trim();

  // 'シート名'!A1 形式
  let ms = arg.match(/^'([^']+)'!([$]?[A-Za-z]+[$]?\d+)$/);
  if (ms) return `'${ms[1]}'!${ms[2]}`;

  // A1 / $A$1 形式
  let ma1 = arg.match(/^([$]?[A-Za-z]+[$]?\d+)$/);
  if (ma1) return ma1[1];

  return null;
}