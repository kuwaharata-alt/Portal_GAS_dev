function hha02_AnnounceWed() {
  // ====== 毎週水曜日に案件管理表の更新を促すチャットを送信する =========

  // 今日の曜日を確認
  const today = new Date();

  // フォーマットを整える（例: YYYY/MM/DD）
  const year1 = today.getFullYear();
  const month1 = String(today.getMonth() + 1).padStart(2, '0'); // 月は0から始まるので+1
  const day1 = String(today.getDate()).padStart(2, '0');

  const formattedDate1 = `${year1}/${month1}/${day1}`;

  let message = `【${formattedDate1} 更新通知】\n`;
  message += `　明日の10:30までに案件管理表を更新してください。 `;
  message += `https://docs.google.com/spreadsheets/d/1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II/edit?gid=517496344#gid=517496344`
  // チャットに投稿
  hhfc_postToChat(message);
}

function hha02_AnnounceThu() {
  // ====== 毎週木曜日に案件管理表の更新日が古い案件を送信する =========

  // スプレッドシートを取得
  var sheetId = '1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II';
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('SV案件管理表');
  const data = sheet.getDataRange().getValues();

  // 今日の曜日を確認
  const today = new Date();

  // フォーマットを整える（例: YYYY/MM/DD）
  const year1 = today.getFullYear();
  const month1 = String(today.getMonth() + 1).padStart(2, '0'); // 月は0から始まるので+1
  const day1 = String(today.getDate()).padStart(2, '0');
  const formattedDate1 = `${year1}/${month1}/${day1}`;

  // メッセージ内容を準備
  let message = `【${formattedDate1} 更新催促(SV案件)】\n以下の案件を更新してください。\n`;

  // 除外するB列の値
  const excludeValues = [
    "81.【作業対象外】\nCLDチーム対応",
    "82.【作業対象外】\nCL案件",
    "83.【作業対象外】\nキャンセル",
    "84.【作業対象外】\n保守",
    "85.【作業対象外】\nプリ/BP対応",
    "98. SE対応完了",
    "99. クローズ",
    "A1.【定常案件】\n対応中",
    "A2.【定常案件】\nクローズ"
  ];

  // 担当者ごとの案件を格納するオブジェクト
  const responsibleMap = {};

  // A列のデータを確認して古いものを収集
  for (let i = 1; i < data.length; i++) { // 1行目はヘッダーと仮定
    const updateDate = data[i][0]; // A列: 更新日
    const status = data[i][1]; // B列: ステータス
    const projectName = data[i][4]; // E列: 案件名
    const responsiblePerson = data[i][9]; // J列: 担当者

    // B列の値が除外リストに含まれる場合はスキップ
    if (excludeValues.includes(status)) continue;

    // B列が空白の場合はスキップ
    if (!status || status.toString().trim() === "") continue;

    // 更新日が空欄の場合は対象に追加
    if (!updateDate || updateDate.toString().trim() === "") {
      if (!responsibleMap[responsiblePerson]) {
        responsibleMap[responsiblePerson] = [];
      }
      responsibleMap[responsiblePerson].push({ name: projectName, days: daysPassed });
      continue;
    }

    // 更新日が有効か確認
    const updateDateObj = new Date(updateDate);
    if (!hhfc_isValidDate(updateDateObj)) continue;

    // 経過日数を計算
    daysPassed = Math.floor((today - updateDateObj) / (1000 * 60 * 60 * 24));

    // 更新日が5日以上前なら対象にする
    if (daysPassed > 5) {
      if (!responsibleMap[responsiblePerson]) {
        responsibleMap[responsiblePerson] = [];
      }
      responsibleMap[responsiblePerson].push({ name: projectName, days: 0 });
    }
  }

  // 担当者ごとにメッセージを作成
  for (const [responsible, projects] of Object.entries(responsibleMap)) {
    
    message += `\n〇 ${responsible}\n`;
    projects.forEach(project => {
        // 更新日が空白だった場合はdaysが0なので、days > 0 の時だけ経過日数を表示
        const daysInfo = (project.days > 0) ? `  （${project.days} 日経過）` : "";
        message += `　・${project.name}${daysInfo}\n`;
    });
  }


  message += `\n\n\n【案件管理表URL】\n`;
  message += `https://docs.google.com/spreadsheets/d/1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II/edit?gid=517496344#gid=517496344`

  // チャットに投稿
  hhfc_postToChat(message);
}

function hhfc_isValidDate(d) {
  // ====== 日付が有効か確認する関数 ====== 
  return d instanceof Date && !isNaN(d);
}

function hhfc_postToChat(message) {
  // ====== Google Chatなどにメッセージを送る関数 ====== 
 // const webhookUrl = "https://chat.googleapis.com/v1/spaces/AAAAZnHwRBY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=K5jYsdl_WTJP1sj9Q3WaRMWXO1tQ5a8srfQ4fYX9_qg"; // Webhook URLを設定
  const webhookUrl = "https://chat.googleapis.com/v1/spaces/AAQAaqoOwVw/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=ljAWx8OcDHHzx3cwnxo4V44qPrYtCxlMkvWSpgm8iOU"; // 萩谷テスト

  const payload = JSON.stringify({ text: message });
  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
  };

  UrlFetchApp.fetch(webhookUrl, options);
}