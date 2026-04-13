/** 
 * =======================================================================
 * 
 *【更新通知1】
 * ・毎週水曜日に案件管理表の更新を促すチャットを送信する
 *  
 * ======================================================================== 
 **/ 

function Gen_AnnounceWed() {
  // フォーマットを整える（例: YYYY/MM/DD）
  const year1 = today.getFullYear();
  const month1 = String(today.getMonth() + 1).padStart(2, '0'); // 月は0から始まるので+1
  const day1 = String(today.getDate()).padStart(2, '0');

  const formattedDate1 = `${year1}/${month1}/${day1}`;

  // 通知内容
  let message = `【⚠️${formattedDate1} 案件管理表 更新通知⚠️】\n`;
  message += `　明日の10:30までに案件管理表を更新してください。 `;
  message += `https://docs.google.com/spreadsheets/d/1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II/edit?gid=517496344#gid=517496344`

  // チャットに投稿
  fc_postToChat(webhookUrl_G, message);
}

function Gen_AnnounceThu() {
  // スプレッドシートを取得
  const data = sheetSV.getDataRange().getValues();

  today.setHours(0, 0, 0, 0); 

  // メンション用マップ（BSOL情報：氏名/メール→USER_ID）
  const mentionMap = buildMentionMapFromBSOL_();

  // フォーマットを整える（例: YYYY/MM/DD）
  const year1 = today.getFullYear();
  const month1 = String(today.getMonth() + 1).padStart(2, '0');
  const day1 = String(today.getDate()).padStart(2, '0');
  const formattedDate1 = `${year1}/${month1}/${day1}`;

  // メッセージ内容を準備
  let message = `【💥${formattedDate1} 案件管理表 更新催促(SV案件)💥】\n以下の案件を更新してください。\n`;

  // 除外するB列の値
  const excludeValues = [
    "81.【作業対象外】\nPSSEチーム対応",
    "82.【作業対象外】\nCL案件",
    "83.【作業対象外】\nキャンセル",
    "84.【作業対象外】\n保守",
    "85.【作業対象外】\nプリ/BP対応",
    "98. SE対応完了",
    "99. クローズ",
    "A1.【定常案件】\n対応中",
    "A2.【定常案件】\nクローズ"
  ];

  // 担当者ごとの案件を格納（キーは「メンション文字列 or 氏名」）
  const responsibleMap = {};

  // A列のデータを確認して古いものを収集
  for (let i = 1; i < data.length; i++) {
    const updateDate = data[i][0];   // A列: 更新日
    const status = data[i][1];       // B列: ステータス
    const projectName = data[i][4];  // E列: 案件名
    const responsiblePerson = data[i][14]; // O列: 担当者

    // 除外ステータス
    if (excludeValues.includes(status)) continue;
    if (!status || status.toString().trim() === "") continue;

    // 担当者ラベル（メンション）
    const responsibleLabel = toMentionLabel_(responsiblePerson, mentionMap);

    // 更新日が空欄 → 対象に追加（未更新日数は空扱いにしないなら0等でもOK）
    if (!updateDate || updateDate.toString().trim() === "") {
      if (!responsibleMap[responsibleLabel]) responsibleMap[responsibleLabel] = [];
      responsibleMap[responsibleLabel].push({ projectName, daysUnupdated: '' });
      continue;
    }

    // 更新日が有効か確認
    const updateDateObj = new Date(updateDate);
    updateDateObj.setHours(0, 0, 0, 0);
    if (!fc_isValidDate(updateDateObj)) continue;

    // 未更新日数を計算
    const daysUnupdated = Math.floor((today - updateDateObj) / (1000 * 60 * 60 * 24));

    // 更新日が5日以上前なら対象にする（>5 は元ロジック維持）
    if (daysUnupdated > 5) {
      if (!responsibleMap[responsibleLabel]) responsibleMap[responsibleLabel] = [];
      responsibleMap[responsibleLabel].push({ projectName, daysUnupdated });
    }
  }

  // 0件なら送らない（要件：案件が一つもなければ送信しない）
  if (Object.keys(responsibleMap).length === 0) return;

  // 担当者ごとにメッセージを作成
  for (const [responsible, projects] of Object.entries(responsibleMap)) {
    message += `\n〇 ${responsible}\n`;
    projects.forEach(item => {
      if (item.daysUnupdated === '') {
        message += `　[更新日 未入力]　${item.projectName}\n`;
      } else {
        message += `　[${item.daysUnupdated}日 未更新]　${item.projectName}\n`;
      }
    });
  }

  message += `\n\n【案件管理表v2】\n`;
  message += `https://docs.google.com/spreadsheets/d/1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II/edit?gid=517496344#gid=517496344`;

  // チャットに投稿
  fc_postToChat(webhookUrl_G, message);
}

