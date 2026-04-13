function notifyUncheckedTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName("社内タスク");
  const chatWebhookUrl2 = "https://chat.googleapis.com/v1/spaces/AAQA1Ou2jsg/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=XNb7GJ7zJmdr_UU3ku-ADYvCAI7gVM9V4YTaX0sez-0"; // ここにGoogle ChatのWebhook URLを入力

  if (!taskSheet) {
    Logger.log("シート「社内タスク」が見つかりません。");
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0); // 今日の日付の時間をリセット

  const lastRow = taskSheet.getLastRow();
  const taskNames = taskSheet.getRange("A3:A" + lastRow).getValues().flat();
  const dueDates = taskSheet.getRange("B3:B" + lastRow).getValues().flat();
  const memberNames = taskSheet.getRange("D2:V2").getValues().flat();
  const checkBoxes = taskSheet.getRange("D3:V" + lastRow).getValues();

  let messageBody = "";
  let taskFound = false;

  for (let i = 0; i < dueDates.length; i++) {
    const dueDate = new Date(dueDates[i]);
    dueDate.setHours(0, 0, 0, 0); // 期限日の時間をリセット
   
    if (dueDate.getTime() === today.getTime()) {
      taskFound = true;
      const taskName = taskNames[i];
      let uncheckedMembers = [];
    
      for (let j = 0; j < memberNames.length; j++) {
        if (checkBoxes[i][j] !== true) { // チェックボックスがOFFの場合
          uncheckedMembers.push("　・" + memberNames[j]);
        }
      }

      if (uncheckedMembers.length > 0) {
        messageBody += "〇タスク名：" + taskName + "\n";
        messageBody += uncheckedMembers.join("\n") + "\n\n";
      }
    }
  }

  if (taskFound && messageBody !== "") {
    const spreadsheetUrl = ss.getUrl() + "#gid=" + taskSheet.getSheetId();
    const formattedDate = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");
    const message = "【" + formattedDate + " タスク実施催促】\n" +
                    "　本日が期限のタスクがあります。完了状況の確認をお願いします。\n" +
                    "　以下のメンバーはタスクが完了しているか確認し、完了していれば管理表にチェックを入れてください。\n\n" +
                    messageBody +
                    "〇リンク\n" + spreadsheetUrl;
    sendGoogleChat2(message, chatWebhookUrl2);
  }
}

function sendGoogleChat2(message, webhookUrl) {
  const payload = {
    "text": message
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  try {
    UrlFetchApp.fetch(webhookUrl, options);
  } catch (e) {
    Logger.log("Google Chatへの送信に失敗しました: " + e);
  }
}


