function e53_InputCalender() {
 // ====== 指定の期間カレンダーへ登録 ======
  const ui = SpreadsheetApp.getUi();

  // ログイン中のGoogleアカウントを取得
  const userEmail = Session.getActiveUser().getEmail();
  
  // 登録日を指定
  const startDateResponse = ui.prompt("開始日を入力してください（例：2/3）");
  const startDateInput = startDateResponse.getResponseText();
  const endDateResponse = ui.prompt("終了日を入力してください（例：2/7）");
  let endDateInput;
  if (endDateResponse.getResponseText() === ''){
    endDateInput = startDateInput;
  }else{
    endDateInput = endDateResponse.getResponseText();
  }

  // 現在の年を取得　※年を跨いだ登録はできない
  const currentYear = new Date().getFullYear();
  
  // 月日を入力された文字列から現在の年を付加して完全な日付を作成
  const startformattedDateInput = `${currentYear}/${startDateInput}`;
  const endformattedDateInput = `${currentYear}/${endDateInput}`;

  // 日付に変換
  const startDate = new Date(startformattedDateInput);
  const endDate = new Date(endformattedDateInput);
  
  // 入力された値が有効な日付かどうかを確認
  if (isNaN(startDate.getTime())) {
    ui.alert("開始日が無効な日付です。正しい形式で入力してください。");
    return;
  }
  // 入力された値が有効な日付かどうかを確認
  if (isNaN(endDate.getTime())) {
    ui.alert("終了日が無効な日付です。正しい形式で入力してください。");
    return;
  }

  // カレンダーへ登録
  var sheetId = '1UZJV7v-q837jjsfD5atYmbvuVaR0v0w3rhFnYebl6II';
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('予定管理表');
  const calendar = CalendarApp.getCalendarById(userEmail);
  const dateColumn = 1; // A列が日付が格納されている列
  const column = fc_getColumnForEmail(userEmail)

  var month = startDate.getMonth() + 1

  //A列で検索の開始と終了行を取得
  const calenderStart = fc_getCalenderCell(month);
  const calenderEnd = calenderStart + 31;
  
  let out = '';
  let work = '';
  let output = '';

  
  for (let currentDate = new Date(startDate); currentDate <= new Date(endDate); currentDate.setDate(currentDate.getDate() + 1)) {
    const dateInput = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // 日付の行を検索
    for (let row = calenderStart; row <= calenderEnd; row++) {
      const dateCell = sheet.getRange(row, dateColumn).getValue();
      
        
      // 日付を比較する際は、時刻部分を無視して日付部分だけを比較
      const dateCellFormatted = new Date(dateCell).setHours(0, 0, 0, 0);
      const dateInputFormatted = new Date(dateInput).setHours(0, 0, 0, 0);


      if (dateCellFormatted === dateInputFormatted) {
        out = sheet.getRange(row, column).getValue();
        work = sheet.getRange(row, column + 1).getValue();
        output = fc_cat(out, work);
        break;  
      }
    }
    
    // out(場所)列がブランクだった場合、次の日へスキップする
    if (out === 0) {
      continue; // 次の日付に進む
    }
    // 場所、対応顧客列がブランクだった場合、次の日へスキップする
    if (out==='' && work===''){
      continue;
    }

    // 登録する時間を指定する
    const startTime = new Date(dateInput);
    startTime.setHours(9, 0, 0);  // 9:00 AM に設定
    const endTime = new Date(dateInput);
    endTime.setHours(18, 0, 0);  // 6:00 PM に設定

    let events = calendar.getEvents(startTime, endTime);

    //カレンダーに既に登録されている場合は削除する
    for (let event of events) {
      if (event.getTitle().includes("/") || event.getTitle().includes("休暇")) { 
        Logger.log(`削除対象のイベント: ${event.getTitle()}`);
        event.deleteEvent(); // 既存イベントを削除
      }
    }   

    // イベントをカレンダーに登録
    calendar.createEvent(output, startTime, endTime);
  
    // 次の日付をチェックするためにoutとworkをリセット
    out = '';
    work = '';
  }
  SpreadsheetApp.getUi().alert(`カレンダーにイベントを追加しました！`);
}

function fc_getCalenderCell(cal){
  // 月,開始行,終了行
    const start2 = 157
    const start3 = start2 + 28
    const start4 = start3 + 31
    const start5 = start4 + 30
    const start6 = start5 + 31
    const start7 = start6 + 30
    const start8 = start7 + 31
    const start9 = start8 + 31
    const start10 = start9 + 30
    const start11 = start10 + 31
    const start12 = start11 + 30

  const setCalenderMap = {
    2: start2,
    3: start3,
    4: start4,
    5: start5,
    6: start6,
    7: start7,
    8: start8,
    9: start9,
    10: start10,
    11: start11,
    12: start12,
  };

  if (cal in setCalenderMap) {
    return setCalenderMap[cal];
  }else{
    return -1
  }
}

function fc_getColumnForEmail(email) {
  // メールアドレスと対応する列を定義したオブジェクト
  const emailColumnMap = {
    "hanadary@systena.co.jp": 3,
    "oosaway@systena.co.jp": 5,
    "taniguchima@systena.co.jp": 7,
    "uchidas@systena.co.jp": 9,
    "hagiharan@systena.co.jp": 11,
    "satouhir@systena.co.jp": 13,
    "satoute@systena.co.jp": 15,
    "kuwaharata@systena.co.jp": 17,
    "itoukai@systena.co.jp": 19,
    "hagiyas@systena.co.jp": 21,
    "ashizawar@systena.co.jp": 23,
    "kitagawat@systena.co.jp": 25,
    "suginoy@systena.co.jp": 27,
    "endoug@systena.co.jp": 29,
    "nakanota@systena.co.jp": 31,
    "ishiwarit@systena.co.jp": 33,
    "ootamak@systena.co.jp": 35,
    "izumiry@systena.co.jp": 37
  };
  // メールアドレスがリストに存在するかチェック
  if (email in emailColumnMap) {
    return emailColumnMap[email]; // 対応する列を返す
  } else {
    return -1; // メールアドレスが見つからない場合
  }
}

function fc_cat(out, work) {
  // メールアドレスと対応する列を定義したオブジェクト
  const outColumnMap = {
    "現地": "外",
    "本社": "内",
    "候補": "仮"
  };

  if (out in outColumnMap) {
    output = outColumnMap[out] + "/" + work
    return output; // 対応する列を返す
  } else if(out === "休暇"){
    output = "休暇"
    return output;
  } else {
    return out + "/" + work; // 社内・社外の場合
  }
}


