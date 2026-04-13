/** 
 * =======================================================================
 * 
 * 【作業依頼登録】
 *  
 * ======================================================================== 
 **/ 

function Gen_WorkRequestRegistration() {

  // フォーマットを整える（例: YYYY/MM/DD）
  const year1 = today.getFullYear();
  const month1 = String(today.getMonth() + 1).padStart(2, '0'); // 月は0から始まるので+1
  const day1 = String(today.getDate()).padStart(2, '0');

  const year2 = yesterday.getFullYear();
  const month2 = String(yesterday.getMonth() + 1).padStart(2, '0'); // 月は0から始まるので+1
  const day2 = String(yesterday.getDate()).padStart(2, '0');

  const formattedDate1 = `${year1}/${month1}/${day1}`;
  const formattedDate2 = `${year2}/${month2}/${day2}`;


  // Gmailの検索条件を指定してメールを取得
  var criteria = 'label:01-it基盤-作業依頼  after:' + formattedDate2 +' before:'+ formattedDate1 +' -subject:Re: -subject:【案件相談】 from:"ソリューション推進部 案件受付" <it-kiban@systena.co.jp>'
    Logger.log(criteria);
  var threads = GmailApp.search(criteria); // 未読かつ返信メールを除外

  var lastrow = sheetReq.getLastRow(); // A列の最終行を取得
  row = lastrow +1 
    Logger.log(row);
    Logger.log(threads.length)
  for (var i = 0; i < threads.length; i++) {
    // row行に1行追加

    messages = threads[i].getMessages();
    message = messages[0];
    console.log(message.getFrom());

    var body = message.getPlainBody();  // メールの本文（プレーンテキスト）

    // 正規表現を使用して、本文から必要な項目を抽出
    var salesRepMatch = body.match(/【担当営業】\s*([^\n]+)/);  // 担当営業
    var solNumberMatch = body.match(/【SOL推進部の案件番号】\s*([^\n]+)/);  // SOL推進部の案件番号
    var estimateNumberMatch = body.match(/【受注したSOL推進部の見積書番号】\s*([^\n]+)/);  // 見積書番号
    var categoryMatch = body.match(/【作業カテゴリ】\s*([^\n]+)/); //作業カテゴリ

    // 抽出した内容を変数に代入（見つからない場合は空文字）
    var salesRep = salesRepMatch ? salesRepMatch[1].trim() : '';
    var solNumber = solNumberMatch ? solNumberMatch[1].trim() : '';
    var estimateNumber = estimateNumberMatch ? estimateNumberMatch[1].trim() : '';
    var categoryS = categoryMatch ? categoryMatch[1].trim() : '';

    // スプレッドシートに転記
    sheetReq.getRange(row, 1).setValue(message.getDate());
    sheetReq.getRange(row, 2).setValue(message.getTo());
    sheetReq.getRange(row, 3).setValue(message.getFrom());
    sheetReq.getRange(row, 4).setValue(message.getReplyTo());
    sheetReq.getRange(row, 5).setValue(message.getSubject());
    sheetReq.getRange(row, 6).setValue(solNumber);
    sheetReq.getRange(row, 7).setValue(estimateNumber);
    sheetReq.getRange(row, 9).setValue(salesRep);
    sheetReq.getRange(row, 10).setValue(categoryS);

    //件名から顧客名を抽出する関数を入力
    var com = '=MID(E' + row + ',15, SEARCH("（", E' + row +')-15)'
    sheetReq.getRange(row, 8).setFormula(com);

    //作業カテゴリを抽出する関数を入力
    var cat = '=if(countif(J' + row + ',"*倉庫作業のみ*"),"PC系","SV系")'
    sheetReq.getRange(row, 11).setFormula(cat);
    row++;

  };
}