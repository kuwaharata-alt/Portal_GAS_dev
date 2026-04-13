/** 
 * 作業依頼登録 
 **/

function Auto_作業依頼登録() {

  // フォーマットを整える（例: YYYY/MM/DD）
  const year1 = today.getFullYear();
  const month1 = String(today.getMonth() + 1).padStart(2, '0');
  const day1 = String(today.getDate()).padStart(2, '0');

  const year2 = yesterday.getFullYear();
  const month2 = String(yesterday.getMonth() + 1).padStart(2, '0');
  const day2 = String(yesterday.getDate()).padStart(2, '0');

  const formattedDate1 = `${year1}/${month1}/${day1}`;
  const formattedDate2 = `${year2}/${month2}/${day2}`;

  //const formattedDate1 = `2026/4/8`;  //before
  //const formattedDate2 = `2026/04/07`;  //after


  // Gmailの検索条件を指定してメールを取得
  var criteria =
    'label:01-it基盤-作業依頼 ' +
    'after:' + formattedDate2 + ' before:' + formattedDate1 +
    ' -subject:Re: -subject:【案件相談】 ' +
    'from:"ソリューション推進部 案件受付" <it-kiban@systena.co.jp>';

  var threads = GmailApp.search(criteria);
  threads.reverse();

  writeLog_('Auto_作業依頼登録', 'threads: ' + threads.length);

  const sh = getSH_('作業依頼');
  var row = sh.getLastRow() + 1;

  for (var i = 0; i < threads.length; i++) {

    var messages = threads[i].getMessages();
    var message = messages[0];

    var body = message.getPlainBody();

    // 正規表現で抽出
    var salesRepMatch = body.match(/【担当営業】\s*([^\n]+)/);
    var solNumberMatch = body.match(/【SOL推進部の案件番号】\s*([^\n]+)/);
    var estimateMatch = body.match(/【受注したSOL推進部の見積書番号】\s*([^\n]+)/);
    var categoryMatch = body.match(/【作業カテゴリ】\s*([^\n]+)/);

    var salesRep = salesRepMatch ? salesRepMatch[1].trim() : '';
    var solNumber = solNumberMatch ? solNumberMatch[1].trim() : '';
    var estimateRaw = estimateMatch ? estimateMatch[1].trim() : '';
    var categoryS = categoryMatch ? categoryMatch[1].trim() : '';

    // --- 修正(2): 見積番号が2個以上入る場合は分割して複数行出力 ---
    // 例）440001-01、440001-02 / 440001-01,440001-02 / 改行区切り などに対応
    var estimateList = splitEstimateNumbers_(estimateRaw);
    if (estimateList.length === 0) estimateList = ['']; // 何も取れない場合も1行は出す

    // 作業カテゴリ（J列の文字列からK列に値貼り）
    // --- 修正(1): 数式ではなく値で貼る ---
    // 既存ロジック：倉庫作業のみ or PC系 → PC系、それ以外 → SV系
    // ※本文の「【作業カテゴリ】」が間違うことはあるが、作業依頼はログなのでここはそのまま記録
    var workType = judgeWorkType_(categoryS);

    for (var e = 0; e < estimateList.length; e++) {
      var estimateNumber = estimateList[e];

      // 転記
      sh.getRange(row, 1).setValue(message.getDate());     // A: 日付
      sh.getRange(row, 2).setValue(message.getTo());       // B: 宛先
      sh.getRange(row, 3).setValue(message.getFrom());     // C: From
      sh.getRange(row, 4).setValue(message.getReplyTo());  // D: ReplyTo
      sh.getRange(row, 5).setValue(message.getSubject());  // E: 件名
      sh.getRange(row, 6).setValue(solNumber);             // F: SOL推進部の案件番号
      sh.getRange(row, 7).setValue(estimateNumber);        // G: 受注したSOL推進見積（分割後）
      // H: 顧客名（件名から抽出：従来通り数式）
      sh.getRange(row, 9).setValue(salesRep);              // I: 担当営業
      sh.getRange(row, 10).setValue(categoryS);            // J: 作業カテゴリ（原文）
      sh.getRange(row, 11).setValue(workType);             // K: 分類

      // 件名から顧客名を抽出する関数（従来通り）
      var com = '=MID(E' + row + ',15, SEARCH("（", E' + row + ')-15)';
      sh.getRange(row, 8).setFormula(com);                 // H: 顧客名

      Logger.log('No.' + i + '/日付:' + message.getDate());
      row++;
    }
  }
}

/**
 * 見積番号を分割して配列で返す
 */
function splitEstimateNumbers_(estimateRaw) {
  if (!estimateRaw) return [];

  // 全角カンマ/読点/セミコロン/改行/タブ/スペースなどで分割
  var normalized = estimateRaw
    .replace(/，/g, ',')
    .replace(/、/g, ',')
    .replace(/　/g, ',')
    .replace(/&/g, ';')
    .replace(/；/g, ';')
    .replace(/[;\n\r\t ]+/g, ','); // まとめて区切り扱い

  return normalized
    .split(',')
    .map(s => (s || '').trim())
    .filter(s => s !== '');
}

/**
 * J列(作業カテゴリ原文)から K列(SV系/PC系) を判定
 * 既存式:
 * =if(or(countif(J,"*倉庫作業のみ*"),countif(J,"*PC系*")),"PC系","SV系")
 */
function judgeWorkType_(categoryS) {
  var s = (categoryS || '').trim();
  if (s === '') return '本社';

  if (s.indexOf('倉庫作業のみ') !== -1) return '倉庫';
  if (s.indexOf('PC系') !== -1) return '本社';

  // 念のため（本文にSV系明記があればSV系）
  if (s.indexOf('SV系') !== -1) return '本社';

  return '本社';
}
