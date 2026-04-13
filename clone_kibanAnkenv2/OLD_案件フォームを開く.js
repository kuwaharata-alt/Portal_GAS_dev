function updatePeopleList(){
  //アクティブなスプレッドシートから名前付き範囲"担当者名"の値を一次元配列で取得
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let people = ss.getRangeByName("桑原拓也").getValues().flat();
  
  //フォームの対象項目のプルダウンの選択肢を配列の値で更新
  let form = FormApp.openByUrl(ss.getFormUrl());  //連携しているフォームを取得 
  let items = form.getItems();                //フォーム内の全アイテムを取得
  let item = items[2];                        //0番目から数えて何番目のアイテムか指定
  item.asListItem().setChoiceValues(people);  //配列peopleの値でプルダウンの選択肢を更新
}

function openUrl_CL() {
  var url = "https://forms.gle/LqXwhXwc47v9pj8B6"; // 開きたいURLを指定します
  var html = HtmlService.createHtmlOutput('<html><script>'
        +'window.open("'+url+'", "_blank");'
        +'google.script.host.close();'
        +'</script>'
        +'<body>Failed to open automatically. <a href="'+url+'" target="_blank">Click here to proceed</a>.</body>'
        +'</html>')
    .setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ...");
}

function openUrl_SV() {
  var url = "https://docs.google.com/forms/d/e/1FAIpQLSf5_PW7O27Yu1b1WyhyyGuEDmirYOnSFWCKen8vTqVeNDnV2Q/viewform"; // 開きたいURLを指定します
  var html = HtmlService.createHtmlOutput('<html><script>'
        +'window.open("'+url+'", "_blank");'
        +'google.script.host.close();'
        +'</script>'
        +'<body>Failed to open automatically. <a href="'+url+'" target="_blank">Click here to proceed</a>.</body>'
        +'</html>')
    .setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ...");
}