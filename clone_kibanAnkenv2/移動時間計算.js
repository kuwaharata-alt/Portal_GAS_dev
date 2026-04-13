function generateMapAndCalculateRoutes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = "客先までの距離"; 
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Browser.msgBox("シート「" + SHEET_NAME + "」が見つかりません。");
    return;
  }
  
  // 客先住所（K2に変更）
  const destination = sheet.getRange("K2").getValue();
  if (!destination) {
    Browser.msgBox("K2セルに客先住所（または駅名）を入力してください");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // A2:D列（名前、最寄り駅1、路線、最寄り駅2）を取得
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 4); 
  const values = dataRange.getValues();
  
  const finalTimeResults = [];

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2; // スプレッドシート上の行番号
    const station1 = values[i][1]; // B列
    const station2 = values[i][3]; // D列
    
    // --- 1. Mapリンクの生成 (F列・G列) ---
    const baseUrl = "https://www.google.com/maps/dir/?api=1" +
                     "&destination=" + encodeURIComponent(destination) +
                     "&travelmode=transit" +
                     "&dirflg=r"; 

    if (station1) {
      const url1 = baseUrl + "&origin=" + encodeURIComponent(station1);
      sheet.getRange(rowIndex, 6).setFormula('=HYPERLINK("' + url1 + '", "電車ルート1")');
    } else {
      sheet.getRange(rowIndex, 6).setValue("-");
    }
    
    if (station2) {
      const url2 = baseUrl + "&origin=" + encodeURIComponent(station2);
      sheet.getRange(rowIndex, 7).setFormula('=HYPERLINK("' + url2 + '", "電車ルート2")');
    } else {
      sheet.getRange(rowIndex, 7).setValue("-");
    }

    // --- 2. 移動時間の計算 (H・I・J列) ---
    let allCandidates = [];
    const origins = [station1, station2].filter(s => s !== "" && s !== null);

    origins.forEach(origin => {
      try {
        const directions = Maps.newDirectionFinder()
          .setOrigin(origin)
          .setDestination(destination)
          .setMode(Maps.DirectionFinder.Mode.TRANSIT)
          .setLanguage('ja')
          .setAlternatives(true) 
          .getDirections();

        if (directions.routes && directions.routes.length > 0) {
          directions.routes.forEach(route => {
            allCandidates.push({
              seconds: route.legs[0].duration.value,
              text: route.legs[0].duration.text
            });
          });
        }
      } catch (e) {
        console.log("計算エラー: " + e.message);
      }
    });

    // 短い順にソートして上位3つを抽出
    allCandidates.sort((a, b) => a.seconds - b.seconds);
    const top3 = [
      allCandidates[0] ? allCandidates[0].text : "-",
      allCandidates[1] ? allCandidates[1].text : "-",
      allCandidates[2] ? allCandidates[2].text : "-"
    ];
    finalTimeResults.push(top3);
  }

  // H・I・J列に時間を一括書き込み
  sheet.getRange(2, 8, finalTimeResults.length, 3).setValues(finalTimeResults);
  
  Browser.msgBox("完了しました！\nF-G列: Mapリンク\nH-J列: 最短3候補の時間\nを確認してください。");
}