// 最新情報の取得 --------------------------------------------------------------------------------------------------- //
function updateData() {
  var localSpread = SpreadsheetApp.getActiveSpreadsheet();
  var category = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "アクセサリー", "メイク"];
  var tgt = localSpread.getSheetByName("コーデ検索").getRange("D23").getValue();
  if (tgt == "アクセ以外") {
    category = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "メイク"];
    updateDataMain(category);
    updateDataOfficial(category);
  } else if (tgt == 'アクセサリー') {
    // アクセのときはアクセシートのみ（処理落ち対策）
    category = [tgt];
    updateDataMain(category);
  } else if (tgt != '') {
    category = [tgt];
    updateDataMain(category);
    updateDataOfficial(category);
  } else {
    // ブランクのとき、とりあえず起動
    updateDataMain(category);
    updateDataOfficial(category);
  }
}

// 最新情報の取得（通常シート用処理/バッチでは直呼び出し） --------------------------------------------------------------------------------------------------- //
function updateDataMain(targetList) {
  Logger.log("start update");
  setStartLog("最新情報取得（通常シート）");
  var dataSpread = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1O1hD48IWzzpGDwKOYlAraTCdO0EzfO1uRrKZ1Ut1luA/edit#gid=1015594370");
  var localSpread = SpreadsheetApp.getActiveSpreadsheet();
  var logRange = localSpread.getSheetByName("コーデ検索").getRange("D24");
  setScriptStatusStart(logRange);
  var updateList = localSpread.getSheetByName("コーデ検索").getRange(1, 22, 9, 2).getValues();
  var addDataCount = 0;
  
  try{
    for (var index in targetList) {
      Logger.log(targetList[index] + "を更新します。");
      setScriptStatusNow(logRange, targetList[index] + "を更新中です。");
      
      // 通常シート --------------------------------------------------------------------------------
      var dataSheet = dataSpread.getSheetByName(targetList[index]);
      var dataLR = dataSheet.getLastRow();
      var updData = dataSheet.getRange(2, 2, dataLR -1, 17).getValues();
      
      var localSheet = localSpread.getSheetByName(targetList[index]);
      var localLR = localSheet.getLastRow();
      var localData = localSheet.getRange(2, 2, localLR -1, 17).getValues();
      
      var newDataCount = 0;
      
      for (var d = 2; d < dataLR -1; d++) {
        var find = 0;
        lSearch : for (var l = 0; l < localData.length; l++) {
          if (returnHalfString(updData[d][1]) == returnHalfString(localData[l][1])) {
            // ナンバー一致
            for (var a = 2; a < localData[l].length; a++){
              localData[l][a] = updData[d][a];
            }
            find = 1;
            break lSearch;
          }
        }
        if (find == 0 && updData[d][1] != "") {
          Logger.log(updData[d][1] + "を追加します。");
          localData.push(updData[d]);
          newDataCount++;
          addDataCount++;
        }
      }
//      localLR = localLR + newDataCount;
      
      
      // 転記
      if (newDataCount!= 0) {
        // データがあるときのみ
        localSheet.insertRows(localSheet.getLastRow() +1, newDataCount);
        var afRow = localData.length;
        localSheet.getRange(2, 2, afRow, 17).setValues(localData);
        
        // ソート
        localLR = localSheet.getLastRow();
        localSheet.getRange(2, 1, localLR, 19).sort(2);
      }

    } // category loop
    
    setScriptStatusEnd(logRange);
    Logger.log("end update");
    setEndLog(2, "正常終了。" + addDataCount + "件のデータを追加しました。")
  } catch(e) {
    Logger.log(e);
    setEndLog(2, e);
    setScriptStatusError(logRange);
  }
}

// 最新情報の取得（公式変更/追加シート用処理/バッチでは直呼び出し） --------------------------------------------------------------------------------------------------- //
function updateDataOfficial(targetList) {
  Logger.log("start update");
  setStartLog("最新情報取得（公式変更/追加シート）");
  var dataSpread = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1O1hD48IWzzpGDwKOYlAraTCdO0EzfO1uRrKZ1Ut1luA/edit#gid=1015594370");
  var localSpread = SpreadsheetApp.getActiveSpreadsheet();
  var logRange = localSpread.getSheetByName("コーデ検索").getRange("D24");
  setScriptStatusStart(logRange);
  var updateList = localSpread.getSheetByName("コーデ検索").getRange(1, 22, 9, 2).getValues();
  var addDataCount = 0;
  
  try{
    for (var index in targetList) {
      Logger.log(targetList[index] + "を更新します。");
      setScriptStatusNow(logRange, targetList[index] + "を更新中です。");
      
      var localSheet = localSpread.getSheetByName(targetList[index]);
      var localLR = localSheet.getLastRow();
      var localData = localSheet.getRange(2, 2, localLR -1, 17).getValues();
      
      var newDataCount = 0;

      // 変更・追加シート --------------------------------------------------------------------------------
      Logger.log("公式変更/追加シートを検索しています。");
      var dataSheet = dataSpread.getSheetByName("公式変更/追加");
      var dataLR = dataSheet.getLastRow();
      var updData = dataSheet.getRange(3, 2, dataLR -1, 17).getValues();
      
      for (var d = 0; d < updData.length; d++) {
        var find = 0;
        var updCategory = updData[d][2];
        var updRare = updData[d][5]; // レア度ではなくSi or Goの文字列
        if (updCategory == "頭飾り" || updCategory == "首飾り" || updCategory == "腕飾り" || updCategory == "手持品" || updCategory == "特殊") {
          // アクセのときは置換
          updCategory = "アクセサリー";
        }
        if (updCategory == targetList[index] && 
            (updRare == "si" || updRare == "go")){
          // カテゴリー一致
          lSearch : for (var l = 0; l < localData.length; l++) {
            if (returnHalfString(updData[d][1]) == returnHalfString(localData[l][1])) {
              // ナンバー一致
              for (var a = 2; a < localData[l].length; a++){
                localData[l][a] = updData[d][a];
              }
              find = 1;
              break lSearch;
            }
          }
          if (find == 0 && updData[d][1] != "") {
            Logger.log(updData[d][1] + "を追加します。");
            localData.push(updData[d]);
            newDataCount++;
            addDataCount++;
          }
        }
      }
      
      // 転記
      if (newDataCount!= 0) {
        // データがあるときのみ
        localSheet.insertRows(localSheet.getLastRow() +1, newDataCount);
        var afRow = localData.length;
        localSheet.getRange(2, 2, afRow, 17).setValues(localData);
        
        // ソート
        localLR = localSheet.getLastRow();
        localSheet.getRange(2, 1, localLR, 19).sort(2);
      }

    } // category loop
    
    setScriptStatusEnd(logRange);
    Logger.log("end update");
    setEndLog(2, "正常終了。" + addDataCount + "件のデータを追加しました。")
  } catch(e) {
    Logger.log(e);
    setEndLog(2, e);
    setScriptStatusError(logRange);
  }
}