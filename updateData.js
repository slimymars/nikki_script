// 最新情報の取得 --------------------------------------------------------------------------------------------------- //
function updateData() {
  Logger.clear();
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
      
      var lst = new Array(localData.length);
      for (var i = 0; i < lst.length; i++) {
        lst[i] = i;
      }
      for (var d = 2; d < dataLR -1; d++) {
        Logger.log('check :' + updData[d][1]);
        var find = 0;
        lSearch : for (var idx = 0; idx < lst.length; idx++) {
          var l = lst[idx];
          if (returnHalfString(updData[d][1]) == returnHalfString(localData[l][1])) {
            // ナンバー一致
            for (var a = 2; a < localData[l].length; a++){
              localData[l][a] = updData[d][a];
            }
            lst.splice(idx, 1);
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

/// add by slimymars --------------------------------------------------------------------------------------------
function getHasList(targetName){
  var localSpread = SpreadsheetApp.getActiveSpreadsheet();
  var localSheet = localSpread.getSheetByName(targetName);
  var localLR = localSheet.getLastRow();
  var localData = localSheet.getRange(2, 1, localLR -1, 2).getValues();

  var result = [];
  for (var i = 0; i < localData.length; i++) {
    if (localData[i][0] !== "") {
      result.push({no: localData[i][1], has: localData[i][0]});
    }
  }
  return result;
}

function getCategorySheetData(targetName){
  Logger.log("データ取得: %s シート", targetName);

  var dataSpread = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1O1hD48IWzzpGDwKOYlAraTCdO0EzfO1uRrKZ1Ut1luA/edit#gid=1015594370");
  var dataSheet = dataSpread.getSheetByName(targetName);
  var dataLR = dataSheet.getLastRow();
  var updData = dataSheet.getRange(2, 2, dataLR -1, 17).getValues();

  var result = [];

  for (var i = 0; i < updData.length; i++) {
    var d = updData[i];
    result.push({
      no: d[0],
      values: d,
      has: ""
    });
  }

  return result;
}

officialSheetDataCache = [];

function makeOfficialSheetDataCache() {
  Logger.log("公式変更/追加シートデータ取得");
  var dataSpread = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1O1hD48IWzzpGDwKOYlAraTCdO0EzfO1uRrKZ1Ut1luA/edit#gid=1015594370");
  var dataSheet = dataSpread.getSheetByName("公式変更/追加");
  var dataLR = dataSheet.getLastRow();
  var updData = dataSheet.getRange(3, 2, dataLR -1, 17).getValues();

  for (var i = 0; i < updData.length; i++) {
    var d = updData[i];
    var category = d[2];
    var updRare = d[5]; // レア度ではなくSi or Goの文字列
    if (category == "頭飾り" || category == "首飾り" || category == "腕飾り" || category == "手持品" || category == "特殊") {
      // アクセのときは置換
      category = "アクセサリー";
    }
    if (updRare == "si" || updRare == "go") {
      officialSheetDataCache.push({
        no: d[0],
        category: category,
        values: d,
        has: ""
      })
    }
  }
}

function getOfficialSheetData(targetName) {
  if (officialSheetDataCache.length === 0){
    makeOfficialSheetDataCache();
  }
  Logger.log("取得数: %s", officialSheetDataCache.length);
  var result = officialSheetDataCache.filter(function (value) {
    return value.category == targetName;
  });
  return result;
}

function margeOfficialToCategorySheet(officialDataList, categoryDataList){
  var officialIncludeNoList = officialDataList.map(function (value) { return value.no });
  var newList = categoryDataList.filter(function (value) {
    return !officialIncludeNoList.includes(value.no);
  });

  return newList.concat(officialDataList);
}

function writeData(targetName, values){
  var localSpread = SpreadsheetApp.getActiveSpreadsheet();
  var toSheet = localSpread.getSheetByName(targetName);
  var endR = toSheet.getLastRow();
  toSheet.getRange(2, 1, endR -1, 17).clear({contentsOnly: true});
  if (values.length > endR - 1){
    toSheet.insertRows(endR+1, values.length-(endR-1));
  }
  var dataList = values.map(function (value) { return [value.has].concat(value.values); });
  toSheet.getRange(2, 1, dataList.length+1, 18).setValues(dataList);
}

function setHasData(hasList, values){
  var hash = values.reduce(function (ac, v) {
    ac[v.no] = v;
  }, {});
  hasList.forEach(function (d) {
    hash[d.no].has = d.has;
  });
}

function dataUpdate(targetList){
  targetList.forEach(function (target) {
    Logger.log("所持リスト取得: %s", target);
    setStartLog("所持リスト取得");
    var nowHasList = getHasList(target);
    Logger.log("カテゴリシート取得");
    var categolyData = getCategorySheetData(target);
    Logger.log("公式追加/更新シート取得");
    var officialData = getOfficialSheetData(target);
    Logger.log("データマージ");
    var data = margeOfficialToCategorySheet(officialData, categolyData);
    setHasData(nowHasList, data);
    Logger.log("書き出し");
    writeData(target, data);
  });
}

function updateTest() {
  dataUpdate(["アクセサリー"]);
}
