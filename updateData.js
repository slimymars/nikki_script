// 最新情報の取得 --------------------------------------------------------------------------------------------------- //
// アクセサリー種定義
var Accessories = ["頭飾り", "首飾り", "腕飾り", "手持品", "特殊", "腰飾り", "耳飾り"];

function updateData() {
  Logger.clear();
  var localSpread = SpreadsheetApp.getActiveSpreadsheet();
  var category = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "アクセサリー", "メイク"];
  var tgt = localSpread.getSheetByName("コーデ検索").getRange("D23").getValue();
  if (tgt == "アクセ以外") {
    category = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "メイク"];
    dataUpdate(category);
  } else if (tgt != '') {
    category = [tgt];
    dataUpdate(category);
  } else {
    // ブランクのとき、とりあえず起動
    dataUpdate(category);
  }
}

function getLocalSheet(target) {
  if (getLocalSheet.memo && getLocalSheet.memo[target]) {return getLocalSheet.memo[target]}
  if (!getLocalSheet.memo) {getLocalSheet.memo = {}}
  getLocalSheet.memo[target] = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(target);
  return getLocalSheet.memo[target];
}

function getDataSheet(target) {
  if (getDataSheet.memo && getDataSheet.memo[target]) { return getDataSheet.memo[target]; }
  if (!getDataSheet.spld) { 
    getDataSheet.spld = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1O1hD48IWzzpGDwKOYlAraTCdO0EzfO1uRrKZ1Ut1luA/edit#gid=1015594370");
    getDataSheet.memo = {};
  }
  getDataSheet.memo[target] = getDataSheet.spld.getSheetByName(target);
  return getDataSheet.memo[target];
}

function getHasList(targetName){
  var localSheet = getLocalSheet(targetName);
  var localLR = localSheet.getLastRow();
  var localData = localSheet.getRange(2, 1, localLR -1, 19).getValues();

  var result = {};
  for (var i = 0; i < localData.length; i++) {
    if (localData[i][0] !== "" || localData[i][18] !== "") {
      result[localData[i][1]] =  {has: localData[i][0], ep: localData[i][18], values: localData[i].slice(1, 18)};
    }
  }
  return result;
}

function getCategorySheetData(targetName){
  Logger.log("データ取得: %s シート", targetName);

  var dataSheet = getDataSheet(targetName);
  var dataLR = dataSheet.getLastRow();
  var updData = dataSheet.getRange(2, 2, dataLR -1, 17).getValues();

  var result = {};

  for (var i = 0; i < updData.length; i++) {
    var d = updData[i];
    result[d[0]] = { values: d, has: ""};
  }

  return result;
}

function getOfficialSheetDataList() {
  if (getOfficialSheetDataList.memo) { return getOfficialSheetDataList.memo; }
  getOfficialSheetDataList.memo = [];
  Logger.log("公式変更/追加シートデータ取得");
  var dataSheet = getDataSheet("公式変更/追加");
  var dataLR = dataSheet.getLastRow();
  var updData = dataSheet.getRange(3, 2, dataLR -1, 17).getValues();

  for (var i = 0; i < updData.length; i++) {
    var d = updData[i];
    var category = d[2];
    var updRare = d[5]; // レア度ではなくSi or Goの文字列
    if (Accessories.indexOf(category) >= 0) {
      // アクセのときは置換
      category = "アクセサリー";
    }
    if (updRare == "si" || updRare == "go") {
      getOfficialSheetDataList.memo.push({
        no: d[0],
        category: category,
        values: d,
        has: ""
      })
    }
  }
  return getOfficialSheetDataList.memo;
}

function getOfficialSheetData(targetName) {
  var officialSheetData = getOfficialSheetDataList();
  Logger.log("取得数: %s", officialSheetData.length);
  return officialSheetData.reduce(function (ac, value) {
    if (value.category != targetName) {
      return ac;
    }
    ac[value.no] = {values: value.values, has: value.has};
    return ac;
  }, {});
}

function margeOfficialToCategorySheet(officialDataList, categoryDataList){
  for (var no in officialDataList) if (officialDataList.hasOwnProperty(no)) {
    categoryDataList[no] = officialDataList[no];
  }
}

function writeData(targetName, values){
  var toSheet = getLocalSheet(targetName);
  var endR = toSheet.getLastRow();
  toSheet.getRange(2, 1, endR -1, 19).clear({contentsOnly: true});
  var dataList = [];
  for (var key in values) if (values.hasOwnProperty(key)) {
    var value = values[key];
    var v = [value.has].concat(value.values);
    if (value.hasOwnProperty("ep")){
      v.push(value.ep);
    } else {
      v.push("");
    }
    dataList.push(v);
  }
  if (dataList.length > endR - 1){
    toSheet.insertRowsAfter(endR, dataList.length-(endR-1));
  }
  toSheet.getRange(2, 1, dataList.length, 19).setValues(dataList);
  // ソート
  var localLR = toSheet.getLastRow();
  toSheet.getRange(2, 1, localLR - 1, 19).sort(2);
}

function setHasData(hasList, values){
  for (var no in hasList) if (hasList.hasOwnProperty(no)) {
    if (values.hasOwnProperty(no)) {
      values[no].has = hasList[no].has;
      values[no].ep = hasList[no].ep;
    } else {
      values[no] = hasList[no];
    }
  }
}

function dataUpdate(targetList){
  Logger.clear();
  Logger.log("start update");
  var logRange = getLocalSheet("コーデ検索").getRange("D24");
  setScriptStatusStart(logRange);
  targetList.forEach(function (target) {
    setStartLog("更新: " + target);
    setScriptStatusNow(logRange, target + "を更新中です。");
    try {
      Logger.log("所持リスト取得: %s", target);
      var nowHasList = getHasList(target);
      Logger.log("カテゴリシート取得");
      var categolyData = getCategorySheetData(target);
      Logger.log("公式追加/更新シート取得");
      var officialData = getOfficialSheetData(target);
      Logger.log("データマージ");
      margeOfficialToCategorySheet(officialData, categolyData);
      setHasData(nowHasList, categolyData);
      Logger.log("書き出し");
      writeData(target, categolyData);
      setEndLog(2, "正常終了: " + target);
    } catch(e) {
      Logger.log(e);
      setEndLog(2, e);
      setScriptStatusError(logRange);
    }
  });
  setScriptStatusEnd(logRange);
  Logger.log("end update");
}

function updateTest() {
  dataUpdate(["アクセサリー"]);
}
