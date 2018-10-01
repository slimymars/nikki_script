// NGリスト作成
function getNgList() {
  var out = [];
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("評価値一覧");
  var rowCount = 0;
  
  // NGリスト
  sheet.getRange("AE:AF").clearContent();
  out = getNgFromNormalStage(); // 通常ステージ
  out = out.concat(getNgFromGuildStage()); // ギルド
  rowCount = out.length +1;
  sheet.getRange("AE2:AF" + rowCount).setValues(out).setHorizontalAlignment("left").setNumberFormat('@');
  
  // 範囲NG
  sheet.getRange("AG:AH").clearContent();
  out = getCaution(); // ～以外NG
  rowCount = out.length +1;
  sheet.getRange("AG2:AH" + rowCount).setValues(out).setHorizontalAlignment("left").setNumberFormat('@');
}

// 通常
// ※「～以外すべてNG」は非対応
function getNgFromNormalStage() {
  var out = [];

  var url = "https://miraclenikki.gamerch.com/NG%E4%B8%80%E8%A6%A7";
  var response = UrlFetchApp.fetch(url);
  
  // 正規表現
  var allReg = /<.*>.*<\/.*>/gi;
  var stageReg = /<div .*><h3><a href.*>(.*)<\/a><\/h3>/gi;
  var itemReg = /<tr><td .*data-col="1"><a href.*>(.*)<\/a><\/td><\/tr>/gi;
  var getNameReg = /.*>(.*)<\/a>.*/i;
  var loopEndReg = /.*コメント.*/; // ループ終了向け

  // 探索
  var stageMatch = stageReg.exec(response.getContentText());
  var t;
  var stage;
  while ((t = allReg.exec(response.getContentText())) !== null ){

    if (stageReg.test(t)) {
      // ステージ名
      stage = t[0].match(getNameReg)[1];
      //Logger.log(stage + "のNG設定を開始します。 -----------------------");
    }
    if (itemReg.test(t)) {
      // アイテム
      var item = [];
      item[0] = stage;
      item[1] = t[0].match(getNameReg)[1];
      out.push(item);
      Logger.log(stage + "に" + item[1] + "のNGを設定します。");
    }
    
    if (loopEndReg.test(t)) {
        break; // loopEndRegと一致するHTMLを取得したら終了する
    }
  }
  return out;
}

// ギルド
function getNgFromGuildStage () {
  var out = [];
  
  var url = "https://miraclenikki.gamerch.com/%E4%BE%9D%E9%A0%BC";
  var response = UrlFetchApp.fetch(url);
  
  // 正規表現
  var allReg = /.*<.*>.*<\/.*>/gi;
  var stageReg = /<span.*name="(.*)"<\/a><\/span>/gi;
  var stageNameReg = /name="(.*)"/i;
  var itemReg = /<tr><th .*>NGコーデ<\/th>.*<a href.*>(.*)<\/a>/gi;
  var itemReg2 = /^【/i;
  var getNameReg = /.*>(.*)<\/a>.*/i;
  var loopEndReg = /.*コメント.*/; // ループ終了向け
  
  // 探索
  var stageMatch = stageReg.exec(response.getContentText());
  var t;
  var stage;
  while ((t = allReg.exec(response.getContentText())) !== null ){

    if (stageNameReg.test(t)) {
      // ステージ名
      stage = t[0].match(stageNameReg)[1] + "G";
      //Logger.log(stage + "のNG設定を開始します。 -----------------------");
    }
    if (itemReg.test(t)) {
      // アイテム
      var item = [];
      item[0] = stage;
      item[1] = t[0].match(getNameReg)[1];
      out.push(item);
      Logger.log(stage + "に" + item[1] + "のNGを設定します。");
    }
    if (itemReg2.test(t)) {
      // アイテム（2行目以降）
      var item = [];
      item[0] = stage;
      item[1] = t[0].match(getNameReg)[1];
      out.push(item);
      Logger.log(stage + "に" + item[1] + "のNGを設定します。");
    }
    
    if (loopEndReg.test(t)) {
        break; // loopEndRegと一致するHTMLを取得したら終了する
    }
  }
  return out;
}

// ～以外NGの有無
function getCaution() {
  var out = [];

  var url = "https://miraclenikki.gamerch.com/NG%E4%B8%80%E8%A6%A7";
  var response = UrlFetchApp.fetch(url);
  
  // 正規表現
  var allReg = /<.*>.*<\/.*>/gi;
  var stageReg = /<div .*><h3><a href.*>(.*)<\/a><\/h3>/gi;
  var stageNameReg = /.*>(.*)<\/a>.*/i;
  var checkReg = /すべてNG/i;
  var getNameReg = /.*color:#ff0000;(.*)すべてNG/i;
  
  // 探索
  var stageMatch = stageReg.exec(response.getContentText());
  var t;
  var stage;
  while ((t = allReg.exec(response.getContentText())) !== null ){
    
    if (stageReg.test(t)) {
      // ステージ名
      stage = t[0].match(stageNameReg)[1];
      //Logger.log(stage + "のNG設定を開始します。 -----------------------");
    }
    if (checkReg.test(t)) {
      // 注意文
      var att = t[0];
      var attReg = />([^<]+)</gi;
      if (attReg.test(att)) {
        var rep = att.match(attReg);
        att = "";
        for (var index in rep){
          att = att + rep[index].replace("<", "").replace(">", "");
        }
        var item = [];
        item[0] = stage;
        item[1] = att;
        out.push(item);
        Logger.log(stage + ":" + att);
      }
    }
  }
  return out;
}

