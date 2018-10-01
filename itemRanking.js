function checkRanking() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("アイテムランキング");
  var category = sheet.getRange("D2").getValue();
  var subCategory = sheet.getRange("G2").getValue();
  var item = sheet.getRange("D3").getValue();
  var attr = sheet.getRange("B5:F6").getValues();
  
  // 必須情報のチェック
  if (category == '') {
    // カテゴリブランク
    SpreadsheetApp.getUi().alert("カテゴリを入力してください。");
    return;
  } else if (subCategory == ""){
    // サブカテゴリがブランク
    if (spreadsheet.getSheetByName("参照用シート").getRange("N20").getValue() != "-") {
      // サブカテゴリ必須
      SpreadsheetApp.getUi().alert("サブカテゴリを入力してください。");
      return;
    } else {
      // それ以外はサブカテゴリをセットして続ける
      sheet.getRange("G2").setValue("-");
    }
  } 
  
  // 属性チェック（タグを除く）
  for (var r = 0; r < 2; r++) {
    for (var c = 0; c < 5; c++) {
      if (attr[r][c] == '') {
        // 属性がブランク
        SpreadsheetApp.getUi().alert("各属性を入力してください。");
        return;
      }
    }
  }
  
  // アイテム名
  if (item == '') {
    item == "カスタムアイテム";
    sheet.getRange("D3").setValue("カスタムアイテム");
  }
  
  // 計算処理
  rankingCal(sheet, category, subCategory, item, attr);
}

// 計算処理
function rankingCal(sheet, category, subCategory, item, attr) {
  var log = sheet.getRange("D8");
  setScriptStatusStart(log);
  var ret = [];
  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var calcSheet = spreadSheet.getSheetByName("ランキング計算");
  var evalSheet = spreadSheet.getSheetByName("評価値一覧");
  var tmpSheet = spreadSheet.getSheetByName("参照用シート");
  
  var subCategoryVal = "";
  if (subCategory != "-") {
    subCategoryVal = tmpSheet.getRange("N19").getValue();
  }
  
  // 前回実行データの消去
  if (sheet.getLastRow()-9 != 0){
    sheet.deleteRows(10, sheet.getLastRow()-9);
  }

  // ステージ一覧
  var stageList = evalSheet.getRange("A2:AD" + evalSheet.getLastRow()).getValues();
  // NGステージ一覧
  var ngStage = tmpSheet.getRange("P2").getValue();
  var ngStageList = [];
  if (ngStage != ''){
    ngStageList = tmpSheet.getRange(tmpSheet.getRange("P2").getValue()).getValues();
  }
  // カテゴリ
  var categorySheet = spreadSheet.getSheetByName(category);
  var categoryData = categorySheet.getRange("A2:R" + categorySheet.getLastRow()).getValues();
  if (item == "カスタムアイテム") {
    categoryData.push(["◯", "", item, "" , "", "", attr[0][0], attr[1][0], attr[0][1], attr[1][1], attr[0][2], attr[1][2], attr[0][3], attr[1][3], attr[0][3], attr[1][3], attr[0][4], attr[1][4], attr[0][5], attr[1][6]]);
  }
  Logger.log(ngStageList);
  
  for (var w = 0; w < 3; w++) {
    // 横展開ループ
    var baseW = (10 * w);

    for (var s = 0; s < stageList.length; s++) {
      // ステージループ
      var totalscore = [];
      
      
      if (stageList[s][baseW] == '') {
        // ステージ名がないときは終了
        break;
      }
      if (ngStage != '') {
        if (ngStageList[0].indexOf(stageList[s][baseW]) != -1) {
          // NGステージのときはパス
          Logger.log(stageList[s][baseW] + "はNGステージです。");
          s++; // 情報を2行ごとに持っているため1加算する
          continue;
        }
      }
      
      Logger.log(stageList[s][baseW] + "における" + category + "のランキング計算を始めます。");
      var tag = [stageList[s][baseW + 6],stageList[s][baseW + 7]];
      var tagWeight = [stageList[s+1][baseW + 6],stageList[s+1][baseW + 7]];
      var value = [stageList[s][baseW + 1],stageList[s][baseW + 2],stageList[s][baseW + 3],stageList[s][baseW + 4],stageList[s][baseW + 5]];
      var weight = [stageList[s+1][baseW + 1],stageList[s+1][baseW + 2],stageList[s+1][baseW + 3],stageList[s+1][baseW + 4],stageList[s+1][baseW + 5]];
      
      for (var i = 0; i < categoryData.length; i++) {
        // カテゴリ内アイテムループ
        var score = [];
        var tagScore = [];
        if ((categoryData[i][0] != "◯" || (subCategory != "-" && subCategoryVal != categoryData[i][4])) && categoryData[i][2] != "カスタムアイテム"){
          // 除外
          // ①未所持
          // ②小分類指定有 かつ 小分類が異なる
          continue;
        }
        
        // タグのスコア
        for (var j = 0;j < 2;j++) {
          if (tag[0][j * 2] == categoryData[i][16] || tag[0][j * 2] == categoryData[i][17]) {
            switch(tagWeight[0][j * 2]) {
              case "SSS":
                tagScore[j] = 3000 * tagWeight[0][j * 2 + 1];
                break;
              case "SS":
                tagScore[j] = 2612.7 * tagWeight[0][j * 2 + 1];
                break;
              case "S":
                tagScore[j] = 2089.35 * tagWeight[0][j * 2 + 1];
                break;
              case "A":
                tagScore[j] = 1690.65 * tagWeight[0][j * 2 + 1];
                break;
              case "B":
                tagScore[j] = 1309.8 * tagWeight[0][j * 2 + 1];
                break;
              case "C":
                tagScore[j] = 817.5 * tagWeight[0][j * 2 + 1];
                break;
              default:
                tagScore[j] = 0;
                break;
            }
          } else{
            tagScore[j] = 0;
          }
        }
        
        for (var j = 0;j < 5;j++) {
          if (categoryData[i][j * 2 + 6] == value[j]) {
            score[j] = categoryData[i][j * 2 + 7];
          } else{
            score[j] = 0;
          }
        }
        for (var j = 0;j < 5;j++) {
          switch(score[j]) {
            case "SSS":
              score[j] = 3000;
              break;
            case "SS":
              score[j] = 2612.7;
              break;
            case "S":
              score[j] = 2089.35;
              break;
            case "A":
              score[j] = 1690.65;
              break;
            case "B":
              score[j] = 1309.8;
              break;
            case "C":
              score[j] = 817.5;
              break;
            default:
              score[j] = 0;
              break;
          }
          score[j] = (score[j] + tagScore[0] + tagScore[1]) * weight[j];
        }
        
        totalscore[i] = [categoryData[i][2] ,(score[0] + score[1] + score[2] + score[3] + score[4])];
      }
      
      // 結果のソート
      var result = totalscore.sort(function(a,b){return(b[1] - a[1]);});
      var rank = sheet.getRange("H3").getValue();
      for (var r = 0; r < result.length; r++) {
        if (Number(rank) <= r) {
          // 指定位以下は除外
          break;
        }
        if(result[r] != null) {
          if (result[r][0] == item) {
            // 名称一致
            var retScore = result[r][1].toFixed(0);
            if (retScore == 0) {
              // スコア0のときNG
              break;
            } else {
              var str = [];
              str = [stageList[s][baseW], (r+1) + "位", retScore];
              Logger.log(str);
              ret.push(str);
            }
          }
        }
      }
      
      s++; // 情報を2行ごとに持っているため1加算する
    }
  }
  
  // 反映
  if (ret.length != 0) {
    // 順位順にソート
    ret.sort(function(a,b){
      // 順位の数値を切り出してソート
      return(a[1].substring(0, a[1].length -1) - b[1].substring(0, b[1].length -1));
    });
    //  Logger.log(ret);
    
    sheet.insertRowsAfter(9, ret.length);
    for (var i = 0; i < ret.length; i++) {
      // 結合のためのブランクを入れる
      ret[i] = [ret[i][1], ret[i][2], "", ret[i][0]];
    }
    // 値のセット
    sheet.getRange(10, 2, ret.length, 4).setValues(ret).setBorder(true, true, true, true, true, true).setBackground("white");
    sheet.getRange(10, 3, ret.length, 2).mergeAcross();
    sheet.getRange(10, 5, ret.length, 3).mergeAcross();
  } else {
    // データなし
    sheet.insertRowsAfter(9, 1);
    sheet.getRange(10, 2).setValue("該当データなし");
    sheet.getRange(10, 2, 1, 4).setBackground("white");
  }
  setScriptStatusEnd(log);
}

// アイテムを選択したときの処理
function selectItem() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("アイテムランキング");
  var category = sheet.getRange("D2").getValue();
  var item = sheet.getRange("D3").getValue();
  Logger.log(item + "の評価を反映します。");
  
  if (category == ''){
    // カテゴリがブランクのとき
    // カテゴリ特定処理
    var categoryList = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "メイク", "頭飾り", "耳飾り", "首飾り", "腕飾り", "手持品", "腰飾り", "特殊"];
    for (var i = 0; i < categoryList.length; i++) {
      Logger.log(item + "を" + categoryList[i] + "で検索します。");
      var categorySheet = spreadsheet.getSheetByName(categoryList[i]);
      var categoryData = categorySheet.getRange("C2:C" + categorySheet.getLastRow()).getValues();
      for (var n=0; n < categoryData.length; n++) {
        if (categoryData[n][0] == item) {
          category = categoryList[i];
          sheet.getRange("D2").setValue(category);
          Logger.log(category + "で" + item + "を発見しました。");
          break;
        }
      }
      if (category != ''){
        break;
      }
    }
  }
  
  if (category != ''){
    // カテゴリがあるとき、評価反映
    var dataSheet = spreadsheet.getSheetByName(category);
    var dataList = dataSheet.getRange("A2:R" + dataSheet.getLastRow()).getValues();
    for (var i = 0; i < dataList.length; i++){
      if (dataList[i][2] == item){
        // 名称一致
        var attr = [dataList[i][6], dataList[i][8], dataList[i][10], dataList[i][12], dataList[i][14]];
        var attrVal = [dataList[i][7], dataList[i][9], dataList[i][11], dataList[i][13], dataList[i][15]];
        sheet.getRange("B5:F5").setValues([attr]);
        sheet.getRange("B6:F6").setValues([attrVal]);
        sheet.getRange("G5").setValue(dataList[i][16]);
        sheet.getRange("I5").setValue(dataList[i][17]);
        break;
      }
    }
  }
}

// カテゴリ入力
function selectCategory(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("アイテムランキング");
  var category = sheet.getRange("D2").getValue();
  if (category != '') {
    sheet.getRange("G2").clearContent();
  }
}

// 属性カスタム
function changeAttr() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("アイテムランキング");
  sheet.getRange("D3").setValue("カスタムアイテム");
}