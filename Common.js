// Spreadsheetが開かれたときに自動的に呼び出される。（スマホでは呼び出されない）
function onOpen()
{
  Logger.log("onOpen start");

  // カスタムメニューの追加
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ニキメニュー')
    .addItem('計算（calc）', 'search')
    .addSeparator()
    .addItem('データ更新（取得）', 'updateData')
    .addItem('NGデータ更新', 'getNg')
    .addSeparator()
    .addItem('コロッセオ前半更新', 'updateOfColosseum1')
    .addItem('コロッセオ後半更新', 'updateOfColosseum2')
  .addItem('ギルド1章更新', 'updateOfGuild1')
  .addItem('ギルド2章更新', 'updateOfGuild2')
  .addItem('ギルド3章更新', 'updateOfGuild3')
  .addItem('ギルド4章更新', 'updateOfGuild4')
  .addItem('ギルド5章更新', 'updateOfGuild5')
  .addItem('ギルド6章更新', 'updateOfGuild6')
  .addItem('ギルド7章更新', 'updateOfGuild7')
    .addSeparator()
    .addItem('所持データ更新', 'koreMotteru')
  .addToUi();

  Logger.log("onOpen end");
}

// セルが編集されたときに自動的に呼び出される。 （どのシートが編集されても呼び出されるので注意）
function onEdit(e)
{
  Logger.log("onEdit start");

  var spreadsheet = e.source; // 編集されたスプレットシート
  var sheet = spreadsheet.getActiveSheet(); // 編集されたシート
  var range = e.range;  // 編集された範囲

  var sheetName = sheet.getName(); // 編集されたシート名
  var rangeName = range.getA1Notation(); // 編集された範囲のA1形式の文字列

  // onEditは、すべてのシートの編集時に呼び出されるので、シートの特定を先に行う。
  if (sheetName == "コーデ検索") {
    // コーデ評価セルリスト
    var codeEvalCellList = ["B2", "C2", "D2", "E2", "F2", "G2", "H2", "I2"];
    
    if (rangeName == "D7") {
      // コーデ名が編集されたときの処理
      renewEval();
    } else if (rangeName == "H18") {
      // スマホ用calc計算時の処理
      Logger.log("onEdit スマホ用calc計算時の処理 start");
      search("1");
      Logger.log("onEdit スマホ用calc計算時の処理 end");
    } else if (rangeName == "C25") {
      // スマホ用最新情報取得時の処理
      Logger.log("onEdit スマホ用最新情報取得時の処理 start");
      Logger.log("openByUrlは動かせないの～");
      //updateData();
      setScriptStatusNow(sheet.getRange("D24"), "動きませんて～");
      Logger.log("onEdit スマホ用最新情報取得時の処理 end");
    } else if (codeEvalCellList.indexOf(rangeName) != -1){
      // コーデ評価セルリストの編集時の処理
      Logger.log("onEdit コーデ評価セルリストの編集時の処理 start");
      if(sheet.getRange("D7").getValue() != "カスタム") {
        // カスタムに変更
        var customVal = [1,1,1,1,1,'','','',''];
        sheet.getRange("B3:J3").setValues([customVal]);
        sheet.getRange("B7").setValue("-");
        sheet.getRange("D7").setValue("カスタム");
      }
      checkCode();
      Logger.log("onEdit コーデ評価セルリストの編集時の処理 end");
    }
  }　else if (sheetName == "所持追加用シート" && rangeName == "B1") {
    Logger.log("onEdit スマホ用所持追加時の処理 start");
    koreMotteru();
    Logger.log("onEdit スマホ用所持追加時の処理 end");
  } else if (sheetName == "アイテムランキング"){
    // 属性セルリスト
    var attrCellList = ["B5", "C5", "D5", "E5", "F5", "G5", "I5", "B6", "C6", "D6", "E6", "F6"];
    
    if (rangeName == "D3") {
      // アイテム入力
      selectItem();
    } else if (rangeName == "D2") {
      // カテゴリ
      selectCategory();
    } else if (attrCellList.indexOf(rangeName) != -1) {
      // 属性セル編集
      changeAttr();
    }
  }
  Logger.log("onEdit end");
}

// 評価値の反映（ステージ選択時の処理）
function renewEval(){
  Logger.log("onEdit ステージ選択時の処理 start");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var evaluationSheet = spreadsheet.getSheetByName("評価値一覧");
  var sheet = spreadsheet.getSheetByName("コーデ検索");
  
  var lastRow = evaluationSheet.getLastRow();
  var stage = sheet.getRange("D7").getValue();
  var valueData = evaluationSheet.getRange(2,1,lastRow - 1,30).getValues();
  
  for (var j = 0;j < 3;j++) {
    for (var i = 0;i < valueData.length;i++) {
      if (stage == valueData[i][j * 10]) {
        var value = evaluationSheet.getRange(i + 2,j * 10 + 2,2,9).getValues();
        break;
      }
    }
  }
  sheet.getRange("B2:J3").setValues(value);
  sheet.getRange("H18").setValue(stage);
  sheet.getRange("B4").clearContent();
  Logger.log("onEdit ステージ選択時の処理 end");
}
