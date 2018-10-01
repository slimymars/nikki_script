/* ユーザ向け使い方メモ --------------------------------------------------------------------

◆ 各スクリプトの内容
　Main：コーデ探索メイン処理
　Common：共通処理。メニュー登録、セル編集時処理、ステージ選択時の処理
　updateData：最新情報取得処理
　koremotteru：所持情報更新処理
　getNG：NGコーデ取得処理（Gamerchのwikiより取得）
　codeSace：コーデ保存処理
　Batch：バッチ処理（日次・月次）
　pdateCoordination：コーデ一括更新（バッチから呼び出される）
　setLog：ログ周辺の処理
　returnRep：変換処理

◆ 手動更新が必要なところ
　【ステージ追加】
 　　評価値一覧シートに対象の評価値を追加。Nikki共有シートからコピペでOK。

　【ギルドステージが追加されたとき】
 　　Batchの最下行あたりに既存処理をまねて「ギルドN章のコーデを更新」処理を追加する。
 　　functionの連番、capterの値をギルドの章に合わせて変更。
    また、deleteTriggerに対し、自身のfunction名を設定する
    例）8章追加
    　function名：updateOfGuild8
      capterの値：8
      deleteTriggerに渡す値：updateOfGuild8

---------------------------------------------------------------------------------------- */


// 最適コーデ計算処理
function search() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("コーデ検索");
  var rangeLog = sheet.getRange("H16");
  
  setScriptStatusStart(rangeLog);
  var msg = raiseAllCategory(onAllCategoryCalc, null);
  setScriptStatusEnd(rangeLog);

  // 最終更新日
  var updSheet = spreadsheet.getSheetByName("コーデ更新日時");
  var updLR = updSheet.getLastRow();
  var updData = updSheet.getRange(1, 1, updLR, 6).getValues();
  var tgtG = sheet.getRange("B7").getValue();
  var tgt = sheet.getRange("D7").getValue();
  for (var i = 0; i < 6; i++) {
    if (updData[0][i] == tgtG) {
      // 区分が一致したとき
      for (var l = 0; l < updLR; l++ ){
        if (updData[l][i] == tgt) {
          // 名称が一致したとき
          updData[l][i + 1] = getDateTimeJSTString(new Date())
          Logger.log(tgt + "の最終更新日時を更新。");
          break;
        }
      }
      break;
    }
  }
  updSheet.getRange(1, 1, updLR, 6).setValues(updData);
  
  // コーデ転記
  Utilities.sleep(5000);
  codeSace();
  
  // メッセージ邪魔なので消す
  //SpreadsheetApp.getUi().alert(msg);
}

// すべてのカテゴリに対して、ループ処理を行う際の共通処理
// 処理の中身は、別のfunctionで指定する。
function raiseAllCategory(callback, params) {
  Logger.log("raiseAllCategory start: " + callback.name);
  var sttime = new Date();

  var category = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "アクセサリー", "メイク"];
  for (var index in category) {
    // コールバック（引数callbackで指定したfunctionを呼び出す）
    // コールバックに指定するfunctionは、以下の二つの引数を受け取るようにすること。
    // 第一引数はカテゴリの文字列を受け取り、第二引数は、パラメータとして自由に使う。
    // パラメータは、呼び出し元と呼び出し先で型を一致させればよい。第二引数だけで複数パラメータも渡せる。
    // ここでは、わかりやすいように、functionの名前は、「onAllCategory～」とする。
    callback(category[index], params);
  }

  var edtime = new Date();
  var time = (edtime - sttime) / 1000;
  var msg = "完了！\n所要時間：" + time + "s";

  Logger.log("raiseAllCategory end: " + callback.name);
  return msg;
}


// コールバックによってカテゴリごとに呼び出される。（計算処理用）
// 第二引数のparamsは、未使用NULL
function onAllCategoryCalc(category, params) {
  calc(category);
}


// カテゴリ別の最適コーデ計算処理 --------------------------------------------------------------------------------------------------- //
function calc(category) {
  Logger.log("calc start: " + category);

  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("コーデ検索");
  var ref = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(category); 
  var lr = ref.getLastRow();
  var value = mySheet.getRange("B2:F2").getValues();
  var weight = mySheet.getRange("B3:F3").getValues();
  var tag = mySheet.getRange("G2:J2").getValues();
  var tagWeight = mySheet.getRange("G3:J3").getValues();
  var data = ref.getRange(2,3,lr - 1,16).getValues(); // 2, 7, lr -1, 12
  var totalscore = [];
  var delSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("参照用シート"); 
  var delList = delSheet.getRange(delSheet.getRange("J2").getValue()).getValues();
  
  for (var i = 0;i < data.length;i++) {
    var score = [];
    var tagScore = [];
    
    for (var j = 0;j < 2;j++) {
      if (tag[0][j * 2] == data[i][14] || tag[0][j * 2] == data[i][15]) { //[10],[11]
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
      }
      else{
        tagScore[j] = 0;
      }
    }
    
    for (var j = 0;j < 5;j++) {//j=0
      if (data[i][j * 2 + 4] == value[0][j]) {
        score[j] = data[i][j * 2 + 5];
        }
      else{
        score[j] = 0;
      }
    }
    for (var j = 0;j < 5;j++) { //j=0
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
      score[j] = (score[j] + tagScore[0] + tagScore[1]) * weight[0][j];
    }
    
    // 除外
    for (var d =0; d < delList.length; d++) {
      if (delList[d] == data[i][0]) {
        score[0] = 0;
        score[1] = 0;
        score[2] = 0;
        score[3] = 0;
        score[4] = 0;
        break;
      }
    }
    totalscore[i] = score[0] + score[1] + score[2] + score[3] + score[4];
  }
  var objtotalscore = [];
  for (var i = 0;i < totalscore.length;i++) {
    objtotalscore.push([totalscore[i]]);
  }
  ref.getRange(2,19,lr - 1,1).setValues(objtotalscore);

  Logger.log("calc end: " + category);
}

