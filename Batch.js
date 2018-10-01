//日トリガー
function setTriggerOfDay(){
  var triggerDay = new Date();
  
  /* ----------------------------------------------------------------------------------
  * 最新情報取得　01:00 ～ 01:40　毎日実行
  * ---------------------------------------------------------------------------------- */
  
  // 前半
  triggerDay.setHours(1);
  triggerDay.setMinutes(0);
  ScriptApp.newTrigger("update1").timeBased().at(triggerDay).create();
  
  // 後半
  triggerDay.setMinutes(10);
  ScriptApp.newTrigger("update2").timeBased().at(triggerDay).create();
  
  // アクセ
  triggerDay.setMinutes(20);
  ScriptApp.newTrigger("updateAcc").timeBased().at(triggerDay).create();
  triggerDay.setMinutes(30);
  ScriptApp.newTrigger("updateAccOfficial").timeBased().at(triggerDay).create();
  
}

// 月トリガー
function setTriggerOfMonth() {
  var triggerDay = new Date();
  
  /* ----------------------------------------------------------------------------------
  * コーデ更新　02:00 ～　月1回実行
  * ---------------------------------------------------------------------------------- */
  
  // コロッセオ前半 02:00
  triggerDay.setHours(2);
  triggerDay.setMinutes(0);
  ScriptApp.newTrigger("updateOfColosseum1").timeBased().at(triggerDay).create();
  
  // コロッセオ後半 02:20
  triggerDay.setMinutes(20);
  ScriptApp.newTrigger("updateOfColosseum1").timeBased().at(triggerDay).create();
  
  // ギルド最終章の取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("評価値一覧");
  var list = sheet.getRange("K2:K" + sheet.getLastRow()).getValues();
  var last = "";
  for (var idx in list) {
    if (list[idx][0] != "") {
      last = list[idx][0].substring(0,1);
    }
  }
  
  // ギルド　02:40以降、20分おきに登録
  var hour = 2;
  var min = 40;
  for(var i = 1;i < last; i++){
    // ギルドのループ
    triggerDay.setHours(hour);
    triggerDay.setMinutes(min);
    ScriptApp.newTrigger("updateOfGuild" + i).timeBased().at(triggerDay).create();
    
    // 20分加算
    min+=20;
    
    if (min == 60) {
      // 1時間
      hour++;
      min = 0;
    }
  }
  
}

// 指定したトリガーを削除する
function deleteTrigger(func) {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == func) {
      ScriptApp.deleteTrigger(triggers[i]);
      break;
    }
  }
}

// バッチ個別処理 ------------------------------------------------------------------------------

// 共通処理（メイン）
function upd(targetList, triggerName){
  Logger.log(targetList);
  updateDataMain(targetList);
  deleteTrigger(triggerName);
}
// 共通処理（公式追加変更）
function updOf(targetList, triggerName){
  Logger.log(targetList);
  updateDataOfficial(targetList);
  deleteTrigger(triggerName);
}

// ヘアスタイル・ドレス・コート・トップスのアイテム情報を取得
function update1() {
  var targetList = ["ヘアスタイル", "ドレス", "コート", "トップス"];
  upd(targetList, "update1");
  updOf(targetList, "update1");
}

// ボトムス・靴下・シューズ・メイクのアイテム情報を取得
function update2() {
  var targetList = ["ボトムス", "靴下", "シューズ", "メイク"];
  upd(targetList, "update2");
  updOf(targetList, "update2");
}

// アクセの情報を取得
function updateAcc() {
  var targetList = ["アクセサリー"];
  upd(targetList, "updateAcc");
}
function updateAccOfficial() {
  var targetList = ["アクセサリー"];
  updOf(targetList, "updateAccOfficial");
}

// コロッセオのコーデを更新（前半9件）
function updateOfColosseum1() {
  var targetList = ["麗しき美女", "オフィスの女王", "華麗な女王", "宮廷舞踏会", "キュートな大人女子", "Xmasのホームパーティー", "ゴールデンホール", "サマーストーリー", "スポーツタイム"];
  updateCoordination(targetList);
  deleteTrigger("updateOfColosseum1");
}

// コロッセオのコーデを更新（前半8件）
function updateOfColosseum2() {
  var targetList = ["絶世の麗人", "夏のガーデンパーティー", "春の香り", "春のピクニック", "ビーチパーティーコーデ", "フェアリーガーデン", "冬の炎", "名探偵ホームズ"];
  updateCoordination(targetList);
  deleteTrigger("updateOfColosseum2");
}

// ギルド1章のコーデを更新
function updateOfGuild1() {
  var capter = 1;
  var targetList = [];
  for (var i = 1; i <= 7; i++) {
    targetList[i-1] = capter + "-" + i + "G";
  }
  updateCoordination(targetList);
  deleteTrigger("updateOfGuild1");
}

// ギルド2章のコーデを更新
function updateOfGuild2() {
  var capter = 2;
  var targetList = [];
  for (var i = 1; i <= 7; i++) {
    targetList[i-1] = capter + "-" + i + "G";
  }
  updateCoordination(targetList);
  deleteTrigger("updateOfGuild2");
}

// ギルド3章のコーデを更新
function updateOfGuild3() {
  var capter = 3;
  var targetList = [];
  for (var i = 1; i <= 7; i++) {
    targetList[i-1] = capter + "-" + i + "G";
  }
  updateCoordination(targetList);
  deleteTrigger("updateOfGuild3");
}

// ギルド4章のコーデを更新
function updateOfGuild4() {
  var capter = 4;
  var targetList = [];
  for (var i = 1; i <= 7; i++) {
    targetList[i-1] = capter + "-" + i + "G";
  }
  updateCoordination(targetList);
  deleteTrigger("updateOfGuild4");
}

// ギルド5章のコーデを更新
function updateOfGuild5() {
  var capter = 5;
  var targetList = [];
  for (var i = 1; i <= 7; i++) {
    targetList[i-1] = capter + "-" + i + "G";
  }
  updateCoordination(targetList);
  deleteTrigger("updateOfGuild5");
}

// ギルド6章のコーデを更新
function updateOfGuild6() {
  var capter = 6;
  var targetList = [];
  for (var i = 1; i <= 7; i++) {
    targetList[i-1] = capter + "-" + i + "G";
  }
  updateCoordination(targetList);
  deleteTrigger("updateOfGuild6");
}

// ギルド7章のコーデを更新
function updateOfGuild7() {
  var capter = 7;
  var targetList = [];
  for (var i = 1; i <= 7; i++) {
    targetList[i-1] = capter + "-" + i + "G";
  }
  updateCoordination(targetList);
  deleteTrigger("updateOfGuild7");
}





