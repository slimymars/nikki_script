// 転記する
function codeSace(){
  var codeSearchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("コーデ検索");
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("コーデ記録");
  var lastRow = outputSheet.getLastRow();
  var find = false;
  
  // コーデ名称取得
  var codeName = codeSearchSheet.getRange("D7").getValue();
  if (codeName == "カスタム") {
    Logger.log("カスタムコーデのため、保存処理をスキップします。");
    return; // カスタムのときは何もしない
  }
  
  for (var i = 1; i <= lastRow; i++) {
    // 行探索
    for (var l = 2; l <= 26; l++ ){
      // 列探索
      var checkName = outputSheet.getRange(i, l).getValue();
      if (checkName == '') {
        // 空白行なら列探索を終了
        break;
      }

      if (checkName == codeName) {
        // コーデ名称一致
        Logger.log(codeName + "のコーデを保存します。");
        
        // 更新日時
        var data = getDateTimeJSTString(new Date());
        outputSheet.getRange(i + 1, l).setValue(data);
        
        // 前半
        data = codeSearchSheet.getRange("M2:M17").getValues();
        var drs = codeSearchSheet.getRange("M3").getBackground();
        if (drs == "#b7e1cd") {
          // ドレスセルの背景が緑（優位）のとき、上下を「-」に置き換える
          data[3][0] = "-";
          data[4][0] = "-";
        } else {
          // ドレスセルの背景が赤のとき、ドレスを「-」に置き換える
          data[1][0] = "-";
        }
        outputSheet.getRange(i + 2, l, 16, 1).setValues(data);
        
        // 後半
        data = codeSearchSheet.getRange("P2:P18").getValues();
        var hnd = codeSearchSheet.getRange("P7").getBackground();
        if (hnd == "#b7e1cd") {
          // 両手持ちセルの背景が緑（優位）のとき、右手持ち・左手持ちを「-」に置き換える
          data[3][0] = "-";
          data[4][0] = "-";
        } else {
          // 両手持ちセルの背景が赤のとき、両手持ちを「-」に置き換える
          data[5][0] = "-";
        }
        outputSheet.getRange(i + 18, l, 17, 1).setValues(data);
        
      }
    
    }
    // 次の行へ
    i = i + 34;
    
    if (find) {
      // データが見つかったら終了
      break;
    }
  }
}
