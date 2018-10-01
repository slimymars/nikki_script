// 同一コーデチェック
function checkCode() {
  var codeSearchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("コーデ検索");
  var searchList = codeSearchSheet.getRange("A2:J2").getValues();
  
  var codeEvalListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("評価値一覧");
  var lastRow = codeEvalListSheet.getLastRow();
  var checkRange = codeEvalListSheet.getRange("A2:J35").getValues();
  for (var i = 0; i < 34; i++){
    if (checkRange[i][1] == searchList[0][1]
       && checkRange[i][2] == searchList[0][2]
       && checkRange[i][3] == searchList[0][3]
       && checkRange[i][4] == searchList[0][4]
       && checkRange[i][5] == searchList[0][5]
       && checkRange[i][6] == searchList[0][6]
       && checkRange[i][8] == searchList[0][8]) {
      // 属性・タグが一致する
      Logger.log(checkRange[i][0] + "と同一のカスタム値です。");
      codeSearchSheet.getRange("B4").setValue(checkRange[i][0] + "と同一のカスタム値です");
      return;
    }
    
    // 次の行
    i++;
  }
  
  // 一致無し
  codeSearchSheet.getRange("B4").clearContent();
  
}
