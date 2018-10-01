// 所持情報の引き継ぎ
function copyData(){
  var url = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("コーデ検索").getRange("B29").getValue();
  var inSpSheet = SpreadsheetApp.openByUrl(url);
  var outSpSheet = SpreadsheetApp.getActiveSpreadsheet();
  var category = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "アクセサリー", "メイク"];
  
  for (var index in category) {
    // category循環
    Logger.log(category[index] + "のデータをコピーします。");
    var outSheet = outSpSheet.getSheetByName(category[index]);
    var outLR = outSheet.getLastRow();
    var copySheet = inSpSheet.getSheetByName(category[index]);
    var copyLR = copySheet.getLastRow();
    
    var copyData = copySheet.getRange("A2:S" + copyLR).getValues();
    var outData = outSheet.getRange("A2:S" + outLR).getValues();
    
    // 出力先の所持データをリセットする
    for (var i = 0; i < outLR -1; i++) {
      outData[i][0] = "";
    }
    
    // 所持データの更新
    for (var i = 0; i < copyLR -1; i++) {
      if (copyData[i][0] == "◯"){
        // 所持していたら
        Logger.log(copyData[i][2] + "を所持に設定します。");
        for (var l = 0; l < outLR -1; l++) {
          // 出力先循環
          if (copyData[i][2] == outData[l][2]) {
            // 名称一致
            outData[l][0] = "◯";
          }
        }
      }
    }
    outSheet.getRange("A2:S" + outLR).setValues(outData);
  }
}
