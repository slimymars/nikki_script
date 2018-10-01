
// 所持更新 ------------------------------------------------------------------------------------------------------------------------- //
function koreMotteru(){
  var logRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("所持追加用シート").getRange("B2");
  try{
    Logger.log("start koreMotteru");
    setStartLog("所持情報更新");
    setScriptStatusStart(logRange);
    var sp = SpreadsheetApp.getActiveSpreadsheet();
    var addSheet = sp.getSheetByName("所持追加用シート");
    var addLR = addSheet.getLastRow();
    var addData = addSheet.getRange(3, 1, addLR - 2, 2).getValues();
    addSheet.getRange(3, 2, addLR - 2, 1).clearContent();
    var category = ["ヘアスタイル", "ドレス", "コート", "トップス", "ボトムス", "靴下", "シューズ", "アクセサリー", "メイク"];
    Logger.log(addData.length + "件の追加を試みます。");
    var repeatRow = addData.length;
    var delCnt = 0;
    
    for (var i = 0; i < repeatRow; i++) {
      search : for (var index in category) {
        // 各シートの探索を行う
        setNowLog(2, addData[i][0] + "を検索しています。");
        setScriptStatusNow(logRange, "検索：" + addData[i][0] + "を検索しています。");
        Logger.log(addData[i][0] + "を" + category[index] + "シートで探索します。");
        var sheet = sp.getSheetByName(category[index]);
        var sheetLR = sheet.getLastRow();
        var sheetData = sheet.getRange(2, 3, sheetLR - 1, 1).getValues();
        for (var s = 0; s < sheetData.length; s++) {
          if (returnHalfString(addData[i][0]) == returnHalfString(sheetData[s][0])) {
            // 名称一致
            sheet.getRange(s + 2, 1).setValue("◯");
            
            addSheet.deleteRow(i + 3 - delCnt);
            addSheet.insertRows(addData.length + 3);
            i - i -1;
            delCnt = delCnt + 1;
            break search;
          }
        }
      }
    }
    
    setScriptStatusEnd(logRange);
    Logger.log("end koreMotteru");
    //SpreadsheetApp.getUi().alert("所持データを更新しました。");
    setEndLog(2, "正常終了")
  } catch(e){
    Logger.log(e);
    setEndLog(2, e);
    setScriptStatusError(logRange);
  }
}
