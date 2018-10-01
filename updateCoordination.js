// 対象ターゲットリストのコーデを更新する（トリガーで呼び出される処理）
function updateCoordination (targetList) {
  for (var index in targetList) {
    if (targetList[index] != "") {

      // 名称を設定
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("コーデ検索").getRange("D7").setValue(targetList[index]);
      renewEval();
      Utilities.sleep(5000);
      
      // 処理実行
      var msg = raiseAllCategory(onAllCategoryCalc, null);
      
      // 終了時刻
      setScriptStatusEnd(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("コーデ検索").getRange("H16"));
      
      // コーデ転記
      Utilities.sleep(5000);
      codeSace();
    }
  }
}
