// 実行開始ログ
function setStartLog(tgt){
  var sp = SpreadsheetApp.getActiveSpreadsheet();
  var addSheet = sp.getSheetByName("実行ログ");
  addSheet.deleteRow(10);
  addSheet.insertRows(2);
  addSheet.getRange(2,1).setValue(getDateTimeJSTString(new Date()));
  addSheet.getRange(2,3).setValue(tgt);
}

// 実行終了ログ
function setEndLog(rowNum, result){
  var sp = SpreadsheetApp.getActiveSpreadsheet();
  var addSheet = sp.getSheetByName("実行ログ");
  addSheet.getRange(rowNum,2).setValue(getDateTimeJSTString(new Date()));
  addSheet.getRange(rowNum,4).setValue(result);
}

// 実行中ログ
function setNowLog(rowNum, log){
  var sp = SpreadsheetApp.getActiveSpreadsheet();
  var addSheet = sp.getSheetByName("実行ログ");
  addSheet.getRange(rowNum,4).setValue(log);
}


// スクリプト動作ログ出力 開始処理
function setScriptStatusStart(range)
{
  // スマホでも動いたことがわかるように、処理中を表示しておく。
  range.setValue("処理中!");
}

// スクリプト動作ログ出力 終了処理
function setScriptStatusEnd(range)
{
  // スマホでも動いたことがわかるように、最終実行時間を表示しておく。
  range.setValue(getDateTimeJSTString(new Date()));
}

// スクリプト動作ログ出力 終了処理（エラー）
function setScriptStatusError(range)
{
  range.setValue("エラー！");
}

// スクリプト動作ログ出力
function setScriptStatusNow(range, value)
{
  range.setValue(value);
}

