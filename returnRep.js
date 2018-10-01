// 日付を日本時間にして返却する
function getDateTimeJSTString(value)
{
  // 日本時間（JST）に変換して、文字列として返却する。
  return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}


// 半角文字にして返すヨ～
function returnHalfString(value){
  var repVal = "";
  value = value.replace(/[‐－―]/g, '-').replace(/　/g, ' ').replace(")","）").replace("(","（");
  // 英数字
  value.split("").forEach(function (s) {
    //repVal += String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    // 文字化けしちょるのでもう文字コードそのまま返そうず
    repVal += s.charCodeAt(0) - 65248 + ",";
  });
  // 記号
  return repVal;
}

// 文字列にして返すヨ～
function returnString(value){
  var repVal = value.toString();
  var idx = repVal.indexOf(".");
  if (idx != -1) {
    repVal = repVal.substring(0, repVal.indexOf("\."));
  }
  return repVal;
}