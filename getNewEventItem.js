function myFunction() {
  var out = [];

  // 【第三回】愛の誓いのURL
  // テスト用。最終的にはシートから取得する。
  var url = "https://miraclenikki.gamerch.com/%E3%80%90%E7%AC%AC%E4%B8%89%E5%9B%9E%E3%80%91%E6%84%9B%E3%81%AE%E8%AA%93%E3%81%84";
  var response = UrlFetchApp.fetch(url);
  
  // 正規表現
  var allReg = /<.*>.*<\/.*>/gi;
  var stageReg = /<div class="t-line-img">.*/gi;
  var endReg = /<\/div>/;
  var rowReg = /<tbody>(.*)<\/tbody>/gi;
  var itemReg = /(.*data-col="2"><a href.*>)(.*)(<\/a><\/td>.*)/gi;
  
  var loopStartRef = /.*js_async_main_column_text.*/;
  var loopEndReg = /.*ui_comment_tab_container.*/; // ループ終了向け
  /*
  div.t-line-img
  table
  
  */
  
  // 探索
  var stageMatch = stageReg.exec(response.getContentText());
  var t;
  var stage;
  while ((t = allReg.exec(response.getContentText())) !== null ){
    var itemList = [];

    if (itemReg.test(t)) {
      // アイテム
      var item = t[0].split("/td>");
      var tmpItem = [];
      for(var i = 0; i < item.length; i++) {
        var str = item[i] + "/td>";
        if (/data-col="2"/.test(str) || /data-col="3"/.test(str)
          || /data-col="4"/.test(str)
          || /data-col="6"/.test(str) || /data-col="7"/.test(str)
          || /data-col="8"/.test(str) || /data-col="9"/.test(str)
          || /data-col="10"/.test(str) || /data-col="11"/.test(str)
          || /data-col="12"/.test(str) || /data-col="13"/.test(str)
          || /data-col="14"/.test(str) || /data-col="15"/.test(str)
          || /data-col="16"/.test(str) || /data-col="17"/.test(str)){
          tmpItem.push(str.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,''));
        } else if (/data-col="5"/.test(str)){
          // レア度のときはハートマークを削除
          str = str.substr(1);
          tmpItem.push(str.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,''));
        }
      }
      if(tmpItem.length == 16){
        itemList.push(tmpItem);
        Logger.log(tmpItem);
      }
    }
    
    if (loopEndReg.test(t)) {
      break; // loopEndRegと一致するHTMLを取得したら終了する
    }
  }
  
}
