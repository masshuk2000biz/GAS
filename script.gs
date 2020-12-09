function Main(){
  var ss = SpreadsheetApp.getActive();
  var setting_sheet = ss.getSheetByName("設定");
  var lineToken = setting_sheet.getRange(3,2).getValue();
  if(lineToken == ""){
    Browser.msgBox('トークンを入力してください');
  }
  else{
    var triggerKey = "trigger";
    getMail()
    setTrigger(triggerKey, "getMail")
    setting_sheet.getRange(5,2).setValue("転送中");
  }
}  

function deleteAllTrigger(){
  var ss = SpreadsheetApp.getActive();
  var setting_sheet = ss.getSheetByName("設定");
  var allTriggers = ScriptApp.getProjectTriggers();
  for( var i = 0; i < allTriggers.length; ++i ){
    if(allTriggers[i].getHandlerFunction() == "getMail"){
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  setting_sheet.getRange(5,2).setValue("停止");
}



function getMail(){
  
  var ss = SpreadsheetApp.getActive();
  var setting_sheet = ss.getSheetByName("設定"); 
  var get_interval = setting_sheet.getRange(4,2).getValue();
  var triggerKey = "trigger";
  
  //取得間隔
  var now_time= Math.floor(new Date().getTime() / 1000) ;//現在時刻を変換
  var time_term = now_time - (60 * get_interval) - 3; //変換
  
  //検索条件指定
  var strTerms = '(after:'+ time_term + ')';
  
  //指定した件名のスレッドを検索して取得 
  var myThreads = GmailApp.search(strTerms); 
  //スレッドからメールを取得し二次元配列に格納
  var myMessages = GmailApp.getMessagesForThreads(myThreads);
  
  
  for(var i in myMessages){
    for(var j in myMessages[i]){
      
      
      
      
      var strDate　=　myMessages[i][j].getDate();
      var strSubject　=　myMessages[i][j].getSubject();
      var strMessage　=　myMessages[i][j].getPlainBody().slice(0,200); //本文を200文字取得
      var attachments = myMessages[i][j].getAttachments(); //添付ファイルを取得
      
      //LINEにメッセージを送信
      //sendLine(strDate,strSubject,strMessage,attachment);

      
      
      var attachment = undefined;
      if(attachments == "" ) { 
        
        sendLine(strDate,strSubject,strMessage,attachment);
      }         
      
      if(attachments != ""){
        for(var k = 0; k < attachments.length; k++) {
          for(var k = 0; k < attachments.length; k++) {
            
            if(attachments[k].getContentType() === "application/octet-stream" ) {        
              attachment = attachments[k]
              sendLine(strDate,strSubject,strMessage,attachment);
            }      
            
            if(attachments[k].getContentType() === 'image/png' ) {
              attachment = attachments[k].getAs('image/png');
              sendLine(strDate,strSubject,strMessage,attachment);
            }
            
            if(attachments[k].getContentType() === 'image/jpeg' ) {
              attachment = attachments[k].getAs('image/jpeg'); 
              sendLine(strDate,strSubject,strMessage,attachment);
            }
            
          }
        }
      }
    }
  }
}



function sendLine(strDate,strSubject,strMessage,attachment){

  //Lineに送信するためのトークン

  var ss = SpreadsheetApp.getActive();
  var setting_sheet = ss.getSheetByName("設定");
  var lineToken = setting_sheet.getRange(3,2).getValue();
  
  var formData = {
    'message' : "\n" + strDate + "\n" + strSubject + "\n" + "\n" + strMessage,
    'imageFile': attachment  // 画像を添付 
  }
  var options =
   {
     "method"  : "post",
     "payload" : formData,
     "headers" : {"Authorization" : "Bearer "+ lineToken}

   };

   UrlFetchApp.fetch("https://notify-api.line.me/api/notify",options);
}

function sendLineAtt(strSubject,attachments){

  //Lineに送信するためのトークン

  var formData = {
   'message' : "\n" + strSubject,
   'imageFile': attachments  // 画像を添付
   }
  var options =
   {
     "method"  : "post",
     "payload" : formData,
     "headers" : {"Authorization" : "Bearer "+ lineToken}

   };

   UrlFetchApp.fetch("https://notify-api.line.me/api/notify",options);
}


//指定したkeyに保存されているトリガーIDを使って、トリガーを削除する
function deleteTrigger(triggerKey) {
  var triggerId = PropertiesService.getScriptProperties().getProperty(triggerKey);
  
  if(!triggerId) return;
  
  ScriptApp.getProjectTriggers().filter(function(trigger){
    return trigger.getUniqueId() == triggerId;
  })
  .forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  PropertiesService.getScriptProperties().deleteProperty(triggerKey);
}
 
//トリガーを発行
function setTrigger(triggerKey, funcName){
  var ss = SpreadsheetApp.getActive();
  var setting_sheet = ss.getSheetByName("設定"); 
  var get_interval = setting_sheet.getRange(4,2).getValue();  
  
  deleteTrigger(triggerKey);   //保存しているトリガーがあったら削除
  var dt = get_interval;
  //dt.setSeconds(dt.setSeconds() + (get_interval * 60000));  //10秒後に再実行
  //dt.setMinutes(dt.setMinutes() + (get_interval ));  //10秒後に再実行
  var triggerId = ScriptApp.newTrigger(funcName).timeBased().everyMinutes(dt).create().getUniqueId();
  //あとでトリガーを削除するためにトリガーIDを保存しておく
  PropertiesService.getScriptProperties().setProperty(triggerKey, triggerId);
}

function getMail_test(){
  
  var ss = SpreadsheetApp.getActive();
  var setting_sheet = ss.getSheetByName("設定"); 
  var get_interval = setting_sheet.getRange(4,2).getValue();

  


  //指定した件名のスレッドを検索して取得 
  //var myThreads = GmailApp.search(strTerms);
  var myThreads = GmailApp.search('', 0, 3);  
  //スレッドからメールを取得し二次元配列に格納
  var myMessages = GmailApp.getMessagesForThreads(myThreads);


  for(var i in myMessages){
    for(var j in myMessages[i]){
      
      
      
      
      var strDate　=　myMessages[i][j].getDate();
      var strSubject　=　myMessages[i][j].getSubject();
      var strMessage　=　myMessages[i][j].getPlainBody(); //本文
      var attachments = myMessages[i][j].getAttachments(); //添付ファイルを取得
      
      //LINEにメッセージを送信
      var attachment = undefined;
      
      if(attachments == "" ) {        
        sendLine(strDate,strSubject,strMessage,attachment);
      }      
      
      
      if(attachments != ""){
     
        for(var k = 0; k < attachments.length; k++) {
          if(attachments[k].getContentType() === "application/octet-stream" ) {        
            attachment = attachments[k]
            sendLine(strDate,strSubject,strMessage,attachment);
          }      
          
          if(attachments[k].getContentType() === 'image/png' ) {
            attachment = attachments[k].getAs('image/png');
            sendLine(strDate,strSubject,strMessage,attachment);
          }
          
          if(attachments[k].getContentType() === 'image/jpeg' ) {
            attachment = attachments[k].getAs('image/jpeg'); 
            sendLine(strDate,strSubject,strMessage,attachment);
          }
          
        }
      }
    }
  }
}
