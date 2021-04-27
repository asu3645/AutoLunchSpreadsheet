/*
需要更新的項目
1.http://gsyan888.blogspot.com/2017/07/apps-script-spreadsheet-shortener-url.html
2.https://webapps.stackexchange.com/questions/76050/google-sheets-function-to-get-a-shortened-url-from-bit-ly-or-goo-gl-etc
3.
4.
------------function說明------------
NotifyNextWeek()     週一、五通知換誰訂餐 
NotifyLunchLink()    將會傳送午餐連結、結單時間到LINE
UpdateLunchLink()    "規則"工作表會更新所有菜單列表
MoveTodayMenu()      把當天的菜單工作表移到第一個位置
doGet(e)             目前利用IFTTT呼叫此函數
MoveSheet()          把當周選出來的菜單移動到最前面
isHoliday()          判斷今天有沒有放假

-----------已設定的時間觸發器------------
呼叫func             呼叫來源               呼叫週期
NotifyNextWeek()      內建                 每周五1600~1700
                     IFTTT(doGet)         每周一1045
NotifyLunchLink()    IFTTT(doGet)         每日0845(放心 假日你不會被吵到 程式會過濾假日)
UpdateLunchLink()     內建                 每日0700~0800
MoveTodayMenu()       內建                 每日0700~0800
*/

var ss = SpreadsheetApp.openById("ID");
var weekDay = new Date().getDay(); //今天星期幾
var nowTimes = new Date();
var ruleSheet = ss.getSheetByName('規則');
var log = ss.getSheetByName('LOG');

function NotifyNextWeek() {
  var people;
    var msg = "";
  
  if(weekDay == 5){ 
    ruleSheet.getRange("E2").setValue(ruleSheet.getRange("E2").getValue() + 1);   //計數器++
    var _cell = ruleSheet.getRange("E3");
    people = ruleSheet.getRange(GetHolder(),2).getValue();//換誰訂餐
    msg = "\n下禮拜換" + people + "訂餐哦!\n" + "可以先設定下禮拜要ㄘ甚麼";
    SendLine(msg);
    log.appendRow([nowTimes,msg]);
  }
  if(weekDay == 1){
    people = ruleSheet.getRange(GetHolder(),2).getValue();//換誰訂餐
    msg = "這禮拜是" + people + "訂餐!";
    SendLine(msg);
  }
  return true;
}

function NotifyLunchLink(){
  if(!isHoliday()){
    var regExp = RegExp("[0-9]+");
    var Link = regExp.exec(ruleSheet.getRange(2, 9).getFormula())[0];
    
    var Msg = "\n今日午餐 09:40結單\n";
    Msg += "https://docs.google.com/spreadsheets/d/ID";
    Msg += Link;
    
    SendLine(Msg);
    SendLine("訂餐的時候記得不用附餐具!");
    log.appendRow([nowTimes,Msg]);
  }
}

//更新菜單網址
function UpdateLunchLink(){
  var sheetLen = ss.getSheets().length;
  
  for(var shet = 2;shet<sheetLen;shet++)
  {
    //名稱 
    var TodayMenu = ss.getSheets()[shet];
    
    //儲存格部分
    //名稱
    var MenuName = ruleSheet.getRange(shet,9);
    //連結
    var MenuLink = ruleSheet.getRange(shet,10);
    
    var HyperLink = "=HYPERLINK(\"#gid=" + TodayMenu.getSheetId() + "\",\"" + TodayMenu.getName() + "\")"
    //設定名稱
    MenuName.setValue(HyperLink);
    //enuLink.setValue(TodayMenu.getSheetId());    
  }
    
  console.log("更新菜單網址 -> 成功");
}

function doGet(e){
  try
  {
    if(!isHoliday() && e != undefined)
    {
      var actionName = e.parameter.action;
      switch(actionName){
        case "collectMoney":
          NotifyCollectMoney();
          break;
        case "NotifyOrdr":
          NotifyNextWeek();
          break;
        case "NotifyMenu":
          NotifyLunchLink();
          break;
        default:
          SendLine("\n不要亂打request! " + actionName);
          break;
      }
      console.log("Exec : " + actionName);
      return true;
    }
  }
  catch(e)
  {
    log.appendRow([new Date(),"doGet: " + actionName]);
     
  }
}

function test(){
    //var MenuName = ss.getSheets()[3].getProtections();
  //var a = String.Format("{0}","test");
  //console.log('GGGGGGGGGGTEST');
  //log.appendRow([new Date(),"E OF TEST"]);
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
        'headers': {
            'Authorization': 'Bearer ' + "TOKEN",
        },
        'method': 'post',
        'payload': {
          'message': "good",
          //'message': '這是測試',
          'stickerPackageId':"1",
          'stickerId':"106"

        }
    });
}

function ChangeMenu(){
  var count = ruleSheet.getRange("E5").getValue();
  var count2 = ruleSheet.getRange("E5").getValue();
  var MenuList = new Array(7);
  var MenuName;
  
  if(count > 0)
  {
    
    for(var shet = 2; count > 0;shet++)
    {
      //取得該行菜單有沒有設定星期幾
      var WhatDay = ruleSheet.getRange(shet,8);
      MenuName = ruleSheet.getRange(shet,9).getValue();
      
      if(WhatDay.getValue() > ""){
        MenuList[WhatDay.getValue() - 1] = MenuName;
        //重設天數
        WhatDay.setValue("");
        count --;
        shet = 2;
      }
    }
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();//活動中的工作表
  
    //移動表單
    for(var i = 0; i < count2; i++)
    {
      MenuName = MenuList[i];
      var newcurrSheet = spreadsheet.getSheetByName(MenuName);
      
      if(i + 3 != newcurrSheet.getIndex())
      {
        Move(spreadsheet,newcurrSheet,i + 3);
      }
    }
    
    SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("規則"));
  }
  
  UpdateLunchLink();
}

//移動菜單順序
//還沒設定執行時間
function MoveSheet(){
  
  var count = ruleSheet.getRange("D5").getValue();
  var MenuList = ["Monday Holiday","Tuesday Holiday","Wensday Holiday","Thusday Holiday","Friday Holiday","Saturday Holiday","Sunday Holiday"];
  
  if(count > 0)
  {
    for(var shet = 2; count > 0;shet++)
    {
      //取得該行菜單有沒有設定星期幾
      var WhatDay = ruleSheet.getRange(shet,8);
      var MenuName = ruleSheet.getRange(shet,9).getValue();
      
      if(WhatDay.getValue() > ""){
        MenuList[WhatDay.getValue() - 1] = MenuName;
        //重設天數
        WhatDay.setValue("");
        count --;
      }
    }
  }
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();//活動中的工作表
  
  //移動表單
  for(var i = 0; i < MenuList.length; i++)
  {
    var MenuName = MenuList[i];
    var currSheet = spreadsheet.getSheetByName(MenuName);
    Move(spreadsheet,currSheet,i + 3);
  }
  
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("規則"));
  
  UpdateLunchLink();
}

//把當天的菜單移到第一個
//每天工作日執行
function MoveTodayMenu(){
  if(!isHoliday()){
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();//活動中的工作表
    var currSheet = spreadsheet.getSheets()[weekDay + 1];
    
    Move(spreadsheet,currSheet,3);
    
    console.log("移動今天菜單 -> 成功");
    
    UpdateLunchLink();
  }
}

function Move(spreadsheet,currSheet,idx){
  SpreadsheetApp.setActiveSheet(currSheet);
  SpreadsheetApp.flush();
  Utilities.sleep(200);
  spreadsheet.moveActiveSheet(idx);
  SpreadsheetApp.flush();
  Utilities.sleep(200);
}

function NotifyCollectMoney(){
  SendLine("收錢囉!");
  log.appendRow([nowTimes,"收錢囉!"]);
}

//判斷有沒有放假
function isHoliday(){
  var range;
  switch(weekDay){
    case 1:
      range = "F2";
      break;
    case 2:
      range = "F3";
      break;
    case 3:
      range = "F4";
      break;
    case 4:
      range = "F5";
      break;
    case 5:
      range = "F6";
      break;
    case 6:
      range = "F7";
      break;
    case 0:
      range = "F8";
      break;
  }
  if(ruleSheet.getRange(range).getValue()){return true;}
  else{return false;}
}

//取得當周負責人 回傳為整數
function GetHolder(){
  var totalPeople = ruleSheet.getRange("E4").getValue(); //總人數
  var result = ruleSheet.getRange("E3").getValue();
  result = result == 0 ? totalPeople : result;
  
  return result;
}

//傳送訊息到LINE
function SendLine(msg) {
  
    var currToken = ruleSheet.getRange(GetHolder(),3).getValue();
  if(currToken > ''){
    UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
        'headers': {
            'Authorization': 'Bearer ' + currToken,
        },
        'method': 'post',
        'payload': {
          'message': msg,
          //'message': '這是測試',
          //'stickerPackageId':1,
          //'stickerId':106

        }
    });
  }
}



//Test
function ShortenUrl(){ 
  var a = ruleSheet.getRange("E2").getValue()
  var regExp = RegExp("[0-9]+");
  var Link = regExp.exec(ruleSheet.getRange(2, 9).getFormula())[0];
  var url = "https://docs.google.com/spreadsheets/d/ID" + Link;
  var longUrl = UrlShortener.newUrl()
      .setLongUrl(url);

      var shortUrl = UrlShortener.Url.insert(longUrl);

  
  var apiKey, post_url, options, result;
  post_url ="https://www.googleapis.com/urlshortener/v1/url";
  apiKey = 'AIzaSyAFUHBZyhKo-v441VTr608dxwi76ppOHQc';//here is real apiKey;
  post_url += '?key=' + apiKey;
  var options =
      {
        'method':'post',
        'headers' : {'Content-Type' : 'application/json'},
        "resource": {"longUrl": url},
        'muteHttpExceptions': true
      };
  result = UrlFetchApp.fetch(post_url, options);
  Logger.log(shortUrl);
  
} 

//https://firebasedynamiclinks.googleapis.com/v1/shortLinks?key=AIzaSyCRcDF28C5BfR7HBYvw4BZekEs7BKsmhKs

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Shorten")
    .addItem("Go !!","rangeShort")
    .addToUi()  
}

function rangeShort() {
  var range = SpreadsheetApp.getActiveRange(), data = range.getValues();
  var output = [];
  for(var i = 0, iLen = data.length; i < iLen; i++) {
    var url = UrlShortener.Url.insert({longUrl: data[i][0]});
    output.push([url.id]);
  }
  range.offset(0,1).setValues(output);
}