// LINE developersのメッセージ送受信設定に記載のアクセストークン
var ACCESS_TOKEN = '自分のアクセストーンを入力する';

function doPost(e) {
  // WebHookで受信した応答用Token
  var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
  // ユーザーのメッセージを取得
  var userMessage = JSON.parse(e.postData.contents).events[0].message.text;
  // 応答メッセージ用のAPI URL
  var url = 'https://api.line.me/v2/bot/message/reply';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();
  //指定するセルの範囲を取得
  var date01 = new Date();
  date01 = new Date(date01.getFullYear(),date01.getMonth(),date01.getDate());
  //くり返し日数
  var count = 1;
  
  //var now2 = new Date(2018, 0, 1);   //1月は0、12月は11
  var month = userMessage.indexOf("月");
  var day = userMessage.indexOf("日");
  
  if(( month != -1 ) && ( day != -1 )){
    var n_day = Number(userMessage.slice(month+1,-1));
    var n_month = Number(userMessage.slice(0,month))-1;
    // もし、月が今より前であれば
    if( n_month < date01.getMonth() ){
      date01 = new Date(date01.getFullYear() + 1, n_month, n_day);
    }else{
      date01 = new Date(date01.getFullYear(), n_month, n_day);
    }
  }
  
  // 日付や用語の設定
  if( userMessage == "明日" ){
    date01.setDate( date01.getDate() + 1 );
  }
  if( userMessage == "明後日" ){
    date01.setDate( date01.getDate() + 2 );
  }
  if( userMessage == "週末" ){
    if( date01.getDay() == 6 ){
      count = count + 1;
    }else if( date01.getDay() == 0 ){
    }else{
      date01.setDate( date01.getDate() + (6 - date01.getDay()) );
      count = count + 1;
    }
  }
  if( userMessage == "来週" ){
    date01.setDate( date01.getDate() + (7 - date01.getDay()) );
    count = count + 6;
  }
  if( userMessage == "再来週" ){
    date01.setDate( date01.getDate() + (14 - date01.getDay()) );
    count = count + 6;
  }
  
  // シートのデータをとる
  var last_row = sheet.getLastRow();
  var str = "";
  var str2 = [ "", "", "", "", "", "", "", "", "", "" , "" , "" ];
  var r_range = sheet.getRange( 1, 8, last_row );
  var date02 = r_range.getValues();
  var r_range2 = sheet.getRange( 1, 9, last_row );    
  var date03 = r_range2.getValues();
  var r_range3 = sheet.getRange( 1, 10, last_row, 2 ).getValues();
  var building = sheet.getRange( 1, 20, last_row ).getValues();
  var eventname = sheet.getRange( 1, 5, last_row ).getValues();  
  var flag2 = 0;

  // データを作った日付けをクリアする
  for (var ix=1; ix <= last_row-1; ix++){
    date02[ix][0] = new Date(date02[ix][0].getFullYear(),date02[ix][0].getMonth(),date02[ix][0].getDate());
    date03[ix][0] = new Date(date03[ix][0].getFullYear(),date03[ix][0].getMonth(),date03[ix][0].getDate());
  }
  
  var d_count = 0;
  var d_count2 = 0;
  
  for( var jx=0; jx < count; jx++ ){
    var flag = 0;

    for (var ix=1; ix <= last_row-1; ix++){
      if( compareDate2(date01,date02[ix][0],date03[ix][0]) ){
        if( flag == 0 ){
          str2[ d_count2 ] = str2[ d_count2 ] + '【' + date01.getFullYear() + "/" + ( date01.getMonth() + 1 )  + "/" + date01.getDate() + "のイベント】\n";
        }

        var hours1 = r_range3[ix][0];
        var hours2 = r_range3[ix][1];
        
        if(( hours1 == "NULL" ) || ( hours2 == "NULL"  )){
          str2[ d_count2 ] = str2[ d_count2 ] + '・' + building[ix][0] + 'で' + eventname[ix][0] + 'が開催。\n';　
        }else{
          str2[ d_count2 ] = str2[ d_count2 ] + '・' + building[ix][0] + 'で' + eventname[ix][0] + 'が';　        
          str2[ d_count2 ] = str2[ d_count2 ] + toDoubleDigits(hours1.getHours())　+ ':' + toDoubleDigits(hours1.getMinutes());
          str2[ d_count2 ] = str2[ d_count2 ] + '-' + toDoubleDigits(hours2.getHours()) + ':' + toDoubleDigits(hours2.getMinutes()) +　'で開催。\n';
        }
        flag = 1;
        flag2 = 1;
        
        d_count++;
        
        if( d_count == 10 ){
          d_count = 0;
          d_count2++;
        }
        if(d_count2 == 5 ){
          break;
        }
      }
    }
    
    if(d_count2 == 5 ){
      break;
    }
    // 日付けを１日増やす
    date01.setDate( date01.getDate() + 1 );
  }
 
  if( flag2 == 0 ){
    str2[0] = 'イベントはありませんでした。';
  }else if( d_count == 0 ){
    d_count2--;
  }
  
//  var json_obj = [{"type":"text","text":str},{"type":"text","text":str}];
  var json_obj = "";
  //json_obj = [{"type":"text","text":str2[0]}];

  if( d_count2 == 0 ){
    json_obj = [{"type":"text","text":str2[0]}];
  }
  if( d_count2 == 1 ){
    json_obj = [{"type":"text","text":str2[0]},{"type":"text","text":str2[1]}];
  }
  if( d_count2 == 2 ){
    json_obj = [{"type":"text","text":str2[0]},{"type":"text","text":str2[1]},{"type":"text","text":str2[2]}];
  }  
  if( d_count2 == 3 ){
    json_obj = [{"type":"text","text":str2[0]},{"type":"text","text":str2[1]},{"type":"text","text":str2[2]},{"type":"text","text":str2[3]}];
  }  
  if( d_count2 == 4 ){
    json_obj = [{"type":"text","text":str2[0]},{"type":"text","text":str2[1]},{"type":"text","text":str2[2]},{"type":"text","text":str2[3]},{"type":"text","text":str2[4]}];
  }  
  if( d_count2 == 5 ){
    json_obj = [{"type":"text","text":str2[0]},{"type":"text","text":str2[1]},{"type":"text","text":str2[2]},{"type":"text","text":str2[3]},{"type":"text","text":str2[4]},{"type":"text","text":str2[5]}];
  }  


  UrlFetchApp.fetch(url, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify(
      {
      'replyToken': replyToken,
      'messages': json_obj,
    }
    ),
    });
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}


/* 2つの日付が等しいかを比較する */
function compareDate(date1,date2){
 
  if(date1.getFullYear() === date2.getFullYear() && date1.getMonth() === date2.getMonth() && date1.getDate() === date2.getDate()){
    return true;
  }else{
    return false;
  }
 
}


/* 2つの日付が等しいかを比較する */
function compareDate2(date1,date2,date3){
  var msecDiff1 = date2.getTime() - date1.getTime();
  var msecDiff2 = date3.getTime() - date1.getTime();
  
  if( ( msecDiff1 <= 0 ) && ( msecDiff2 >= 0 ) ){
    return true;
  }else{
    return false;
  }
 
}


var toDoubleDigits = function(num) {
  num += "";
  if (num.length === 1) {
    num = "0" + num;
  }
 return num;     
};