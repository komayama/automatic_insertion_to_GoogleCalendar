//Webを公開するための関数
function doGet(){
  //アクティブなユーザを取得
  var objUser = Session.getActiveUser();
  var mail = objUser.getEmail();

  //月を取得
  var Datetime = new Date();
  var month = Datetime.getMonth()+1;

  //カレンダー取得（詳細な部分は画面に表示しない）
  var calendarid = CalendarApp.getCalendarById(mail);
  var date = Datetime;　//対象月を指定
  var startDate=new Date(date); //取得開始日
  var endDate=new Date(date);
  endDate.setMonth(endDate.getMonth()+1);　//取得終了日
  var myEvents = calendarid.getEvents(startDate,endDate);
  var usename = "";

  var maxRow=1;
  for each(var evt in myEvents){
      usename += evt.getTitle()+"|"+ //イベントのタイトル

    maxRow++;
  }

  t = HtmlService.createTemplateFromFile('index.html');
  t.title = '共和サービスセンター';
  t.mailaddr = mail;
  t.calendar = calendarid;
  t.putonmonth = month;
  t.buttonname = 'Googleカレンダーに追加';
  return t.evaluate();
}

function doPost(e){
  //選択されたスタッフの名前を取得
  var name = e.parameter.staffname;

  //メールアドレスをindexからとってくる技が見当たらなかったので応急処置
  var objUser = Session.getActiveUser();
  var mail = objUser.getEmail();

  Logger.log(name+" "+mail);

  getPartTime(name,mail);

  return HtmlService.createHtmlOutputFromFile("output");
}

function getPartTime(name,mail){
  var Datetime = new Date();
  var year = Datetime.getFullYear();
  var month = Datetime.getMonth()+1;

  names = getName(month);
  //列を取得
  for(var i=0;i < names.length; i++){
    if(names[i] == name){
      var clumnum = i+5;
    }
  }
  jobDateTime = getspreadonPartTime(clumnum,month);
  Logger.log(jobDateTime);
  for(var day=1;day<30;day++){
    if(jobDateTime[day-1] != ""){
      setCalendar(jobDateTime[day-1],day,month,year,mail);
    }
  }
}

//シフトの名前を取得
function getName(month){
  month = 2;
  var ss  = SpreadsheetApp.openById('1QjAblUNiY21ynf-NCiYM0awkx5GPy7x0RbLuUFzahFE');
  var sheet = ss.getSheetByName(month + '月');
  var parttimeName = sheet.getRange("A5:A13").getValues();
  Logger.log(parttimeName);
  return parttimeName;
}

//スプレッドシートのアルファベットを読み込む
function getspreadonPartTime(clumnum,month){
  var ss  = SpreadsheetApp.openById('1QjAblUNiY21ynf-NCiYM0awkx5GPy7x0RbLuUFzahFE');
  var sheet = ss.getSheetByName(month + '月');
  var partTime = [];

  //ここで日付分回す
  for(var i=1; i<=31; i++){
    partTime[i-1] = sheet.getRange(clumnum,i+1).getValue();
  }
  return partTime;
}

function setCalendar(alphabet,day,month,year,mail) {
  var jobtime = timeCongnition(alphabet);
  var calendar = CalendarApp.getCalendarById(mail);

  calendar.createEvent('共和 '+alphabet+'番',
                       new Date(year + '/' + month + '/' + day + ' ' + jobtime[0]),
                       new Date(year + '/' + month + '/' + day + ' ' + jobtime[1])).addPopupReminder(1440);
}


//文字から時間に変更
function timeCongnition(alphabet){
  var starttime,endtime

  switch(alphabet){
    case 'B':
      starttime = "08:30:00";
      endtime = "13:30:00";
      break;
    case 'A':
      starttime = "08:00:00";
      endtime = "17:00:00";
      break;
    case 'D':
      starttime = "10:00:00";
      endtime = "20:00:00";
      break;
    case 'AD':
      starttime = "08:00:00";
      endtime = "20:00:00";
      break;
    case 'N':
      starttime = "10:00:00";
      endtime = "15:00:00";
      break;
    case 'E':
      starttime = "15:00:00";
      endtime = "19:00:00";
      break;
    case 'L':
      starttime = "15:00:00";
      endtime = "20:00:00";
      break;
    case 'X':
      starttime = "17:00:00";
      endtime = "20:00:00";
      break;
    case 'H':
      starttime = "08:30:00";
      endtime = "16:00:00";
      break;
    case 'F':
      starttime = "10:00:00";
      endtime = "19:00:00";
      break;
    case 'K':
      starttime = "13:30:00";
      endtime = "19:00:00";
      break;
    case 'C':
      starttime = "12:00:00";
      endtime = "17:00:00";
      break;
    case 'G':
      starttime = "11:00:00";
      endtime = "19:00:00";
      break;
    case 'EX':
      starttime = "15:00:00";
      endtime = "20:00:00";
      break;
    default:
      starttime = null;
      endtime = null;
      break;
  }

  return[starttime,endtime];
}
