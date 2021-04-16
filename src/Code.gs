/****************************************************************
Backend Team勉強会抽選Bot
抽選ルール： Round-robin + Random
データ保存用Google Sheetsファイル：
  A列： Slackのdisplay_name　（手動で記入）
  Slackのuser_id　（手動で記入 メンションするのにrequired）
  C列： １サイクルにおける発表status (1:発表済 0:未発表) （プログラムで埋める）
  D列: 発表された日付 （プログラムで埋める）
****************************************************************/

const SPREAD_SHEETS_ID = "xxxxx";
const SLACK_WEB_HOOK_URL = "https://hooks.slack.com/services/xxx/xxxxxx";

// 列
const COLUMN_NAME = 1;
const COLUMN_ID = 2;
const COLUMN_STATUS = 3;
const COLUMN_DATE = 4;

// １サイクルにおける発表status値
const STATUS_ASSIGNED = 1;
const STATUS_NOT_ASSIGNED = 0;

// header行数
const HEADER_ROWS = 1;

// データ開始行
const DATA_START_ROW_NO = HEADER_ROWS + 1;

// Interval Days
const INTERVAL_DAYS = 14;

// 勉強会発表者抽選function
function RobinStudySession() {
  
    var date = new Date();
    var day = date.getDate();  // 日付のみ取得
    date.setDate(day + INTERVAL_DAYS);
    Logger.log(date);
    if (isJapaneseHoliday(date)) {
        Logger.log(date + " はお休みでした！");
        SendSlackMessage("", date);
        return;
    }
  
    // スプレッドシートを指定（URLからコピペする）
    var spreadsheet = SpreadsheetApp.openById(SPREAD_SHEETS_ID);
    //シートの名前を指定する
    var sheet = spreadsheet.getSheetByName('members');
  
    var names = sheet.getRange(DATA_START_ROW_NO, COLUMN_NAME, sheet.getLastRow()-HEADER_ROWS).getValues();
    var members = sheet.getRange(DATA_START_ROW_NO, COLUMN_ID, sheet.getLastRow()-HEADER_ROWS).getValues();
    var roundRobinStatus = sheet.getRange(DATA_START_ROW_NO, COLUMN_STATUS, sheet.getLastRow()-HEADER_ROWS).getValues();
    var spokenDate = sheet.getRange(DATA_START_ROW_NO, COLUMN_DATE, sheet.getLastRow()-HEADER_ROWS).getValues();
    console.log("members:", members)
    console.log("roundRobinStatus", roundRobinStatus)
    console.log("spokenDate", spokenDate)

    // 未発表者の中からランダムで抽選
    var unassignedIndices = [];
    var firstRound = (getRowCountNotBlank(spokenDate) !== members.length) ? true : false;
    if (firstRound) {
        for (var i = 0; i < roundRobinStatus.length; i++) {
            if (roundRobinStatus[i][0] !== STATUS_ASSIGNED) {
                unassignedIndices.push(i)
            }
        }    
    }
    else{ // ２週目以降は、過去の発表日の古い者からランダム抽選（前のサイクルにおける後半の発表者は次のサイクルでも後半になるように調整）
        var data = []
        for (var i = 0; i < members.length; i++) {
          if (roundRobinStatus[i][0] !== STATUS_ASSIGNED) {
            var date = new Date(spokenDate[i][0]).getTime();
            var item = [date, i]
            data.push(item)
          }
        }
        console.log("data:",data);
        var dataSortedBySpokenDate = data.sort();
        console.log("dataSortedBySpokenDate", dataSortedBySpokenDate);
        var range = dataSortedBySpokenDate.length / 2;
        console.log("range", range);
        for (var i = 0; i < range; i++) { 
          unassignedIndices.push(dataSortedBySpokenDate[i][1]);
        }
    }

    console.log("unassignedIndices:",unassignedIndices);
    var index = Math.floor(Math.random() * Math.floor(unassignedIndices.length));
    var memberIndex = unassignedIndices[index]
    console.log("memberIndex", memberIndex);
    SendSlackMessage(members[memberIndex], date)
    
    // データを更新
    sheet.getRange(memberIndex+DATA_START_ROW_NO, COLUMN_STATUS).setValue(STATUS_ASSIGNED);
    sheet.getRange(memberIndex+DATA_START_ROW_NO, COLUMN_DATE).setValue(new Date());  
  
    // 一週を回した場合はデータを「Status」列をReset
    var roundRobinStatus = sheet.getRange(DATA_START_ROW_NO, COLUMN_STATUS, sheet.getLastRow()-HEADER_ROWS).getValues();
    if (getRowCountNotBlank(roundRobinStatus) === members.length) {
       for (var i = 0; i < members.length; i++) {
         sheet.getRange(i+DATA_START_ROW_NO, COLUMN_STATUS).setValue(STATUS_NOT_ASSIGNED);
       }
    }
}


/************************************************************************
 *
 * Gets the last row number based on a selected column range values
 *
 * @param {array} range : takes a 2d array of a single column's values
 *
 * @returns {number} : the row numbers with a value. 
 * @example:
 *
 * var columnToCheck = sheet.getRange("A:A").getValues();
 * var lastRow = getRowCountNotBlank(columnToCheck);
 */ 
function getRowCountNotBlank(range){
  var rowNum = 0;
  for(var row = 0; row < range.length; row++){
    if(range[row][0] !== "" && range[row][0] !== STATUS_NOT_ASSIGNED){
      rowNum = rowNum + 1;
    }
  }
  return rowNum;
};

// Slackへ送信するfunction
function SendSlackMessage(slackId, date) {  
  var formatedDate = date.toISOString().slice(0,10);

  var jsonData =
      {
        "text" : `<!channel> * ${formatedDate} の勉強会の発表者は* ` + "<" + slackId + ">" + "です！ 宜しくお願いします! :beer:"
      };
  
  // 休日の場合はSKIP
  if (slackId === "")
  {
    jsonData = {
        "text" : `<!channel> ${formatedDate} はお休みのため、勉強会無し :pear: です。宜しくお願いします! :relieved:`
      };
  }
  
  Logger.log(jsonData);
  var payload = JSON.stringify(jsonData);
  var options =
      {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : payload
      };
  UrlFetchApp.fetch(SLACK_WEB_HOOK_URL, options);
}


function isJapaneseHoliday(today){
  const calendars = CalendarApp.getCalendarsByName('Holidays in Japan');
  const count = calendars[0].getEventsForDay(today).length;
  return count > 0;
}


