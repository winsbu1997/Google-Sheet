var bg = 7;
var fn = 8;
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var dataSheet = ss.getSheetByName("Database");
var lastCol = dataSheet.getLastColumn();
var lastRow = dataSheet.getLastRow();
var data = dataSheet.getDataRange().getValues();

var month = [0,1,2,3,4,5,6,7,8,9,10,11,12];
var dayOfMonth = [0,31,28,31,30,31,30,31,30,31,30,31,30];
function check(year){
  if ((year % 4===0 &&year%100 !==0 && year % 400 !==0)||(year%100===0 && year % 400===0)){
    return 1;
  }
  return 0;
}
function subTime(str1, str2){
  var kq1, kq2;
  var date1 = new Date(str1);
  var day1 = date1.getDate(); 
  var month1 = date1.getMonth() + 1; 
  var year1 = date1.getFullYear();
  
  var date2 = new Date(str2);
  var day2 = date2.getDate(); 
  var month2 = date2.getMonth() + 1; 
  var year2 = date2.getFullYear();
  
}

function onOpen1(){
  var a = []; 
  for(var i = 1; i < data.length; i++){
    var list = data[i];
    var time1 = new Date(list[0]);    
    var timeZone = ss.getSpreadsheetTimeZone(); 
    time1 = Utilities.formatDate(time1, timeZone, "dd/MM/yyyy");
    a.push(time1);
  }
  dataSheet.getRange(2, 1).setValue(a);
}

function convertArr(x){
  var kq = [];
  var date1 = new Date(x);
  kq.push(date1.getDate()); 
  kq.push(date1.getMonth() + 1); 
  kq.push(date1.getFullYear());
  return kq;
}

function Kq(startDate, endDate, t){
  var x1 = convertArr(startDate);
  var x2 = convertArr(endDate);
  var x = convertArr(t);
  if(x1[2] > x[2] || x2[2] < x[2]) return 0;
  if(x1[2] == x[2]){
    if(x1[1] > x[1]) return 0;
    if(x1[1] == x[1]){
      if(x1[0] > x[0]) return 0;
    }
  }
  if(x2[2] == x[2]){
    if(x2[1] < x[1]) return 0;
    if(x2[1] == x[1])
      if(x2[0] < x[0]) return 0;
  }
  return 1;
}

function onEdit(){
  var filterData = [];
  if(sheet.getSheetName() == "Dữ liệu"){
    var activeCell = sheet.getActiveCell();
    var editRow = activeCell.getRow();
    var editCol = activeCell.getColumn();
    var bg = sheet.getRange(2,7).getValue();
    var fn = sheet.getRange(2,8).getValue();
    if(editRow == 2 && editCol == 7 || editRow == 2 && editCol == 8){
      for(var i = 1; i < data.length; i++){
        var list = data[i];
        if(Kq(bg, fn, list[0])){
          filterData.push(list);
        }
      }
      
      var range = sheet.getRange(2,1,lastRow, lastCol);
      range.clear();
      sheet.getRange(2,1,filterData.length, lastCol).setValues(filterData);
    }   
  }
}
