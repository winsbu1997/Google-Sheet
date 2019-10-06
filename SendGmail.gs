// cột chứa ô check box
var colCheckBox = 25;

// hàng chứa tên trường Vd: Số báo giá
var rowHeader = 3;

// trường lấy subject để gửi mail
var colSubject = 2;

// chiều dài số cột cần lấy dữ liệu
var lastValCol = 21;

// thay đổi giá trị ở trên.
function sortRange(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var list = sheet.getDataRange().getValues().filter(function (row) {
                                                         return row[colCheckBox] == true;
                                                     });
  list.sort(function(x,y){
               if (x[colCheckBox - 1] > y[colCheckBox - 1]) {return -1;}
               if (x[colCheckBox - 1] < y[colCheckBox - 1]) {return 1;}
               return 0;            
            });
  return list;
}

function deduplicate(arr){
  return arr.filter(function (value, index, arr) { return arr.indexOf(value) === index});
}

function checkIndex(val, arr){
  for(i in arr){
    if(arr[i] == val) return i;
  }
}

function sendEmail(Email, Subject, html){
  var template = HtmlService.createTemplateFromFile("template");
  template.tableName = "Bảng Danh Sách";
  template.tableData = html;
  //Reset();
  try{
    GmailApp.sendEmail('', Subject, '', {
      //name: Name,
      htmlBody: template.evaluate().getContent(),
      bcc: Email,
      from: 'tuyetnt@vinhhungjsc.com'
    });   
  }
  catch(e){
    Browser.msgBox("Thất Bại!");
    Logger.log(e);
  }
}

function separateEmail(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var list = sortRange();
  var arr = [];
  var message = [];
  var name = [];
  var subject = [];
  for(i in list){
    var str = list[i][colCheckBox - 1].split(", ");
    arr = arr.concat(str);     
  }
  arr = deduplicate(arr);
  arr = arr.sort();
  for(i in arr){
    message[i] = [];
    name[i] = i;
    subject[i] = i;
  }
  for(i in list){
    var str = list[i][colCheckBox - 1].split(", ");
    var row = list[i];
    for(j in str){
      var index = checkIndex(str[j], arr);      
      message[index].push(row);  
      if(subject[index] == index) subject[index] = "Số báo giá: " + row[colSubject];
    }
  }
  for(i in arr){
    var html = [];
    html.push(sheet.getRange(rowHeader,1,1,lastValCol).getValues()[0]);
    for(t in message[i]){
      var x = message[i][t];
      html[t + 1] = [];
      for(var j = 0; j < lastValCol; j++){       
        html[t + 1].push(x[j]);
      }
    }
    //Browser.msgBox(html);
    sendEmail(arr[i], subject[i], html);
  }
}

function Reset(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getDataRange().getLastRow();
  sheet.getRange(rowHeader + 1,colCheckBox + 1,lastRow,1).setValue(false);
}

function onClickImg() { 
  //var html = HtmlService.createHtmlOutputFromFile('FormDataSend'); 
  //SpreadsheetApp.getUi()
//.showModalDialog(html, 'Send Email'); 
 // checkBox();  
  separateEmail();
  Browser.msgBox("Gửi thành công!");
}