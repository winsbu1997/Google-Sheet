var mainList;
var keyList;
var subList;
var beginRowMain = 4;
var arrMainHeader = [21,22,23,24,25,26,27,28,29,30,31,32];
var arrSubHeader = [5,6,9,10,11,12,13,14,15,16,17,21];

var ss = SpreadsheetApp.getActiveSpreadsheet();
// update subSheet => mainSheet
var subSheet = ss.getSheetByName("N.M.Tuấn");
var mainSheet = ss.getSheetByName("BGTT");
var key = subSheet.getRange(3, 2).getValue() + subSheet.getRange(3, 4).getValue();
var nsx = subSheet.getRange(3, 6).getValue();
// Gán mảng tạo bởi 4 trường keys + Row cho cả 2 sheet 
function AssignKeyRow(list, index, keyCol1, keyCol2, keyCol3, keyCol4){
  var count = index;
  var newList = [];
  var newSTT = [];
  for(i in list){
    var row = list[i];
    var s = "";
    newSTT.push(count++);    
    for(j in row){      
      if(j == keyCol1 || j == keyCol2 || j == keyCol3 || j == keyCol4){
        s += row[j];
      }
    }
    newList.push(s);
  }
  return [newList, newSTT];
}
function Init() {
  // kiem tra du lieu den row trong sheet BGTT
  var lastRowMain = 1000;
  // kiem tra du lieu den row trong sheet N.M.Tuấn
  var lastRowSub = 1000;
      //subSheet.getDataRange().getLastRow();
  
  mainList = mainSheet.getRange(beginRowMain, 4, lastRowMain - beginRowMain + 1, 6).getValues().filter(function (row) {
                                                         return row[0] != "";
                                                     });
  keyList = subSheet.getRange(11, 2, lastRowSub - 10, 1).getValues().filter(function (row) {
                                                         return row != "";
                                                     }); 
  subList = subSheet.getRange(11, 5, lastRowSub - 10, 17).getValues().filter(function (row) {
                                                         return row[0] != "";
                                                   }); 
}

// khi nút update được click
function EventClickUpdate(){
  if(subSheet.getRange(5, 4).getValue() != "Đã duyệt"){    
    Init();
    for(i in keyList){
      keyList[i] = key + keyList[i] + nsx;
    }
    //Browser.msgBox(keyList);
    var [mainValues, mainRows] = AssignKeyRow(mainList, beginRowMain, 0,1,2,5);
    for(i in keyList){      
      var index = mainValues.indexOf(keyList[i]);
      if(index == -1){
        Browser.msgBox("Không tìm thấy với Key là: " + keyList[i]);
      }
      else{
        for(j in arrMainHeader){    
          mainSheet.getRange(mainRows[index],arrMainHeader[j]).setValue(subList[i][arrSubHeader[j]-5]);
        }
      }
    } 
  }
  else{
    Browser.msgBox("Bạn không được cập nhập dữ liệu khi đang khóa");
  }
}
