var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

var activeCell = ss.getActiveRange();
var editRow = activeCell.getRow();
var editCol = activeCell.getColumn();
var activeCellValue = activeCell.getValue();

var dbSheet = ss.getSheetByName('Phụ');
var lastCol = dbSheet.getLastColumn();
var lastRow = dbSheet.getLastRow();

function takeArrCount(key){
  var list = dbSheet.getRange(4,4,editRow,2).getValues().filter(function(row){
    return row[0] == key;
  });
  var count = parseInt(list[0][1]);
  var arr = [];
  for(var i = 1; i <= count; i ++){
    arr.push(i);
  }
  return arr;
}
function takeArrProducts(key){
  var arr = [];
  var list = dbSheet.getRange(4,8,editRow,2).getValues().filter(function(row){
    return row[0] == key;
  });
  for(i in list){
    arr.push(list[i][1]);
  }
  return arr;
}

function protect(col){
  var lastRange = activeCell.getLastRow();
  if(activeCellValue != ""){  
    sheet.getRange(editRow, col + 1, lastRange - editRow + 1, 1).setValue("Gửi mail");
    for(var j = editRow; j <= lastRange; j++){ 
      var range1 = sheet.getRange(String(j),1, 1, col);     
      var protection = range1.protect().setDescription(String(j));
      protection.removeEditors(protection.getEditors());
      protection.addEditor("tuyetnt@vinhhungjsc.com");
    }
  }
  else{   
    sheet.getRange(editRow, col + 1, lastRange - editRow + 1, 1).setValue("");
    for(var j = editRow; j <= lastRange; j++){  
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < protections.length; i++) {
        if (protections[i].getDescription() == String(j)) {
          protections[i].remove();
        }
      }  
    }
  }
}
function ProtectOnEdit() { 
  if(sheet.getSheetName() == "YCMH từ YCBG"){
    
    if (editCol == 2 && editRow > 3) {
      activeCell.offset(0, 1).clearContent().clearDataValidations();
      activeCell.offset(0, 2).clearContent().clearDataValidations();
      
      var select1 = takeArrCount(activeCellValue);
      var rule1 = SpreadsheetApp.newDataValidation().requireValueInList(select1).build();
      activeCell.offset(0, 1).setDataValidation(rule1);
      
      var select2 = takeArrProducts(activeCellValue);
      var rule2 = SpreadsheetApp.newDataValidation().requireValueInList(select2).build();
      activeCell.offset(0, 2).setDataValidation(rule2);
    }
    
    if(editCol == 7 && editRow > 3){  
      protect(7);
    }   
  }
  else if(sheet.getSheetName() == "YCMH không YCBG"){
    if(editCol == 20 && editRow > 3){  
      protect(20);
    }   
  }
}
