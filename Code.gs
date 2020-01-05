var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

function deleteSelectRow(){
  var response = SpreadsheetApp.getUi().alert('詢問', '請問是否要刪除?', SpreadsheetApp.getUi().ButtonSet.YES_NO);
  if (response == SpreadsheetApp.getUi().Button.YES){
     sheet.deleteRow(sheet.getActiveRange().getRow());
  }
}

function onEdit(e) {
  // Set a comment on the edited cell to indicate when it was changed.
  var range = e.range;
  toUpdateWorkByRange(sheet,range);
  if (sheet.getName().equals("工作清單")) {
    if (range.getA1Notation().substring(0,1).equals("U")){
      //實際完成日期當填入時，會自動將前面的進度，列為完成
      sheet.getRange("C" + range.getRow()).setValue("完成");
    }
  }
}

function toUpdateWorkByRange(sheet,range){

  if (sheet.getName().equals("工作清單")) {

    changeHourData(sheet,range,"AC");
    changeHourData(sheet,range,"AD");
    changeHourData(sheet,range,"AE");
    changeHourData(sheet,range,"AF");
    changeHourData(sheet,range,"AG");
    changeHourData(sheet,range,"AH");
    changeHourData(sheet,range,"AI");
    changeHourData(sheet,range,"AJ");
    
    updateLastUpdateDateColumn(range);
    
  } else if (sheet.getName().equals("時數對應表")){
    if (sheet.getRange(4, range.getColumn()).getValue().equals("處利內容說明")){
      sheet.getRange(range.getRow(), range.getColumn() + 1).setValue(new Date());
    }
  }
}

//異動更新時間資料
function updateLastUpdateDateColumn(range){
  //尋找欄位
  for(var i = 1; i <= sheet.getLastColumn(); i ++) {
    if (sheet.getRange(2, i).getValue().equals("更新時間")) {
      sheet.getRange(range.getRow(), i).setValue(new Date());
      break;
    }
  }
}

function changeHourData(sheet,range,needColumn){
  var row = range.getRow();
 
  if (sheet.getRange(row, range.getColumn() + 21).getA1Notation().equals(sheet.getRange(needColumn + row).getA1Notation())
      &&  sheet.getRange(2,range.getColumn()).getValue().indexOf("時數") > -1){
    
    var currentHourCell = sheet.getRange(needColumn + row);
    var currentHour = currentHourCell.getValue();
    
     //加總
    currentHourCell.setValue(currentHour + range.getValue());
    //還原
    var currentChangeCell = sheet.getRange(row,range.getColumn());
    currentChangeCell.setValue(0);
  }
  
} 


function insertDriveComment(fileId, comment, context) {
  var driveComment = {
    content: comment,
    context: {
      type: 'text/html',
      value: context
    }
  };
  Drive.Comments.insert(driveComment, fileId);  
}

//新增資訊申辦單
function insertNormalProblem(){
  //塞入目前sheet第4列 (其實是第2列，之所以不放在第1列，是為了屬性套用方便) 
 
  sheet.getRange('3:3').activate();
  sheet.insertRowsBefore(sheet.getActiveRange().getRow(), 1);
  sheet.getActiveRange().offset(0, 0, 1, sheet.getActiveRange().getNumColumns()).activate();
  sheet.getRange('B3').activate();
  sheet.getCurrentCell().setValue('Ⅹ');
  sheet.getRange('C3').activate();
  sheet.getCurrentCell().setValue('尚未');
  sheet.getRange('H3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('I3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('J3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('K3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('L3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('M3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('N3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('O3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('Z3').activate();
  sheet.getRange('Z4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('AA3').activate();
  sheet.getCurrentCell().setValue('Ⅹ');
  sheet.getRange('AK3').activate();
  sheet.getRange('AK4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('AT3').activate();
  sheet.getRange('AT4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('AV3').activate();
  sheet.getRange('AV4:BF4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

//新增線上問題
function insertOnlineProblem(){
  //先取目前線上問題編號再加1
  var settingSheet = ss.getSheetByName("設定");
  var num = settingSheet.getRange(2, 5).getValue();
  var onlineNum = "LINE-" +  paddingLeft((num + 1).toString(),6);
  //取號後塞回
  settingSheet.getRange(2, 5).setValue((num + 1));
  
  sheet.getRange('3:3').activate();
  sheet.insertRowsBefore(sheet.getActiveRange().getRow(), 1);
  sheet.getActiveRange().offset(0, 0, 1, sheet.getActiveRange().getNumColumns()).activate();
  sheet.getRange('A3').activate();
  sheet.getCurrentCell().setValue(onlineNum);
  sheet.getRange('B3').activate();
  sheet.getCurrentCell().setValue('√');
  sheet.getRange('C3').activate();
  sheet.getCurrentCell().setValue('尚未');
  sheet.getRange('D3').activate();
  sheet.getCurrentCell().setValue('11.線上問題');
  sheet.getRange('H3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('I3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('J3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('K3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('L3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('M3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('N3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('O3').activate();
  sheet.getCurrentCell().setValue('0');
  sheet.getRange('Z3').activate();
  sheet.getRange('Z4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('AA3').activate();
  sheet.getCurrentCell().setValue('Ⅹ');
  sheet.getRange('AK3').activate();
  sheet.getRange('AK4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('AT3').activate();
  sheet.getRange('AT4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('AV3').activate();
  sheet.getRange('AV4:BF4').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function sendEMail(){
  var emailAddress =  sheet.getActiveCell().getValue();
  var subject = "Google 表單資料請記得更新";
  //申辦單內容
  
  var rangeArray = sheet.getSelection().getActiveRangeList().getRanges();
  for (var i = 0; i < rangeArray.length; i ++) {
    var data = rangeArray[i].getValues();
    for (var j = 0 ; j < data.length; j ++) {
      var receiver = data[j];
      var startRow = rangeArray[i].getRow() + j; 
      var issueNo = sheet.getRange("A" + startRow).getValue();
      var mainMessage = sheet.getRange("G" + startRow).getValue(); 
      var templ = HtmlService .createTemplateFromFile('updateIssueDataMail');
      
      templ.candidate =  {
        name: receiver.toString().substring(0,3),
        issueNo : issueNo,
        issueContent: mainMessage 
      };
      //將candidate內的json資料寫入html content
      var message = templ.evaluate().getContent();
      
      MailApp.sendEmail({
        to: receiver.toString(),
        subject: issueNo + "," + subject,
        cc:"peterchen.ipo@gmail.com",
        htmlBody: message
      });
    }
  }
  SpreadsheetApp.getUi().alert("寄信成功");
}


// The onOpen function is executed automatically every time a Spreadsheet is loaded
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
  // executed.
  menuEntries.push({name: "再想想", functionName: "function1"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "再想想2", functionName: "function2"});

  ss.addMenu("客製化功能頁by Peter", menuEntries);
  
  

}

//字串補0
function paddingLeft(str,lenght){
  if(str.length >= lenght)
    return str;
  else
    return paddingLeft("0" +str,lenght);
}

//時數對應表-新增一列
function insertARow(){
  var spreadsheet = ss.getSheetByName("時數對應表");  
  spreadsheet.getRange('5:5').activate();
  spreadsheet.insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('C5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('D5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('F5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('G5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('H5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('I5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('J5').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('K5').activate();
}

function insertHourData(){
  var spreadsheet = ss.getSheetByName("時數對應表");  
  var rangeArray = spreadsheet.getActiveRangeList().getRanges();
  var allowInsert = false;
  for (var i =0 ; i < rangeArray.length; i ++) {
    var data = rangeArray[i].getValues();
    //基本檢查
    if ("".equals (spreadsheet.getRange(rangeArray[i].getRow(), 1).getValue())){
      SpreadsheetApp.getUi().alert("單號沒寫!");
      return;
    }
    if ("".equals (spreadsheet.getRange(rangeArray[i].getRow(), 11).getValue())){
      SpreadsheetApp.getUi().alert("處理內容說明沒寫!");
      return;
    }
    //取時數相加
    var i3 = spreadsheet.getRange(rangeArray[i].getRow(), 3).getValue();
    var i4 = spreadsheet.getRange(rangeArray[i].getRow(), 4).getValue();
    var i5 = spreadsheet.getRange(rangeArray[i].getRow(), 5).getValue();
    var i6 = spreadsheet.getRange(rangeArray[i].getRow(), 6).getValue();
    var i7 = spreadsheet.getRange(rangeArray[i].getRow(), 7).getValue();
    var i8 = spreadsheet.getRange(rangeArray[i].getRow(), 8).getValue();
    var i9 = spreadsheet.getRange(rangeArray[i].getRow(), 9).getValue();
    var i10 = spreadsheet.getRange(rangeArray[i].getRow(), 10).getValue();
    
    if (i3 + i4 + i5 + i6 + i7 + i8 + i9 + i10 ==0) {
      SpreadsheetApp.getUi().alert("時數沒寫!");
      return;
    }
    
    var issueNo = data[i][0];
    var name = data[i][1];
    var h1 = data[i][2];
    var h2 = data[i][3];
    var h3 = data[i][4]
    var h4 = data[i][5];
    var h5 = data[i][6];
    var h6 = data[i][7];
    var h7 = data[i][8];
    var h8 = data[i][9];
    var process = data[i][10];

    inertDetailBoard(issueNo,name,h1,h2,h3,h4,h5,h6,h7,h8,process);
  }
  SpreadsheetApp.getUi().alert("完成!");
}

//新增至工作清單
function inertDetailBoard(issueNo,name,h1,h2,h3,h4,h5,h6,h7,h8,process){
  var sheet = ss.getSheetByName("工作清單");
  ss.setActiveSheet(sheet);
  
  var getRowId = findInColumn("A",issueNo,sheet);
  if (getRowId != -1) {
   handelHourUpdate(sheet,getRowId,h1,h2,h3,h4,h5,h6,h7,h8);
    
    var preProcess = sheet.getRange("P" + getRowId).getValue();
    if (preProcess.length > 0) {
       sheet.getRange("P" + getRowId).setValue(  sheet.getRange("P" + getRowId).getValue() + "\r\n" + handelProcess(process,h1,h2,h3,h4,h5,h6,h7,h8) );
    } else {
       sheet.getRange("P" + getRowId).setValue(  handelProcess(process,h1,h2,h3,h4,h5,h6,h7,h8) );
    }

    var needAddNoteCell = sheet.getRange("A" + getRowId );
    var note = needAddNoteCell.getNote();
    var newNote =name + "\r\t" + getUpdateStr(h1,h2,h3,h4,h5,h6,h7,h8) + "\r\t"
    +  Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm:ss");
    needAddNoteCell.setNote(newNote + "\r\n" + note);
   
    
  } else {
    SpreadsheetApp.getUi().alert(issueNo + ",單號可能寫錯了，目前沒有這筆單號，請重新確認!");
  }
}


function handelHourUpdate(sheet,getRowId,h1,h2,h3,h4,h5,h6,h7,h8){
  subhandelHourUpdate(sheet,"H",getRowId,h1);
  subhandelHourUpdate(sheet,"I",getRowId,h2);
  subhandelHourUpdate(sheet,"J",getRowId,h3);
  subhandelHourUpdate(sheet,"K",getRowId,h4);
  subhandelHourUpdate(sheet,"L",getRowId,h5);
  subhandelHourUpdate(sheet,"M",getRowId,h6);
  subhandelHourUpdate(sheet,"N",getRowId,h7);
  subhandelHourUpdate(sheet,"O",getRowId,h8);
}

function subhandelHourUpdate(sheet,coulmnName,rowId,data){
  if (data > 0) {
    sheet.getRange(coulmnName + rowId).setValue(data); 
    var range = sheet.getRange(coulmnName + rowId);
    toUpdateWorkByRange(sheet,range);
  }
}

function handelProcess(process,h1,h2,h3,h4,h5,h6,h7,h8){
  var data = "" ;
  
  if (h1 > 0) {
    data += "需求分析" + h1 + "小時，"
  }
  if (h2 > 0) {
    data += "需求規劃" + h2 + "小時，"
  }
  if (h3 > 0) {
    data += "需求開發" + h3 + "小時，"
  }
  if (h4 > 0) {
    data += "功能驗測" + h4 + "小時，"
  }
  if (h5 > 0) {
    data += "系統文件" + h5 + "小時，"
  }
  if (h6 > 0) {
    data += "建構文件" + h6 + "小時，"
  }
  if (h7 > 0) {
    data += "資料驗證" + h7 + "小時，"
  }
  if (h8 > 0) {
    data += "資訊驗測" + h8 + "小時，"
  }
  data = data.substring(0,data.length -1);
  return process + "，" + data;
}

function getUpdateStr(h1,h2,h3,h4,h5,h6,h7,h8){
  var data = "異動" ;
  
  if (h1 > 0) {
    data += "需求分析時數、"
  }
  if (h2 > 0) {
    data += "需求規劃時數、"
  }
  if (h3 > 0) {
    data += "需求開發時數、"
  }
  if (h4 > 0) {
    data += "功能驗測時數、"
  }
  if (h5 > 0) {
    data += "系統文件時數、"
  }
  if (h6 > 0) {
    data += "建構文件時數、"
  }
  if (h7 > 0) {
    data += "資料驗證時數、"
  }
  if (h8 > 0) {
    data += "資訊驗測時數、"
  }
  data = data.substring(0,data.length -1);
  return data;
}


function findInColumn(column, data,sheet) {
  var column = sheet.getRange(column + ":" + column);  // like A:A
  var values = column.getValues(); 
  var row = 0;
  while ( values[row] && values[row][0] !== data ) {
    row++;
  }
  if (values[row][0] === data) 
    return row+1;
  else 
    return -1;
}

function findInRow(data,sheet) {
  var rows  = sheet.getDataRange.getValues(); 
  for (var r=0; r<rows.length; r++) { 
    if ( rows[r].join("#").indexOf(data) !== -1 ) {
      return r+1;
    }
  }
  return -1;
}
