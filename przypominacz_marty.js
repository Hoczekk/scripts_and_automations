var sheetName = "Formularz zatwierdzenia zmiany layoutu";
var dueDateCell = "C6";
var deadlineDateCell = "E6";
var versionCell = "F18";
var extraTextToAdd = "G31";
var startRow = 8; 
// var reminderColStart = 26;
var folderID = "1T0MEthwSTHAEDeQaQgoTlzl2DPUZmOw2";
var subject = "FORMULARZ ZATWIERDZANIA ZMIAN LAYOUTu";
var daysToRunTrigger = 365;
var myEmail = "kamil.hoczek.jv@valeo.com";
var mySubject = "FO0006 Daily email trigger";
var ss = SpreadsheetApp.getActiveSpreadsheet()
var sheet = ss.getSheetByName(sheetName)
var currentDate = new Date();
var deadlineDate = sheet.getRange(deadlineDateCell)
                        .getValue();
var timeAfterDeadline = currentDate.getTime() - deadlineDate.getTime();
var daysAfterDeadline = Math.trunc(timeAfterDeadline / (1000 * 3600 * 24));

// TODO LIST:
//////////////////////////////////////////////////////
// DueDate    C6                              ✓    //
// Deadline   E6                              ✓    //
// Email      C8-C18                          ✓    //
// Version    F18                             ✓    //
// Comment    G31                             ✓    //
// 1st Email                                  ✓    //
// 2nd Email                                  ✓    //
// Email based on E8-E18                       ✓    //
// Calculation between today and deadline      ✓    //
// Documentation and test cases               ✗    //
/////////////////////////////////////////////////////

function main() {
  var dueDate = sheet.getRange(dueDateCell)
                     .getValue();
  var currentDate = new Date();
  var expiryDate = function(){
    var result = new Date(dueDate);
    result.setDate(result.getDate() + daysToRunTrigger)
    return result;
  };
  
  if(currentDate <= dueDate){
    return;
  }else if(currentDate > expiryDate()){
    var cancelTrigger = cancelTimeTrigger();
    return;
  };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var dueDateFormatted = Utilities.formatDate(dueDate, "GMT+2", "EEE, dd MMM yyyy");
  //var daysAfterDeadlineFormatted = Utilities.formatDate(daysAfterDeadline, "GMT+2", "EEE, dd MMM yyyy");
  var range = sheet.getRange(startRow,1,lastRow - startRow+1, lastCol); 
  var rangeVals = range.getValues();
  var badEmailList = [];
  var reminder = [];


 for(row in rangeVals){

    if(!rangeVals[row][4] && daysAfterDeadline >= 1){
        var badEmailAftedDeadline = sendDeadlineEmail(rangeVals[row], ss.getUrl, daysAfterDeadline);
        badEmailList.push(badEmailAftedDeadline);
      }else if (!rangeVals[row][4] && daysAfterDeadline < 1){
      var badEmail = sendEmail(rangeVals[row], ss.getUrl, dueDateFormatted);
      badEmailList.push(badEmail);
    }else{
          var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
          file.getParents().next().removeFile(file);
          DriveApp.getFolderById(folderID).addFile(file);
    };   
  }; 
  //var reminderRange = sheet.getRange(startRow, reminderColStart,lastRow - startRow+1, 2);
  //reminderRange.setValues(reminder);
  var complete = sendDailySummaryEmail(badEmailList); 
};

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Powiadomienia")
  .addItem("Wyślij powiadomienie", "main")
  .addItem("Uruchom zegar powiadomień","startTimeTrigger")   
  .addItem("Zatrzymaj zegar powiadomień","cancelTimeTrigger")
  .addToUi();
};

// -------------------------------------------------------------------------------------------------------

function sendDeadlineEmail(overDueStaff, sheetURL, dateDue){
  var staff = {
    "name": overDueStaff[1],
    "email": overDueStaff[2],
    "dueDate": dateDue
//    "daysOverdue": overDueStaff[4]
  };
//  staff.daysOverdue = (staff.daysOverdue > 1) ? staff.daysOverdue + " days" : staff.daysOverdue + " day";

// Deadline calculation
var deadlineDate = sheet.getRange(deadlineDateCell)
                        .getValue();
var versionOfDoc = sheet.getRange(versionCell)
                        .getValue();
var extraComment = sheet.getRange(extraTextToAdd)
                        .getValue();

var dateString = deadlineDate;
var versionString = versionOfDoc;
var emailComment = extraComment;

  var url = '';
  url += ss.getUrl();
  url += '#grid=';
  url += sheet.getSheetId();
  var body = HtmlService.createTemplateFromFile("email2");
  body.name = staff.name;
  body.due = dateString;
  body.version = versionString;
  body.comment = emailComment;
  body.overdue = staff.daysOverdue;
  body.reportName = sheetName;
  body.urlToSheet = url;
  body.deadlineDays = daysAfterDeadline;
  try{
      MailApp.sendEmail({
      to: staff.email,
      subject: subject,
        htmlBody: body.evaluate().getContent(), 
                      });
  }
  catch(error){
    return staff.email;
  }
};  

function sendEmail(overDueStaff, sheetURL, dateDue){
  var staff = {
    "name": overDueStaff[1],
    "email": overDueStaff[2],
    "dueDate": dateDue
//    "daysOverdue": overDueStaff[4]
  };
//  staff.daysOverdue = (staff.daysOverdue > 1) ? staff.daysOverdue + " days" : staff.daysOverdue + " day";

// Deadline calculation
var deadlineDate = sheet.getRange(deadlineDateCell)
                        .getValue();
var versionOfDoc = sheet.getRange(versionCell)
                        .getValue();
var extraComment = sheet.getRange(extraTextToAdd)
                        .getValue();

var dateString = deadlineDate;
var versionString = versionOfDoc;
var emailComment = extraComment;

  var url = '';
  url += ss.getUrl();
  url += '#grid=';
  url += sheet.getSheetId();
  var body = HtmlService.createTemplateFromFile("email");
  body.name = staff.name;
  body.due = dateString;
  body.version = versionString;
  body.comment = emailComment;
  body.overdue = staff.daysOverdue;
  body.reportName = sheetName;
  body.urlToSheet = url;
  body.deadlineDays = daysAfterDeadline;
  try{
      MailApp.sendEmail({
      to: staff.email,
      subject: subject,
        htmlBody: body.evaluate().getContent(), 
                      });
  }
  catch(error){
    return staff.email;
  }
};  


function sendDailySummaryEmail(badEmailList){
  MailApp.sendEmail({
    to: myEmail,
    subject: mySubject, 
    htmlBody: "<p> Reminder emails send for today </p>" + 
              "<p> The following emails could not be sent:</p>" +
              badEmailList.join("<br>")
  });
};

function getSheetUrl() {
  var url = '';
  url += ss.getUrl();
  url += '#grid=';
  url += sheet.getSheetId();
  return url;
}
