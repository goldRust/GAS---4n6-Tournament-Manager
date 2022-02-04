


function logUse() {
var ssa = SpreadsheetApp;
//https://docs.google.com/spreadsheets/d/1Y2AiCdliM56blHTggx8EQIp3o3tQWJE5L6vwlrdFPNM/edit?usp=sharing
var ssLog = ssa.openById("1Y2AiCdliM56blHTggx8EQIp3o3tQWJE5L6vwlrdFPNM");
//var ssLog = ssa.openById('1Y2AiCdliM56blHTggx8EQIp3o3tQWJE5L6vwlrdFPNM');
var sheet = ssLog.getActiveSheet();

  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  
  var nextRow = sheet.getLastRow()*1+1;
  var user = Session.getActiveUser().getEmail();
 var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  var yyyy = today.getYear();

  
 today = mm+'/'+dd+'/'+yyyy;

  
  sheet.getRange(nextRow,1).setValue(user);
  sheet.getRange(nextRow,2).setValue(today);
}
