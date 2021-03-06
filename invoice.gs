var ssa = SpreadsheetApp;
var ss = ssa.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var ui = ssa.getUi();
var drive = DriveApp;
var gmail = GmailApp;

var uiInvoice = HtmlService
.createHtmlOutputFromFile('invoiceMenu')
.setWidth(450)
.setHeight(400);

var invoiceMessage = HtmlService
.createHtmlOutputFromFile('invoiceMessage');

function invMenuMaker(){
//Set the variables 
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
//Find HTML to edit
var strHtml = uiInvoice.getContent();
var htmlMarker = strHtml.indexOf('</ol>');
var topHtml = strHtml.slice(0,htmlMarker);
var midHtml='';
var bottomHtml = strHtml.slice(htmlMarker);

//get the data
var data = tSheet.getDataRange().getValues();
var events = data[0];
events.shift();
events.shift();
events.shift();
events.shift();



//add each event to the middle of the html
for(i=0;i<events.length;i++){
midHtml=midHtml+'<li>'+events[i]+' -  <input type="text" value="$4.00" class= "prices" name="prices"></li>';
}
//stitch the html back together.
var fullHtml= topHtml+midHtml+bottomHtml;
Logger.log(fullHtml);
uiInvoice.clear();
uiInvoice.append(fullHtml);
ui.showModelessDialog(uiInvoice, 'Invoice Setup');
}

function createInvoice(form){
//Set the variables 
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var inSheet = ss.getSheetByName("test");
//get the data
var data = tSheet.getDataRange().getValues();
var topLine = data[0];
topLine.shift();
//Swap the data around to get a list of teams  
var tranData = transpose(data);
var teams = clearBlanks(tranData[1]);
teams.shift();
Logger.log(teams);
var numTeams = teams.length;
var folder= DriveApp.createFolder('4n6 Tourney Invoice');
  var fLink = folder.getUrl();
//Set the form data
prices = form.prices;
payCheck = form.payCheck;
send = form.send;
var tPrices = prices;
while(tPrices.length<topLine.length){
tPrices.unshift("");}




//Cycle through each team
  for(i=0;i<numTeams;i++){
  
    var thisTeamData = data[i*1+1];
    //Trim the fat
     thisTeamData.shift();
    //Create a new spreadsheet

    var invoice =ssa.create(thisTeamData[0]+"'s Invoice")
    var inSheet = invoice.getActiveSheet();
    var inId = invoice.getId();
    inId=drive.getFileById(inId);
    
    folder.addFile(inId);
    //Logger.log(topLine +'\n'+ tPrices);
var newData=[topLine,thisTeamData,tPrices];
 newData= transpose(newData);
 Logger.log(newData.length);
  inSheet.getRange(1,1,newData.length,newData[0].length).setValues(newData);
 var col = inSheet.getLastColumn()+1;
 var rows = inSheet.getLastRow()-3;
 inSheet.getRange(4,col,rows,1).setFormulaR1C1('=(R[0]C[-2]*R[0]C[-1])');
 inSheet.getRange(inSheet.getLastRow()+1,inSheet.getLastColumn()).setFormulaR1C1("=sum(R[-"+inSheet.getLastRow()+"]C[0]:R[-1]C[0])");
 inSheet.getRange(inSheet.getLastRow(),3).setValue("Total:");
 inSheet.getRange(inSheet.getLastRow(),inSheet.getLastColumn()).setNumberFormat("$#.##");
 
 
 /*var dSheetData = dataSheet.getDataRange().getValues();
 dSheetData = transpose(dSheetData);
 var row = 2;
 if(dSheetData.length <3){
 dSheetData.push("");
 }
 row = 2*1+dSheetData[2].length;
 var sheetLink = invoice.getUrl();
 dataSheet.getRange(row,3).setValue('=HYPERLINK("'+sheetLink+'","'+thisTeamData[0]+' Invoice")');*/
 var pCheckRow = inSheet.getLastRow()*1+1;
 var mCheckRow = inSheet.getRange(pCheckRow,1,1,4);
 mCheckRow.mergeAcross();
inSheet.getRange(pCheckRow,1).setValue("Please make checks payable to: " + payCheck);

}
 var folderLink = folder.getUrl();
 dataSheet.getRange("B5").setValue('=HYPERLINK("'+folderLink+'","Invoice Folder")')
Logger.log(numTeams);
ssa.flush();
if (send){messageMenu()}
}



function messageMenu(){
ui.showModelessDialog(invoiceMessage, 'Invoice Message');

}

function sendInvoice(form){

 var dataSheet=ss.getSheetByName("TM Data");

var folderCell = dataSheet.getRange("B5").getFormula();


var idMarker = folderCell.indexOf('folders/')+8;

var folderId= folderCell.slice(idMarker);

var quoteMarker= folderId.indexOf('"');
folderId=folderId.slice(0,quoteMarker);

var folder = drive.getFolderById(folderId);
if (folder.getName()!=null){
var files = folder.getFiles();
while (files.hasNext()){
var file = files.next();
var fileId = file.getId();
var blob = file.getBlob();
var sheet = ssa.openById(fileId);
var data = sheet.getDataRange().getValues();
var coachEmail = sheet.getRange("B3").getValue();
gmail.sendEmail(coachEmail, "Tournament Dues", " See the attached file for your tournament invoice. \n <i>This invoice has been generated by 4N6 Tournament Manager.</i>",{attachments: [file.getAs(MimeType.PDF)],name: 'Tournament Invoice'});
;}
dataSheet.getRange("C5").setValue("Sent.");


}
else {ui.alert("You must first create the invoices.")}
}

function testNow(){
var dataSheet=ss.getSheetByName("TM Data");

var folderCell = dataSheet.getRange("B5").getFormula();


var idMarker = folderCell.indexOf('folders/')+8;

var folderId= folderCell.slice(idMarker);
var quoteMarker= folderId.indexOf('"');
folderId=folderId.slice(0,quoteMarker);
Logger.log(folderId);
var folder = drive.getFolderById(folderId);

Logger.log(folder.getName());

}