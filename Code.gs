var ssa = SpreadsheetApp;
var ss = ssa.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var ui = ssa.getUi();

//ui attached html files
var uiSetUp = HtmlService
.createHtmlOutputFromFile('setUp')
.setWidth(250)
.setHeight(300);

var uiDonate = HtmlService
.createHtmlOutputFromFile('donate')
.setWidth (250)
.setHeight(300)
.setTitle('Donate');

var uiHelp = HtmlService
.createHtmlOutputFromFile('help')
.setWidth (600)
.setHeight(400)
.setTitle('Help');

var tabsMenu = HtmlService
.createHtmlOutputFromFile('tabsMenu')
.setWidth(350)
.setHeight(150);





function onOpen(e){
//Create the menu  
  
   if (e && e.authMode == ScriptApp.AuthMode.NONE) {
         var menu =  [{name: "Set Up", functionName: "setUp"}]
     // Add a normal menu item (works in all authorization modes).
    ss.addMenu("Tournament Manager", menu);
   } else {
  var menu =  [{name: "Set Up", functionName: "setUp"},
               {name: "Code Contestants", functionName: "repGen"}, {name: "Assign Rooms", functionName: "improvedAssign"},{name: "Make Docs",functionName:"makeDocs2"},{name: "Create Invoices", functionName: "invMenuMaker"},{name:"Send Invoice", functionName: "messageMenu"},{name: "Create Tab Sheet", functionName: "tabMenuUi"},{name:"Help",functionName:"help"},{name: "Donate",functionName:"donate"}];
 
  ss.addMenu("Tournament Manager", menu);

    // logUse();
     }
 
  
}

function setUp(){
var demSheet = ["Teams","Events","Rooms","Contestants"];
var makeSheet = ss.getSheetByName(demSheet);


ui.showSidebar(uiSetUp);


}

function clearBlanks(array){
     var newArray = [];
     for(i=0;i<array.length;i++){
            if(array[i]!=""){newArray.push(array[i]);}
                                     }
  return newArray;};


function transpose(a){
 
  
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

function oneToTwo(array){
    var newArray = [];
  
  for(i=0;i<array.length;i++){
var arrayItem = [];
  arrayItem.push(array[i]);
  newArray.push(arrayItem);
  
  } 
  return newArray;
}
function twoToOne(array){
  var newArray = [];
  for (i=0;i<array.length;i++){
    var arrayElement = array[i][0];
   
    newArray.push(arrayElement);
    
                  }
  return newArray;
}


function repGen(){
  
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
 
 //clears data that remains in other sheets
 cSheet.clear();
 rSheet.clear();
 
 var tEvents2 = eSheet.getDataRange().getValues();
  
  var tEvents3 = transpose(tEvents2);
  tEvents3[0].shift();

  var tEvents = clearBlanks(tEvents3[0]);
  var nEvents= tEvents.length;
  

  //Compiles numbers for each event by team
 var data = tSheet.getDataRange().getValues();
  var allTeams = [];
  var allEvents = [];

  
  var numTeams = tSheet.getLastRow() - 1;
  for(i=0;i<numTeams;i++){
 var team = data[1+1*i];
  for(j=0;j<4;j++){
    team.shift()}
    
    if(data[i*1+1][0] != ''){
      allTeams.push(team)}
  }
  
   var numEvents = allTeams[0].length;
  
  for(i=0;i<numEvents;i++){

    
    var thisEvent = [];
    for (j=0;j<allTeams.length;j++){
      var thisTeam = allTeams[j];
      thisEvent.push(thisTeam[i]);}
  
    allEvents.push(thisEvent);
  }
  
  
   eSheet.getRange("B1").setValue("Totals");
  //sets math for events sheet
  var sums =[];
 
  var letters = "E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,"
  var letter=letters.split(",");
  for(i=0;i<numEvents;i++){
    sums.push("=sum(Teams!"+letter[i]+"2:"+letter[i]+"40)");
    
   
  }
  Logger.log(sums);
  var sumX = oneToTwo(sums);
  Logger.log(sumX);
 
   
  eSheet.getRange(2,2,numEvents).setValues(sumX);

  eSheet.getRange("C1").setValue("Rooms Needed:");
  eSheet.getRange(2,3,numEvents).setFormulaR1C1("=CEILING(R[0]C[-1]/8)")

  eSheet.getRange("D1").setValue("Avg. Per Room");
  eSheet.getRange(2,4,numEvents).setNumberFormat("0.00").setFormulaR1C1("=R[0]C[-2]/R[0]C[-1]")
  var glr = eSheet.getLastRow();
  eSheet.getRange(eSheet.getLastRow()+1,2).setValue("Total Rooms:");
  eSheet.getRange(eSheet.getLastRow(),3).setFormulaR1C1('=sum(R[-'+glr+']C[0]:R[-1]C[0])');
   var e=0;
  
 
  var aEvents = [];

  
  //cycle through each event
 codeCon();
  
  
}

function shuffle(array) {
  var currentIndex = array.length, temporaryValue, randomIndex;

  // While there remain elements to shuffle...
  while (0 !== currentIndex) {

    // Pick a remaining element...
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex -= 1;

    // And swap it with the current element.
    temporaryValue = array[currentIndex];
    array[currentIndex] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }

  return array;
}

function help(){
  ui.showModalDialog(uiHelp,"Help");
}


function tabMenuUi(){
  ui.showModalDialog(tabsMenu,"Tab Room Selection");
}
function rTabsMenu(tab){
  Logger.log(tab.tTab);
 var whichTab = tab.tTab;
  if (whichTab == "makeTabs1"){
    makeTabs1();}
  else if(whichTab == "makeTabs2"){
    makeTabs2();}
  else if(whichTab== "makeTabs3"){
    makeTabs3();}
  
}


function setEvents(form){
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  
tEvents = form.foo;
var nEvents =tEvents.length;
 eSheet.getRange("A1").setValue("Events");
  eSheet.getRange("B1").setValue("Totals");
  tSheet.getRange("A1").setValue("Team #");
  tSheet.getRange("B1").setValue("Team");
   tSheet.getRange("C1").setValue("Coach");
   tSheet.getRange("D1").setValue("email");
  var y=2;
  var e=0;
  var m=5;
  while(e<nEvents){
  eSheet.getRange(y,1).setValue([[tEvents[e]]]);
  tSheet.getRange(1,m).setValue([[tEvents[e]]]);
  e++;
  y++;
  m++;
  }
}



function makeForm(form){
tEvents= form.foo;
var nEvents = tEvents.length;
tName = form.tName;
tDate= form.tDate;
maxE= form.maxE;
var entryForm = FormApp.create(tName +" "+ tDate + " entry form");
entryForm.addTextItem().setTitle("Team:");
entryForm.addTextItem().setTitle("Coach:");
entryForm.addTextItem().setTitle("Coach Email:");
var e=0;
while(e<nEvents){
var item = entryForm.addScaleItem();
item.setTitle(tEvents[e]);
item.setBounds(0,maxE);
e=e+1;

}
  entryForm.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
}
// Replace this with ID of your template document.
var TEMPLATE_ID = '1skkuhjPZZBGN_uxINR7DCdYdSlEDwJiN5hU4CgwMqdU';


function createDoc(start,end,folder) {
  
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");

  if (TEMPLATE_ID === '') {
    
    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }

  // Set up the docs and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(folder);
    var  copyId = copyFile.getId();
     var copyDoc = DocumentApp.openById(copyId);
      var copyBody = copyDoc.getActiveSection();
      
   
      var numberOfRows = rSheet.getLastRow();
      
      var activeRowIndex = rSheet.getActiveRange().getRowIndex();
      var activeColumnIndex= rSheet.getActiveRange().getColumnIndex();
      var numberOfColumns = (end-start);
     // var activeRow = rSheet.getRange(activeRowIndex, 1, 1, numberOfColumns).getValues();
   var  nRooms = rSheet.getRange(1,start +1).getValue();
 //     var headerRow = rSheet.getRange(1, 1, 1, numberOfColumns).getValues();
     
      var rowIndex=2;
      var startPoint= (rowIndex,start);
     var  endPoint=(numberOfRows,end) ; 
  
 
  // Replace the keys with the spreadsheet values
 
 
    var event =  rSheet.getRange(1,start).getValue();
  var array2=[]
    var array1 = rSheet.getRange(rowIndex,start,numberOfRows,(end-start)).getValues();
  var filledRows = (array1.filter(String).length)+nRooms;
    var text = array1.filter(String);
 
  copyBody.replaceText('<<Event>>',event);
  var table = copyBody.appendTable(text);

  
  
  

   

 copyDoc.setName(event + " Room Schedule");
  copyDoc.saveAndClose();
  
  
} 

function makeDocs2(){
  
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  var folderIterator = DriveApp.getFoldersByName("4n6 Room Schedules");
  if (folderIterator.hasNext() ){
   var folder = folderIterator.next();
  }else{
     var folder= DriveApp.createFolder('4n6 Room Schedules');
  }
  var fLink = folder.getUrl();
 var allEvents= rSheet.getLastColumn();
  
   var lastDataRow = dataSheet.getLastRow();
    
  var eInd = dataSheet.getRange(2,1,lastDataRow).getValues();   eInd = clearBlanks(eInd);     eInd = twoToOne(eInd); 
  
  for(i=0;i<eInd.length-1;i++){
    var start = eInd[i];
    var end = eInd[i+1] - 1;
    createDoc(start, end, folder);}  
   SpreadsheetApp.getUi().alert('Room Schedules created in your 4n6 Room Schedules folder.');
   dataSheet.getRange(3,2).setValue('= HYPERLINK("'+fLink+'","4n6 Room Folder")');
   dataSheet.autoResizeColumn(2);
}

function makeTabs1(){
  
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");

 var folders= DriveApp.getFoldersByName('4n6 Room Schedules');
  var folder = folders[0];
  var template = "1T5J_k99VGB3I8FL_e7HVDjcXi9Z4Uv0oAKfH2ZHpab4";
//var tabSheet = ssa.create('TM TABS');
  var source = ssa.openById(template);
  var tabSheet =source.copy('TM TABS');
  
  var tabSheetId = tabSheet.getId();
  var tabSheetUrl = tabSheet.getUrl();
      var tEvents2 = eSheet.getDataRange().getValues();   var tEvents3 = transpose(tEvents2);   tEvents3[0].shift();   tEvents3[2].shift();      var tEvents = clearBlanks(tEvents3[0]);

  
  var sourceSheet = source.getSheetByName("Event");
tabSheet.insertSheet();
    tabSheet.deleteSheet(tabSheet.getSheetByName('SWEEPS'));
  tabSheet.deleteSheet(tabSheet.getSheetByName('Event'));


  for(q=0,r=2;q<tEvents.length;q++,r++){
   sourceSheet.copyTo(tabSheet);
    var theseEvents = tabSheet.getSheets();
    var allCon = cSheet.getLastRow();
    var contestants = cSheet.getRange(2,q+1,allCon,1).getValues();
    theseEvents[q+1].setName(tEvents[q]);
    theseEvents[q+1].getRange(3,2,allCon,1).setValues(contestants);
    theseEvents[0].hideSheet();
  
    
    // This is where I put room numbers on the tab sheet  
  var thisSheet=theseEvents[q+1]; 
    var dataRange = thisSheet.getRange(3,2,allCon,1).getValues();
    var lastDataRow = 1+1*dataSheet.getLastRow();
   var eInd = dataSheet.getRange(2,1,lastDataRow).getValues();   eInd = clearBlanks(eInd);     eInd = twoToOne(eInd);  
   
    var lastRow = rSheet.getLastRow();
    var lastColumn = rSheet.getLastColumn();
    var start = eInd[q];
    var end = eInd[q+1] - 1;
    var numRooms = end - start;
    var roomNames=rSheet.getRange(2,2,1,lastColumn).getDisplayValues();  
      
      for(j=0;j<numRooms;j++){
        var startCol = start*1+j;
      
       var roomSchedule = rSheet.getRange(3,startCol,lastRow,1).getDisplayValues();
       
        var strRounds = roomSchedule.join();
        var rounds = strRounds.split(",,");
        var roundOne= rounds[0].split(",");
        var roundTwo= rounds[1].split(",");
        var roundThree= rounds[2].split(",");
         var rIndex = startCol-2;
        var roomName = roomNames[0][rIndex];
        for (p=0;p<roundOne.length;p++){
          roundOne[p]="^"+roundOne[p]+"^" 
        }
        for (p=0;p<roundTwo.length;p++){
          roundTwo[p]="^"+roundTwo[p]+"^" 
        }
        for (p=0;p<roundThree.length;p++){
          roundThree[p]="^"+roundThree[p]+"^" 
        }
        roundOne = roundOne.join();
        roundTwo = roundTwo.join();
        roundThree = roundThree.join();
        for(k=0;k<allCon;k++){
         var value = dataRange[k];
         var thisRow = k*1+3;
          if(value !=""){
         
            if(roundOne.indexOf("^"+value+"^")!=-1){
            
               thisSheet.getRange(thisRow,4).setValue(roomName);} 
          
          
          if(roundTwo.indexOf("^"+value+"^")!=-1){
            
            
              thisSheet.getRange(thisRow,7).setValue(roomName);}
          
          if(roundThree.indexOf("^"+value+"^")!=-1){
           
              thisSheet.getRange(thisRow,10).setValue(roomName);}
       
      
          
        }
      }
    }
   }

  var allTeams = tSheet.getLastRow();
  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
  source.getSheetByName('SWEEPS').copyTo(tabSheet);
  tabSheet.getSheets()[tEvents.length+1].setName('Sweeps');
  var sweepSheet = tabSheet.getSheetByName('Sweeps');
  sweepSheet.getRange(2,2,allTeams,1).setValues(teamNames);
    
  
//  var allTeams = tSheet.getLastRow();
//  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
//  
//  var allEntries = tSheet.getRange(2,5,allTeams,10).getValues();
//  var eventTracker = [3,3,3,3,3,3,3,3,3,3];

  

//  //cycle through each team
//  for(a=0;a<allEntries.length;a++){
//    var thisTeam = allEntries[a];
//    //start the code for this team
//    var theCode = "=sum("
//    //cycle through each event
//    for(b=0;b<thisTeam.length;b++){
//      var tEvent = thisTeam[b];
//      //add some code for each event
//      for(c=0;c<tEvent;c++){
//        var hTrack = eventTracker[b]*1;
//        var cTrack = (c*1)+hTrack;
//      theCode=theCode+tabSheet.getSheets()[b+1].getName()+ "!Y"+hTrack+"+";
//     
//        Logger.log(hTrack);
//        eventTracker[b]=hTrack+1;
//      }
//    }
//  
//    sweepSheet.getRange(2+a,3).setValue(theCode +"0)");
//   
//  }
  
  
   dataSheet.getRange(4,2).setValue('=HYPERLINK("'+tabSheetUrl+'","Tab Sheet")');
    SpreadsheetApp.getUi().alert('Tabs Spreadsheet Created at this URL:' + tabSheetUrl)
}

function makeTabs2(){
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  
  Logger.log("I'm running.");
 var folders= DriveApp.getFoldersByName('4n6 Room Schedules');
  var folder = folders[0];
  var template = "1-ie5JElg67mTEbhmrx-0I4OPAqM_sR7tUo5Qv-WM9YQ";

  var source = ssa.openById(template);
  var tabSheet =source.copy('TM TABS');
  
  var tabSheetId = tabSheet.getId();
  var tabSheetUrl = tabSheet.getUrl();
      var tEvents2 = eSheet.getDataRange().getValues();   var tEvents3 = transpose(tEvents2);   tEvents3[0].shift();   tEvents3[2].shift();      var tEvents = clearBlanks(tEvents3[0]);

  
  var sourceSheet = source.getSheetByName("Event");
tabSheet.insertSheet();
    tabSheet.deleteSheet(tabSheet.getSheetByName('SWEEPS'));
  tabSheet.deleteSheet(tabSheet.getSheetByName('Event'));

 Logger.log(tEvents);
 Logger.log(tEvents.length);
;
  for(q=0,r=2;q<tEvents.length;q++,r++){
   sourceSheet.copyTo(tabSheet);
    var theseEvents = tabSheet.getSheets();
    var allCon = cSheet.getLastRow();
    var contestants = cSheet.getRange(2,q+1,allCon,1).getValues();
    theseEvents[q+1].setName(tEvents[q]);
    theseEvents[q+1].getRange(3,2,allCon,1).setValues(contestants);
    theseEvents[0].hideSheet();
    
    Logger.log(tEvents[q]);
   
 // This is where I put room numbers on the tab sheet     
    var thisSheet=theseEvents[q+1]; 
    var dataRange = thisSheet.getRange(3,2,allCon,1).getValues();
     var lastDataRow = dataSheet.getLastRow();
     var eInd = dataSheet.getRange(2,1,lastDataRow).getValues();   eInd = clearBlanks(eInd);     eInd = twoToOne(eInd);  
   
    var lastRow = rSheet.getLastRow();
    var lastColumn = rSheet.getLastColumn();
    var start = eInd[q];
    var end = eInd[q+1] - 1;
    var numRooms = end - start;
    var roomNames=rSheet.getRange(2,2,1,lastColumn).getDisplayValues();  
      
      for(j=0;j<numRooms;j++){
        var startCol = start*1+j;
      
       var roomSchedule = rSheet.getRange(3,startCol,lastRow,1).getDisplayValues();
       
       var strRounds = roomSchedule.join();
        var rounds = strRounds.split(",,");
        var roundOne= rounds[0].split(",");
        var roundTwo= rounds[1].split(",");
        var roundThree= rounds[2].split(",");
         var rIndex = startCol-2;
        var roomName = roomNames[0][rIndex];
        for (p=0;p<roundOne.length;p++){
          roundOne[p]="^"+roundOne[p]+"^" 
        }
        for (p=0;p<roundTwo.length;p++){
          roundTwo[p]="^"+roundTwo[p]+"^" 
        }
        for (p=0;p<roundThree.length;p++){
          roundThree[p]="^"+roundThree[p]+"^" 
        }
        roundOne = roundOne.join();
        roundTwo = roundTwo.join();
        roundThree = roundThree.join();
        for(k=0;k<allCon;k++){
         var value = dataRange[k];
         var thisRow = k*1+3;
          if(value !=""){
         
            if(roundOne.indexOf("^"+value+"^")!=-1){
            
               thisSheet.getRange(thisRow,4).setValue(roomName);} 
          
          
          if(roundTwo.indexOf("^"+value+"^")!=-1){
            
            
              thisSheet.getRange(thisRow,7).setValue(roomName);}
          
          if(roundThree.indexOf("^"+value+"^")!=-1){
           
              thisSheet.getRange(thisRow,10).setValue(roomName);}
       
      
          
        }
      }
    }
   }

  var allTeams = tSheet.getLastRow();
  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
  source.getSheetByName('SWEEPS').copyTo(tabSheet);
  tabSheet.getSheets()[tEvents.length+1].setName('Sweeps');
  var sweepSheet = tabSheet.getSheetByName('Sweeps');
  sweepSheet.getRange(2,2,allTeams,1).setValues(teamNames);
    
 //This makes the Sweeps formula 
//  var allTeams = tSheet.getLastRow();
//  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
//  
//  var allEntries = tSheet.getRange(2,5,allTeams,10).getValues();
//  var eventTracker = [3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3];
//
//  
//
//  //cycle through each team
//  for(a=0;a<allEntries.length;a++){
//    var thisTeam = allEntries[a];
//    //start the code for this team
//    var theCode = "=sum("
//    //cycle through each event
//    for(b=0;b<thisTeam.length;b++){
//      var tEvent = thisTeam[b];
//      //add some code for each event
//      for(c=0;c<tEvent;c++){
//        var hTrack = eventTracker[b]*1;
//        var cTrack = (c*1)+hTrack;
//      theCode=theCode+tabSheet.getSheets()[b+1].getName()+ "!AA"+hTrack+"+";
//     
//        Logger.log(hTrack);
//        eventTracker[b]=hTrack+1;
//      }
//    }
//  
//    sweepSheet.getRange(2+a,3).setValue(theCode +"0)");
//   
//  }
  
  
  dataSheet.getRange(4,2).setValue('=HYPERLINK("'+tabSheetUrl+'","Tab Sheet")');
    SpreadsheetApp.getUi().alert('Tabs Spreadsheet Created at this URL:' + tabSheetUrl)
}


function makeTabs3(){
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  
  Logger.log("I'm running.");
 var folders= DriveApp.getFoldersByName('4n6 Room Schedules');
  var folder = folders[0];
  var template = "1H7d-IKHtFmbZaoW3fHlpC6mV1APYDoixb8bWsERvpbU";
//var tabSheet = ssa.create('TM TABS');
  var source = ssa.openById(template);
  var tabSheet =source.copy('TM TABS');
  
  var tabSheetId = tabSheet.getId();
  var tabSheetUrl = tabSheet.getUrl();
      var tEvents2 = eSheet.getDataRange().getValues();   var tEvents3 = transpose(tEvents2);   tEvents3[0].shift();   tEvents3[2].shift();      var tEvents = clearBlanks(tEvents3[0]);

  
  var sourceSheet = source.getSheetByName("Event");
tabSheet.insertSheet();
    tabSheet.deleteSheet(tabSheet.getSheetByName('SWEEPS'));
  tabSheet.deleteSheet(tabSheet.getSheetByName('Event'));

 Logger.log(tEvents);
 Logger.log(tEvents.length);
;
  for(q=0,r=2;q<tEvents.length;q++,r++){
   sourceSheet.copyTo(tabSheet);
    var theseEvents = tabSheet.getSheets();
    var allCon = cSheet.getLastRow();
    var contestants = cSheet.getRange(2,q+1,allCon,1).getValues();
    theseEvents[q+1].setName(tEvents[q]);
    theseEvents[q+1].getRange(3,2,allCon,1).setValues(contestants);
    theseEvents[0].hideSheet();
    
   
    // This is where I put room numbers on the tab sheet  
  var thisSheet=theseEvents[q+1]; 
    var dataRange = thisSheet.getRange(3,2,allCon,1).getValues();
     var lastDataRow = dataSheet.getLastRow();
     var eInd = dataSheet.getRange(2,1,lastDataRow).getValues();   eInd = clearBlanks(eInd);     eInd = twoToOne(eInd);   
   
    var lastRow = rSheet.getLastRow();
    var lastColumn = rSheet.getLastColumn();
    var start = eInd[q];
    var end = eInd[q+1] - 1;
    var numRooms = end - start;
    var roomNames=rSheet.getRange(2,2,1,lastColumn).getDisplayValues();  
      
      for(j=0;j<numRooms;j++){
        var startCol = start*1+j;
      
       var roomSchedule = rSheet.getRange(3,startCol,lastRow,1).getDisplayValues();
       
       var strRounds = roomSchedule.join();
        var rounds = strRounds.split(",,");
        var roundOne= rounds[0].split(",");
        var roundTwo= rounds[1].split(",");
        var roundThree= rounds[2].split(",");
         var rIndex = startCol-2;
        var roomName = roomNames[0][rIndex];
        for (p=0;p<roundOne.length;p++){
          roundOne[p]="^"+roundOne[p]+"^" 
        }
        for (p=0;p<roundTwo.length;p++){
          roundTwo[p]="^"+roundTwo[p]+"^" 
        }
        for (p=0;p<roundThree.length;p++){
          roundThree[p]="^"+roundThree[p]+"^" 
        }
        roundOne = roundOne.join();
        roundTwo = roundTwo.join();
        roundThree = roundThree.join();
        for(k=0;k<allCon;k++){
         var value = dataRange[k];
         var thisRow = k*1+3;
          if(value !=""){
         
            if(roundOne.indexOf("^"+value+"^")!=-1){
            
               thisSheet.getRange(thisRow,4).setValue(roomName);} 
          
          
          if(roundTwo.indexOf("^"+value+"^")!=-1){
            
            
              thisSheet.getRange(thisRow,6).setValue(roomName);}
          
          if(roundThree.indexOf("^"+value+"^")!=-1){
           
              thisSheet.getRange(thisRow,8).setValue(roomName);}
       
      
          
        }
      }
    }
   
      
      
      
    Logger.log(tEvents[q]);
   }

  var allTeams = tSheet.getLastRow();
  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
  source.getSheetByName('SWEEPS').copyTo(tabSheet);
  tabSheet.getSheets()[tEvents.length+1].setName('Sweeps');
  var sweepSheet = tabSheet.getSheetByName('Sweeps');
  sweepSheet.getRange(2,2,allTeams,1).setValues(teamNames);
    
  
//  var allTeams = tSheet.getLastRow();
//  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
//  
//  var allEntries = tSheet.getRange(2,5,allTeams,10).getValues();
//  var eventTracker = [3,3,3,3,3,3,3,3,3,3];
//
//  
//
//  //cycle through each team
//  for(a=0;a<allEntries.length;a++){
//    var thisTeam = allEntries[a];
//    //start the code for this team
//    var theCode = "=sum("
//    //cycle through each event
//    for(b=0;b<thisTeam.length;b++){
//      var tEvent = thisTeam[b];
//      //add some code for each event
//      for(c=0;c<tEvent;c++){
//        var hTrack = eventTracker[b]*1;
//        var cTrack = (c*1)+hTrack;
//      theCode=theCode+tabSheet.getSheets()[b+1].getName()+ "!U"+hTrack+"+";
//     
//        Logger.log(hTrack);
//        eventTracker[b]=hTrack+1;
//      }
//    }
//  
//    sweepSheet.getRange(2+a,3).setValue(theCode +"0)");
//   
//  }
  
 
   dataSheet.getRange(4,2).setValue('=HYPERLINK("'+tabSheetUrl+'","Tab Sheet")');
    SpreadsheetApp.getUi().alert('Tabs Spreadsheet Created at this URL:' + tabSheetUrl)
}


function donate(){
  ui.showSidebar(uiDonate)}

function test(){
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  var lastDataRow = dataSheet.getLastRow();
  var eInd = dataSheet.getRange(2,1,lastDataRow).getValues();   eInd = clearBlanks(eInd);     eInd = twoToOne(eInd);   
  eInd = clearBlanks(eInd);
 
  eInd = twoToOne(eInd);
  eInd.pop();

  }
    
  
      
   function makeTabsPanel(){
  
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");

 var folders= DriveApp.getFoldersByName('4n6 Room Schedules');
  var folder = folders[0];
  var template = "1skkuhjPZZBGN_uxINR7DCdYdSlEDwJiN5hU4CgwMqdU";
//var tabSheet = ssa.create('TM TABS');
  var source = ssa.openById(template);
  var tabSheet =source.copy('TM TABS');
  
  var tabSheetId = tabSheet.getId();
  var tabSheetUrl = tabSheet.getUrl();
      var tEvents2 = eSheet.getDataRange().getValues();   var tEvents3 = transpose(tEvents2);   tEvents3[0].shift();   tEvents3[2].shift();      var tEvents = clearBlanks(tEvents3[0]);

  
  var sourceSheet = source.getSheetByName("Event");
tabSheet.insertSheet();
    tabSheet.deleteSheet(tabSheet.getSheetByName('SWEEPS'));
  tabSheet.deleteSheet(tabSheet.getSheetByName('Event'));


  for(q=0,r=2;q<tEvents.length;q++,r++){
   sourceSheet.copyTo(tabSheet);
    var theseEvents = tabSheet.getSheets();
    var allCon = cSheet.getLastRow();
    var contestants = cSheet.getRange(2,q+1,allCon,1).getValues();
    theseEvents[q+1].setName(tEvents[q]);
    theseEvents[q+1].getRange(3,2,allCon,1).setValues(contestants);
    theseEvents[0].hideSheet();
  
    
    // This is where I put room numbers on the tab sheet  
  var thisSheet=theseEvents[q+1]; 
    var dataRange = thisSheet.getRange(3,2,allCon,1).getValues();
     var lastDataRow =1+1* dataSheet.getLastRow();
     var eInd = dataSheet.getRange(2,1,lastDataRow).getValues();   eInd = clearBlanks(eInd);     eInd = twoToOne(eInd);  
   
    var lastRow = rSheet.getLastRow();
    var lastColumn = rSheet.getLastColumn();
    var start = eInd[q];
    var end = eInd[q+1] - 1;
    var numRooms = end - start;
    var roomNames=rSheet.getRange(2,2,1,lastColumn).getDisplayValues();  
      
      for(j=0;j<numRooms;j++){
        var startCol = start*1+j;
      
       var roomSchedule = rSheet.getRange(3,startCol,lastRow,1).getDisplayValues();
       
        var strRounds = roomSchedule.join();
        var rounds = strRounds.split(",,");
        var roundOne= rounds[0].split(",");
        var roundTwo= rounds[1].split(",");
        var roundThree= rounds[2].split(",");
         var rIndex = startCol-2;
        var roomName = roomNames[0][rIndex];
        for (p=0;p<roundOne.length;p++){
          roundOne[p]="^"+roundOne[p]+"^" 
        }
        for (p=0;p<roundTwo.length;p++){
          roundTwo[p]="^"+roundTwo[p]+"^" 
        }
        for (p=0;p<roundThree.length;p++){
          roundThree[p]="^"+roundThree[p]+"^" 
        }
        roundOne = roundOne.join();
        roundTwo = roundTwo.join();
        roundThree = roundThree.join();
        for(k=0;k<allCon;k++){
         var value = dataRange[k];
         var thisRow = k*1+3;
          if(value !=""){
         
            if(roundOne.indexOf("^"+value+"^")!=-1){
            
               thisSheet.getRange(thisRow,4).setValue(roomName);} 
          
          
          if(roundTwo.indexOf("^"+value+"^")!=-1){
            
            
              thisSheet.getRange(thisRow,14).setValue(roomName);}
          
          if(roundThree.indexOf("^"+value+"^")!=-1){
           
              thisSheet.getRange(thisRow,24).setValue(roomName);}
       
      
          
        }
      }
    }
   }

  var allTeams = tSheet.getLastRow();
  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
  source.getSheetByName('SWEEPS').copyTo(tabSheet);
  tabSheet.getSheets()[tEvents.length+1].setName('Sweeps');
  var sweepSheet = tabSheet.getSheetByName('Sweeps');
  sweepSheet.getRange(2,2,allTeams,1).setValues(teamNames);
    
  
  var allTeams = tSheet.getLastRow();
  var teamNames = tSheet.getRange(2,2,allTeams,1).getValues();
  
  var allEntries = tSheet.getRange(2,5,allTeams,10).getValues();
  var eventTracker = [3,3,3,3,3,3,3,3,3,3];

  

  //cycle through each team
  for(a=0;a<allEntries.length;a++){
    var thisTeam = allEntries[a];
    //start the code for this team
    var theCode = "=sum("
    //cycle through each event
    for(b=0;b<thisTeam.length;b++){
      var tEvent = thisTeam[b];
      //add some code for each event
      for(c=0;c<tEvent;c++){
        var hTrack = eventTracker[b]*1;
        var cTrack = (c*1)+hTrack;
      theCode=theCode+tabSheet.getSheets()[b+1].getName()+ "!AT"+hTrack+"+";
     
        Logger.log(hTrack);
        eventTracker[b]=hTrack+1;
      }
    }
  
    sweepSheet.getRange(2+a,3).setValue(theCode +"0)");
   
  }
  
  
   dataSheet.getRange(4,2).setValue('=HYPERLINK("'+tabSheetUrl+'","Tab Sheet")');
    SpreadsheetApp.getUi().alert('Tabs Spreadsheet Created at this URL:' + tabSheetUrl)
}  
   
   
      
    function codeCon(){
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  var data = tSheet.getDataRange().getValues();
  var dataOne = transpose(data);
  var teams = dataOne[0];
  var events = data[0];
      events.shift();
      events.shift();
      events.shift();
      events.shift();
    events =  oneToTwo(events);
    
  teams = clearBlanks(teams);
  teams.shift();
 var numTeams = teams.length;
 
  
  data.shift();
 var letters = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,"
  var letter=letters.split(",");

      for(i=0;i<numTeams;i++){
      var thisTeam = teams[i];
      var tEvents = data[i];
        tEvents.shift();
        tEvents.shift();
        tEvents.shift();
        tEvents.shift();
        
        for (j=0;j<events.length;j++){
          for(k=0;k<tEvents[j];k++){
           var newCont = thisTeam+letter[k]
         
            events[j].push(newCont);
            
          }
         
        
        }
      
      }
      for(i=0;i<events.length;i++){
        while(events[i].length<100){
          events[i].push("");}          
          }
      events = transpose(events);
      Logger.log(events);
      range = cSheet.getRange(1, 1,events.length,events[0].length).setValues(events);
}
   
