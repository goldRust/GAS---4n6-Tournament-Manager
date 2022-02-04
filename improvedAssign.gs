


function improvedAssign(){
  
  var eSheet= ss.getSheetByName("Events");
var tSheet=ss.getSheetByName("Teams");
var cSheet=ss.getSheetByName("Contestants");
var rSheet=ss.getSheetByName("Rooms");
var dataSheet=ss.getSheetByName("TM Data");
var testSheet=ss.getSheetByName("Test");
  
  
    //clears the roomsheet incase of old data
    rSheet.clear();
   // Check for dataSheet - delete and replace if it exists
  var dataSheet = ss.getSheetByName('TM Data');
  if (dataSheet == null){
   ss.insertSheet('TM Data') ;}

  var dataSheet = ss.getSheetByName('TM Data');
  dataSheet.getRange(1,2).setValue("This sheet contains data vital to running Tournament Manager. Do not adjust or delete the information here.");
  dataSheet.getRange(1,1,dataSheet.getLastRow(),1).clear();
  
 var tEvents2 = eSheet.getDataRange().getValues();
  var tEvents3 = transpose(tEvents2);
  tEvents3[0].shift();
  tEvents3[2].shift();
  
  var tEvents = clearBlanks(tEvents3[0]);
   var numEvents = tEvents.length;
  var mostEntries = cSheet.getLastRow();
  var entryNumbers = eSheet.getRange(2,2,numEvents).getValues();
  var conData = cSheet.getDataRange().getValues();
  var tConData = transpose(conData);
  var tsConData = [];

 
  for (d=0;d<tConData.length;d++){

  tsConData.push(clearBlanks(tConData[d]));
    tsConData[d].shift();

  
  }

var everyCon = tsConData;
  
  
 Array.prototype.swapItems = function(a, b){
    this[a] = this.splice(b, 1, this[a])[0];
    return this;
}
  
  var col = 2;
  var cols = 2;

  var nRooms = clearBlanks(tEvents3[2]);

  var dezRooms = [];
 
  //Cycle through each event.
  var numEvents = tEvents.length;
  for(a=0;a<numEvents;a++){
   var eventName = tEvents[a];
   var roomText=1;
  var eRooms = [];  
  var counter = 0;
  var thisEvent = everyCon[a];    
  var nRoom = nRooms[a];
  var row = 3;
    col = cols;
    //Inform dataSheet
    dataSheet.getRange(a+2,1).setValue(col);

   //Label events and rooms 
  rSheet.getRange(1,col).setValue(eventName);
    for(e=0;e<nRoom;e++){
      var rCol = col*1+e*1;
      var r=e*1+1;
      rSheet.getRange(2,rCol).setValue("Room "+r);}
    
  while(thisEvent.length % nRoom !=0){
    thisEvent.push("Open");}
      
  
    
  var figure = thisEvent.length / nRoom;
   
    //Cycle through Rooms
    for(b=0;b<figure;b++){
      var endPoint = counter*1+nRoom*1;
      var thisRoom = thisEvent.slice(counter,endPoint);
      
     
   counter = endPoint;
      
      eRooms.push(thisRoom);  

      }//End of Round 1 Cycle
    var rows = eRooms.length;
       
                rSheet.getRange(row,col,rows,nRoom).setValues(eRooms);
    row = row*1+rows*1+1;
    //Move Contestants
    
     if(nRoom %2==0){
       //2 Room Problem round 2
       if(nRoom ==2){
         var newRooms=eRooms;
         var swaps = Math.round(eRooms.length/2);
         Logger.log(eRooms);
         for(y=0;y<swaps;y+=2){
           var conChange = newRooms[y][0];
           newRooms[y].shift();
           newRooms[y].push(conChange);
         }
         eRooms = newRooms;
       Logger.log(eRooms);
       }else{
         eRooms.forEach(shuffle);}
    }
    
    if(nRoom %2!=0){
    for(c=0,d=0;c<thisEvent.length-1;c+=2,d++){
      var next = c*1+1;
      
      var temp1= thisEvent[c];
      var temp2 = thisEvent[next];
      if(temp2!= undefined){
      thisEvent[c] = temp2;
        thisEvent[next]= temp1;}
    }
    eRooms=[];
    //do round 2
    
    counter=0;
    for(b=0;b<figure;b++){
      var endPoint = counter*1+nRoom*1;
      var thisRoom = thisEvent.slice(counter,endPoint);
      
     
   counter = endPoint;
      
      eRooms.push(thisRoom);  

    }}//End of Round 2 Cycle
    var rows = eRooms.length;
   
                rSheet.getRange(row,col,rows,nRoom).setValues(eRooms);
    row = row*1+rows*1+1;
    
    //move contestants again
    if(nRoom %2==0){
     if(nRoom ==2){
         var newRooms=eRooms;
       var len = eRooms.length;
         var third = Math.round(eRooms.length/3);
       Logger.log(third);
         Logger.log(eRooms);
         for(y=third-1;y<third+1;y++){
           Logger.log('ERROR HERE:')
           
           Logger.log('newRooms: ' + newRooms)
           Logger.log('y: ' + y)  
           Logger.log('newRooms[y]: ' + newRooms[y])
           Logger.log('newRooms[y][0]' + newRooms[y][0])
           
           
           var conChange = newRooms[y][0];
           newRooms[y].shift();
           newRooms[y].push(conChange);
         }
         eRooms = newRooms;
       Logger.log(eRooms);
       }else{
         eRooms.forEach(shuffle);}
    }
    
     if(nRoom %2 != 0){
    thisEvent = everyCon[a];
    for(c=1,d=0;c<thisEvent.length-1;c+=2,d++){
      var next = c*1+1;
      
      var temp1= thisEvent[c];
      var temp2 = thisEvent[next];
     
     
      if(temp2!= undefined){
      thisEvent[c] = temp2;
        thisEvent[next]= temp1;}
      
      eRooms=[]}
      
     
    
    

    //do round 3
    
    counter=0;
    for(b=0;b<figure;b++){
      var endPoint = counter*1+nRoom*1;
      var thisRoom = thisEvent.slice(counter,endPoint);
      
     
   counter = endPoint;
      
      eRooms.push(thisRoom);  

    }}//End of Round 3 Cycle
 
          var rows = eRooms.length;
       
                rSheet.getRange(row,col,rows,nRoom).setValues(eRooms);
    row = row*1+rows*1+1;
    
 cols = cols*1+nRoom*1+1*1; }//End of Event Cycle
var finalRow = dataSheet.getLastRow()*1+1;
dataSheet.getRange(finalRow,1).setValue(rSheet.getLastColumn()+2)  
dataSheet.hideColumns(1);
}//End of Function



    

   
  