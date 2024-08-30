/* User Supplied Variables Start*/

var url = "https://api.maytapi.com/api/"+ '4dcb7d66-457a-40cf-84dd-1cc8a47940d0'
var token = '9b1af49a-3ee8-4978-bd32-3a681f086cc7'
var apiInstance = '53776'

var api =  url + "/" + String(apiInstance) + "/sendMessage?token=" + token

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email and Whatsapp')
        .addItem('Send Now', 'email')
      .addItem('Set Trigger (Daily at 8)', 'createTrigger')
      .addItem('Remove Trigger','removeTrigger')
      .addToUi();
}

function createTrigger() {
  removeTrigger()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('assigntasktodoer')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
}

function removeTrigger() {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[i].getUniqueId() == allTriggers[i].getUniqueId()) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}

function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  try{
  for(var row = 0; row < range.length; row++){
 
    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;
    }else if(range[row][0] !== ""){
      blank = false;
    };
  };
  return rowNum;
  }
  catch(e){}
  };

function findinB(name,data) {
  var valB= name
  try{
  for(nn=0;nn<data.length;++nn){
    if (data[nn][1]==valB){break} ;// if a match in column B is found, break the loop
      }
  return data[nn][0];// show column A
}
  catch(e){}
  }




function assigntasktodoer() {  
  var ss = SpreadsheetApp.getActive()
  var sheet = ss.getSheetByName("Master")
  var lastrow = sheet.getRange('B:B').getValues().filter(String).length-1

  Logger.log(lastrow)
  var data = sheet.getRange(2, 1, lastrow, 14).getValues()
  var nlen2 = data.filter(function(value){return value[0]}).length
  for (var i = 0; i < nlen2; ++i){
    var phone = sheet.getRange(2+i,20).getValue().toString();
    var taskid = sheet.getRange(2+i,1).getValue();
    var name = sheet.getRange(2+i,3).getValue();
    var task = sheet.getRange(2+i,4).getValue();
    var lastDate = sheet.getRange(2+i,11).getValue();
    var status = sheet.getRange(2+i,13).getValue();
    var revisions = sheet.getRange(2+i,12).getValue();
    var sentstatus0 = sheet.getRange(2+i,22).getValue();
    var sentstatus1 = sheet.getRange(2+i,23).getValue();
    var sentstatus2 = sheet.getRange(2+i,24).getValue();
    var email = sheet.getRange(2+i,21).getValue();
    var taskExplanation = sheet.getRange(2+i,5).getValue();
    var refDocument = sheet.getRange(2+i,7).getValue();
    Logger.log(email)
    
  var finaltargetDatetargetDate = new Date(lastDate)
  finaltargetDatetargetDate.setDate(finaltargetDatetargetDate.getDate() + 1);
  var date = Utilities.formatDate(finaltargetDatetargetDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  var tasklink = "https://docs.google.com/forms/d/e/1FAIpQLScuWsKZUDZik8JqFp6fJSv5LrjjBTyEkTGOHls1bBY8XH4BLQ/viewform?usp=pp_url&entry.76073654=" + taskid;


var subject0 = "New Delegation Task";
var message0 = "You have received a new delegation task. Please click on the link below to complete the given task as per Deadline :\n\n" +
              "Task ID: " + taskid + "\n" +
              "Name: " + name + "\n" +
              "Task: " + task + "\n" +
              "Task Explanation: " + taskExplanation + "\n" +
              "Reference Document/ Image: " + refDocument + "\n" +
              "Target Date: " + date + "\n" +
              "Completion Link: " + tasklink + "\n\n" +
              "Note: Please do not remove the Unique ID from the Google Form.";

var whatMessage0 = "Hello *"+name+"*,\n\nYou have received a new delegation task Today.\n\n*Task ID: "+taskid+"*\n\n*Task : "+task+"*\n\nPlease Completed this task as per Deadline:- *"+date+"*\n\nCompletion Link \n"+tasklink+ "\n\nThanks!\nTeam TPC";

var subject1 = "Delegation Task (Revision 1)";
var message1 = "You have received a new delegation task. Please click on the link below to complete the given task as per Deadline :\n\n" +
              "Task ID: " + taskid + "\n" +
              "Name: " + name + "\n" +
              "Task: " + task + "\n" +
              "Task Explanation: " + taskExplanation + "\n" +
              "Reference Document/ Image: " + refDocument + "\n" +
              "Target Date: " + date + "\n" +
              "Completion Link: " + tasklink + "\n\n" +
              "Note: Please do not remove the Unique ID from the Google Form.";


var whatMessage1 = "Hello *"+name+"*,\n\nYou You have received a delegation task *(Revision 1)*.\n\n*Task ID: "+taskid+"*\n\n*Task : "+task+"*\n\nPlease Completed this task as per Deadline:- *"+date+"*\n\n Completion Link \n"+tasklink+ "\n\nThanks!\nTeam TPC";

var subject2 = "Delegation Task (Revision 2)";
var message2 = "You have received a new delegation task. Please click on the link below to complete the given task as per Deadline :\n\n" +
              "Task ID: " + taskid + "\n" +
              "Name: " + name + "\n" +
              "Task: " + task + "\n" +
              "Task Explanation: " + taskExplanation + "\n" +
              "Reference Document/ Image: " + refDocument + "\n" +
              "Target Date: " + date + "\n" +
              "Completion Link: " + tasklink + "\n\n" +
              "Note: Please do not remove the Unique ID from the Google Form.";

var whatMessage2 = "Hello *"+name+"*,\n\nYou have received a delegation task *(Revision 2)*.\n\n*Task ID: "+taskid+"*\n\n*Task : "+task+"*\n\nPlease Completed this task as per Deadline :- *"+date+"*\n\nCompletion Link \n"+tasklink+ "\n\nThanks!\nTeam TPC";


if(revisions==0){

  if (!sentstatus0 && lastDate!="" && phone!="" && email != ""){
   
    Logger.log(true);
    Logger.log(phone);   
    sheet.getRange(2 + i,22).setValue("Sent");
    sheet.getRange(2 + i,25).setValue(new Date());
    sendMessageW(phone,whatMessage0)
    MailApp.sendEmail(email, subject0, message0);
  
  }
  else if(!sentstatus0 && lastDate!="" && email != ""){
    sheet.getRange(2 + i,22).setValue("Sent");
    sheet.getRange(2 + i,25).setValue(new Date());    
    MailApp.sendEmail(email, subject0, message0);
  
  }

  

}else if(revisions == 1 ){

  if (!sentstatus1 && lastDate!=""&&phone!="" && email != ""){  

    sheet.getRange(2 + i,23).setValue("Sent");
    sendMessageW(phone,whatMessage1)
    MailApp.sendEmail(email, subject1, message1);
  } 
  else if(!sentstatus1 && lastDate!="" && email != ""){
    sheet.getRange(2 + i,23).setValue("Sent");        
    MailApp.sendEmail(email, subject1, message1);
  
  }

}else if(revisions == 2 ){

  if (!sentstatus2 && lastDate!=""&&phone!="" && email != ""){
     
    sheet.getRange(2 + i,24).setValue("Sent");
    sendMessageW(phone,whatMessage2)
    MailApp.sendEmail(email, subject2, message2);

  }else if(!sentstatus2 && lastDate!="" && email != ""){
    sheet.getRange(2 + i,24).setValue("Sent");    
    MailApp.sendEmail(email, subject2, message2);
  
  }
}
}
}


function sendreminder() {  
  var ss = SpreadsheetApp.getActive()
  var sheet = ss.getSheetByName("Master")
  var lastrow = sheet.getRange('B:B').getValues().filter(String).length-1
  var data = sheet.getRange(2, 1, lastrow, 28).getValues()
  var nlen2 = data.filter(function(value){return value[0]}).length
  for (var i = 0; i < nlen2; ++i){
    var phone = sheet.getRange(2+i,20).getValue().toString();
    var taskid = sheet.getRange(2+i,1).getValue();
    var name = sheet.getRange(2+i,3).getValue();
    var task = sheet.getRange(2+i,4).getValue();
    var lastDate = sheet.getRange(2+i,11).getValue();
    var status = sheet.getRange(2+i,13).getValue();
    var currentDate = new Date();

    
    
    var today = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");
   
    var dayNumber = currentDate.getDay();    
    Logger.log(dayNumber)
    if(dayNumber === 0){
      return;
    }

    var finaltargetDatetargetDate = new Date(lastDate)
    finaltargetDatetargetDate.setDate(finaltargetDatetargetDate.getDate() + 1);
    var targetDate = Utilities.formatDate(finaltargetDatetargetDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");

    var tasklink = "https://docs.google.com/forms/d/e/1FAIpQLScuWsKZUDZik8JqFp6fJSv5LrjjBTyEkTGOHls1bBY8XH4BLQ/viewform?usp=pp_url&entry.76073654=" + taskid;

    var whatMessage = "Hello *"+name+"*,\n\nYou have a Delegated task pending.\n\n*Task ID: "+taskid+"*\n\n*Task : "+task+"*\n\nPlease Completed this task as per Deadline:- *"+targetDate+"* \n\nPlease ignore this message if you have already completed this task" + "\n\nCompletion Link: \n"+tasklink+ "\n\nThanks!\nTeam TPC";

 
    
    Logger.log(true)
    if(phone !=""){
      if (targetDate && targetDate <= today && status == "Pending"){
        //Logger.log(targetDate)
        //Logger.log(date)
       
        sendMessageW(phone,whatMessage)
      }
    }
       
            
      
    
  }
}


function sendpendingreminder() {
  //var today = new Date((new Date().setHours(0,0,0,0)).valueOf() + 1000*3600*24);
  var ss = SpreadsheetApp.getActive()
  var sheet = ss.getSheetByName("Master")
  var lastrow = sheet.getRange('B:B').getValues().filter(String).length-1
  var data = sheet.getRange(2, 1, lastrow, 12).getValues()
  var nlen2 = data.filter(function(value){return value[0]}).length
  for (var i = 0; i < nlen2; ++i){
    var phone = sheet.getRange(2+i,20).getValue().toString();
    var taskid = sheet.getRange(2+i,1).getValue();
    var name = sheet.getRange(2+i,3).getValue();
    var task = sheet.getRange(2+i,4).getValue();
    var lastDate = sheet.getRange(2+i,11).getValue();
    var status = sheet.getRange(2+i,13).getValue();    
    
  var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEE");
  var date = Utilities.formatDate(new Date(lastDate), Session.getScriptTimeZone(), "dd-MMM-yyyy");
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy");
  var yesterdayDate = new Date();
  yesterdayDate.setDate(yesterdayDate.getDate() - 1);
  var yesterday = Utilities.formatDate(yesterdayDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  
  
  Logger.log(date)
  Logger.log(today)
  Logger.log(yesterday)
  if(date<today){
    Logger.log(true)
  }

  return;

  var tasklink = "https://docs.google.com/forms/d/e/1FAIpQLScuWsKZUDZik8JqFp6fJSv5LrjjBTyEkTGOHls1bBY8XH4BLQ/viewform?usp=pp_url&entry.76073654=" + taskid;
  
  var whatMessage = "Hello *"+name+"*,\n\nYou have a Delegated task pending from yesterday.\n\n*Task ID: "+taskid+"*\n\n*Task : "+task+"*\n\nPlease Completed this task as per Deadline:- *"+date+"* \n\nPlease ignore this message if you have already completed this task"+"\n\nCompletion Link: \n"+tasklink+ "\n\nThanks!\nTeam TPC";

   if(phone !=""){   
   if (lastDate < today && status == "Pending"){     
   sendMessageW(phone,whatMessage)

  }
  }
  }
}

function sendMessageW(contact,message){
  var whatSend = {
  "to_number": "91" + contact.toString(),
  "type": "text",
  'message': message,
   
  };
 
  var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'payload' : JSON.stringify(whatSend)
         };
  var sendNow = UrlFetchApp.fetch(api, options); 

}



function archive() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Dashboard")
  var data = sheet.getRange(4, 1, 6, 4).getValues()
  var name = sheet.getRange("A2").getValue()
  var week = sheet.getRange("D2").getValue()
  var arch = ss.getSheetByName("Archive")
  var archLast = arch.getLastRow() - 1
  var ared = data[3][1]
  var ayellow = data[3][2]
  var agreen = data[3][3]
  var pred = data[5][1]
  var pyellow = data[5][2]
  var pgreen = data[5][3]
  arch.appendRow([name,week,pred,pyellow,pgreen,ared,ayellow,agreen])
  var spreadsheet = ss.getSheetByName("Dashboard")
  spreadsheet.getRange('B9:D9').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
}

function fix() {
  var ss = SpreadsheetApp.getActive()
  var sh = ss.getSheetByName("Master")
  sh.getRange("F2").setFormula('=arrayformula(if(E2:E<>"",E2:E,if(D2:D<>"",D2:D,if(C2:C<>"",C2:C,""))))')
  sh.getRange("G2").setFormula('=ARRAYFORMULA(if(E2:E<>"",2,if(D2:D<>"",1,if(C2:C<>"",0,""))))')
  sh.getRange("I2").setFormula('=arrayformula(if(F2:F=$K$1,if(H2:H<="Pending","Today",""),""))')
  sh.getRange("K1").setFormula('=Today()')
  sh.getRange("F3:G").clearContent();
  sh.getRange("I3:I").clearContent();
  sh.getRange('F:F').setNumberFormat('dd/MM/yyyy');
  var dash = ss.getSheetByName("Dashboard")
    dash.getRange("B2").setFormula('=if(D2,filter(\'Week List\'!B2:B53,\'Week List\'!A2:A53=D2),"")')
    dash.getRange("C2").setFormula('=if(D2,filter(\'Week List\'!B2:B54,\'Week List\'!A2:A54=D2+1),"")')
    dash.getRange("B5").setFormula('=ifna(QUERY(Archive!A:E,"select C where A=\'"&A2&"\' and B="&D2-1&" limit 1 label C \'\'"),"")')
    dash.getRange("C5").setFormula('=ifna(QUERY(Archive!A:E,"select D where A=\'"&A2&"\' and B="&D2-1&" limit 1 label D \'\'"),"")')
    dash.getRange("D5").setFormula('=ifna(QUERY(Archive!A:E,"select E where A=\'"&A2&"\' and B="&D2-1&" limit 1 label E \'\'"),"")')
    dash.getRange("B7").setFormula('=iferror(count(filter(Master!G:G,Master!F:F>=B2,Master!F:F<=C2,Master!A:A=A2,Master!G:G=2))/count(filter(Master!G:G,Master!F:F>=B2,Master!F:F<=C2,Master!A:A=A2))*100,"Not Found")')
    dash.getRange("C7").setFormula('=iferror(count(filter(Master!G:G,Master!F:F>=B2,Master!F:F<=C2,Master!A:A=A2,Master!G:G=1))/count(filter(Master!G:G,Master!F:F>=B2,Master!F:F<=C2,Master!A:A=A2))*100,"Not Found")')
    dash.getRange("D7").setFormula('=iferror(count(filter(Master!G:G,Master!F:F>=B2,Master!F:F<=C2,Master!A:A=A2,Master!G:G=0))/count(filter(Master!G:G,Master!F:F>=B2,Master!F:F<=C2,Master!A:A=A2))*100,"Not Found")')

}

function getCurrentDayInDDDFormat() {
  // Get the current date
  var currentDate = new Date();

  // Format the date to get the day in 'ddd' format
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "EEE");

  // Log the result to the console (you can remove this line in production)
  Logger.log(formattedDate);

  // Return the formatted date
  return formattedDate;
}



