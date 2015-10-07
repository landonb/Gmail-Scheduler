













/* This Google Script was written by Amit Agarwal */

/* Web: http://www.labnol.org/?p=24867            */

/* Apps Script Development: http://ctrlq.org      */



/* Please retain this message in your copy        */

















/*
* @OnlyCurrentDoc
*/


function help_() {
  var html = HtmlService.createHtmlOutputFromFile('help')
  .setTitle("Google Scripts Support")
  .setWidth(550)
  .setHeight(350);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}




function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [ 
    {name: "☎ Help and Support »",functionName: "help_"},    
    null, 
    {name: "Step 1: Authorize", functionName: "authorize_"},
    {name: "Step 2: Fetch Messages", functionName: "initialize_"},
    {name: "Step 3: Schedule Messages", functionName: "setSchedule_"},
    null,
    {name: "✘ Cancel Pending Jobs", functionName: "cancelJobs_"},
    null,
    {name: "Schedule Recurring Email", functionName: "premium_"},
    {name: "Increase Message Limit", functionName: "premium_"},
    {name: "Send Emails from another Address", functionName: "premium_"}
  ];  
  ss.addMenu("➪ Gmail Scheduler", menu);
}

function premium_() {
  Browser.msgBox("This is available in the Premium edition of Gmail Scheduler. Visit http://www.labnol.org/?p=24867 for details");
}

function authorize_() {  
  deleteTriggers_();
  Browser.msgBox("Your messages will be scheduled in the " + SpreadsheetApp.getActive().getSpreadsheetTimeZone() + " timezone. See Help under the Gmail Scheduler menu above for instructions on how to change your time zone.");
}

function fetchDraftMessages_() {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var drafts = GmailApp.getDraftMessages();
  
  var msgIDs = sheet.getRange("A2:A").getValues().toString();
  
  if (drafts.length > 0) {
    
    for (var i=0; i<drafts.length; i++) {
      if ((drafts[i].getTo() !== "") && (msgIDs.indexOf(drafts[i].getId()) === -1)) {
        sheet.appendRow([drafts[i].getId(), drafts[i].getTo(), drafts[i].getSubject(), "", ""]);
      }
    }
    
    SpreadsheetApp.getActive().toast("Excellent. Now please enter the date and time when you would like these messages to be delivered."); 
    
  } else {
    SpreadsheetApp.getActive().toast("We could not find any messages in your Gmail drafts folder.");
  } 
}

function cancelJobs_() {
  deleteTriggers_();
  var sheet = SpreadsheetApp.getActiveSheet();
  var data  = sheet.getDataRange().getValues(); 
  for (var row=2; row<data.length; row++) {
    sheet.getRange("E"+(row+1)).setValue("Not Scheduled");
  } 
  SpreadsheetApp.getActive().toast("There are no pending email messages in the queue.");
}

function initialize_() {
  deleteTriggers_();
  fetchDraftMessages_();
}

function clearSheet_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(2, 1, sheet.getLastRow()+1, 5).clearContent();
}

function setSchedule_() {
  
  try {
    
    deleteTriggers_();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var data  = sheet.getDataRange().getValues();
    var time  = new Date().getTime();
    var oneminute = 1000*60*1;
    var code  = [];
    var diff, triggers = [];
    
    for (var row=2; row<data.length; row++ ) {
      
      if (data[row][0] === "") continue;
      
      var schedule = data[row][3];
      
      if ( schedule !== "" ) {
        
        diff = schedule.getTime() - time;
        
        if ( diff > 1000*60 ) {
          
          var roundTime = Math.ceil(schedule.getTime()/oneminute)*oneminute;
          
          if (triggers.indexOf(roundTime) === -1) {
            
            triggers.push(roundTime);
            
            ScriptApp.newTrigger("sendMails")
            .timeBased()
            .at(new Date(roundTime))
            .create();
            
          }
          code.push("Scheduled " + humanReadable_(diff) + " from now");
          
        } else {
          code.push("Date is in the past. Have you set the correct timezone?");  
        }
      } else {
        code.push("Not Scheduled");
      }
    }
    
    for (var i=0; i<code.length; i++) {
      sheet.getRange("E" + (i+3)).setValue(code[i]);
    }
    
    if (code.length) {
      ss.toast("The mails have been scheduled. You can close this sheet. For help, contact amit@labnol.org", "Success", -1);
    } else {
      ss.toast("Sorry but no emails were scheduled.", "Oops", -1);
    }
    
  } catch (e) {
    ss.toast(e.toString());
  }
}

function deleteTriggers_() { 
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendMails") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


function sendMails() {
  
  try {
    
    var sheet = SpreadsheetApp.getActiveSheet();
    var data  = sheet.getRange(3,1,sheet.getLastRow(), 5).getValues(); 
    var time  = new Date().getTime();
    for (var row=0; row<data.length; row++) {
      var schedule = data[row][3];
      if ((schedule !== "") && data[row][4].match(/^Scheduled/)) {             
        if ( schedule.getTime() <= time ) {
          var status = dispatchDraft_(data[row][0]);
          sheet.getRange("E"+(row+3)).setValue(status);
          SpreadsheetApp.flush();
        } 
      }
    }
  } catch (e) {
    SpreadsheetApp.getActive().toast(e.toString());
  }
}

function getSSTZ(e) {
  return "Your timezone is " + SpreadsheetApp.getActive().getSpreadsheetTimeZone();  
}

function dispatchDraft_(id) {
  
  try {
    
    var message = GmailApp.getMessageById(id);
    
    if (message) {
      
      var body = message.getBody();
      var raw  = message.getRawContent();
      
      var regMessageId = new RegExp(id, "g");
      if (body.match(regMessageId) != null) {
        var inlineImages = {};
        var nbrOfImg = body.match(regMessageId).length;
        var imgVars = body.match(/<img[^>]+>/g);
        var imgToReplace = [];
        if(imgVars != null){
          for (var i = 0; i < imgVars.length; i++) {
            if (imgVars[i].search(regMessageId) != -1) {
              var id = imgVars[i].match(/realattid=([^&]+)&/);
              if (id != null) {
                id = id[1];
                var temp = raw.split(id)[1];
                temp = temp.substr(temp.lastIndexOf('Content-Type'));
                var imgTitle = temp.match(/name="([^"]+)"/);
                var contentType = temp.match(/Content-Type: ([^;]+);/);
                contentType = (contentType != null) ? contentType[1] : "image/jpeg";
                var b64c1 = raw.lastIndexOf(id) + id.length + 3; // first character in image base64
                var b64cn = raw.substr(b64c1).indexOf("--") - 3; // last character in image base64
                var imgb64 = raw.substring(b64c1, b64c1 + b64cn + 1); // is this fragile or safe enough?
                var imgblob = Utilities.newBlob(Utilities.base64Decode(imgb64), contentType, id); // decode and blob
                if (imgTitle != null) imgToReplace.push([imgTitle[1], imgVars[i], id, imgblob]);
              }
            }
          }
        }
        
        for (var i = 0; i < imgToReplace.length; i++) {
          inlineImages[imgToReplace[i][2]] = imgToReplace[i][3];
          var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + imgToReplace[i][2] + "\"");
          body = body.replace(imgToReplace[i][1], newImg);
        }
      }
      
      var options = {
        cc          : message.getCc(),
        bcc         : message.getBcc(),
        htmlBody    : body,
        replyTo     : message.getReplyTo(),
        inlineImages: inlineImages,
        attachments : message.getAttachments()
      }
      
      GmailApp.sendEmail(message.getTo(), message.getSubject(), body, options);
      message.moveToTrash();
      return "Delivered";
    } else {
      return "Message not found in Drafts";
    }
  } catch (e) {
    return e.toString();
  }
}



function humanReadable_(milliseconds) {
  
  function numberEnding (number) {
    return (number > 1) ? 's' : '';
  }
  
  var temp = Math.floor(milliseconds / 1000);
  
  var years = Math.floor(temp / 31536000);
  if (years) {
    return years + ' year' + numberEnding(years);
  }
  var days = Math.floor((temp %= 31536000) / 86400);
  if (days) {
    return days + ' day' + numberEnding(days);
  }
  var hours = Math.floor((temp %= 86400) / 3600);
  if (hours) {
    return hours + ' hour' + numberEnding(hours);
  }
  var minutes = Math.floor((temp %= 3600) / 60);
  if (minutes) {
    return minutes + ' minute' + numberEnding(minutes);
  }
  var seconds = temp % 60;
  if (seconds) {
    return seconds + ' second' + numberEnding(seconds);
  }
  return 'just now';
}


