// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Extract Emails')
      .addItem('Extract Emails...', 'extractEmails')
      .addToUi();
}

// extract emails from label in Gmail
function extractEmails() {
  
  // get the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var label = sheet.getRange(1,2).getValue();
  var emailAddress = "volunteers@rethinkfoundation.in";
  
  // get all email threads that match label from Sheet
  //var threads = GmailApp.search ("to:volunteers@rethinkfoundation.in");
  //var threads = GmailApp.getInboxThreads()
  var threads = GmailApp.search ("label:" + label);
  
  // get all the messages for the current batch of threads
  var messages = GmailApp.getMessagesForThreads (threads);
  Logger.log(threads.length);
  
  var emailArray = [];
  
  // get array of email addresses
  messages.forEach(function(message) {
    message.forEach(function(d) {
      Logger.log(d.getFrom()+" -> "+d.getTo());
      var to = d.getTo()+"";
      var cc = d.getCc()+"";
      if(to.match(emailAddress) || cc.match(emailAddress)){
      //if(d.getTo() == emailAddress){
        //Logger.log(d.getFrom()+", "+d.getTo()+", "+d.getSubject());
        emailArray.push(d.getFrom(),d.getTo());
      }
    });
  });
   
  // de-duplicate the array
  var uniqueEmailArray = emailArray.filter(function(item, pos) {
    return emailArray.indexOf(item) == pos;
  });
  
  var cleanedEmailArray = uniqueEmailArray.map(function(el) {
    var name = "";
    var email = "";
    
    var matches = el.match(/\s*"?([^"]*)"?\s+<(.+)>/);
   
    if (matches) {
      name = matches[1]; 
      email = matches[2];
    }
    else {
      name = "N/k";
      email = el;
    }
    
    return [name,email];
  }).filter(function(d) {
    if (
         d[1] !== "benlcollins@gmail.com" &&
         d[1] !== "drive-shares-noreply@google.com" &&
         d[1] !== "wordpress@www.benlcollins.com"
       ) {
      return d;
    }
  });
  
  // clear any old data
  //sheet.getRange(4,1,sheet.getLastRow(),2).clearContent();
  
  // paste in new names and emails and sort by email address A - Z
  if(cleanedEmailArray.length == 0){
    var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Alert!',
     'There is no email.',
      ui.ButtonSet.OK);
  }
  else{   
    sheet.getRange(4,1,cleanedEmailArray.length,2).setValues(cleanedEmailArray).sort(2);
  }
  
  //sheet.getRange(4,1,emailArray.length,2).setValues(emailArray).sort(2);
}