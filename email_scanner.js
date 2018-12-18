var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

// add menu to Sheet
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Extract Emails')
        .addItem('Extract emails from inbox', 'extractEmailFromInbox')
        .addItem('Extract emails from label', 'extarctEmailFromLablel')
        .addToUi();
}

function extractEmailFromInbox() {
    var threads = GmailApp.getInboxThreads();
    extractEmails(threads);
}

function extarctEmailFromLablel() {
    var label = sheet.getRange(1, 1).getValue();
    var threads = GmailApp.search("label:" + label);
    extractEmails(threads);
}

function extractEmails(threads) {
    var messages = GmailApp.getMessagesForThreads(threads);
    var emailArray = [];

    // get array of email addresses
    messages.forEach(function(message) {
        message.forEach(function(d) {
            emailArray.push(d.getFrom());
        });
    });

    // de-duplicate the array
    var uniqueEmailArray = emailArray.filter(function(item, pos) {
        return emailArray.indexOf(item) == pos;
    });

    //Clean email to seperate email and name
    var cleanedEmailArray = uniqueEmailArray.map(function(el) {
        var name = "";
        var email = "";
        var matches = el.match(/\s*"?([^"]*)"?\s+<(.+)>/);

        if (matches) {
            name = matches[1];
            email = matches[2];
        } else {
            name = "N/k";
            email = el;
        }
        return [name, email];

    }).filter(function(d) {
        if (
            d[1] !== "zac@zachariasmanuel.com" &&
            d[1] !== "drive-shares-noreply@google.com" &&
            d[1] !== "mailer-daemon@googlemail.com"
        ) {
            return d;
        }
    });

    //Show alert if there is no email
    if (cleanedEmailArray.length == 0) {
        var ui = SpreadsheetApp.getUi();
        var result = ui.alert(
            'Alert!',
            'There is no email.',
            ui.ButtonSet.OK);
    } else {
        sheet.getRange(3, 1, cleanedEmailArray.length, 2).setValues(cleanedEmailArray).sort(2);
    }
}