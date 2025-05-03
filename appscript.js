// ====== MYSTUDIO CONFIGURATION ======
// Replace with your MyStudio API endpoint and API key
var MYSTUDIO_ENDPOINT = 'https://api.mystudio.com/v1/submissions';
var MYSTUDIO_API_KEY = 'YOUR_API_KEY_HERE';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Scripts')
    .addItem('Create Trigger', 'createTrigger')
    .addToUi();
}

function sendNotificationEmail(e) {
  Logger.log("sendNotificationEmail triggered.");

  // Pull each form field
  var values = e.values;
  var timestamp       = values[0];
  var email           = values[1];
  var parentFirstName = values[2];
  var parentLastName  = values[3];
  var parentPhone     = values[4];
  var parentEmail     = values[5];
  var studentName     = values[6];

  Logger.log("Submission: " + values.join(', '));

  // 1) Send Notification Email
  var subject = "New Form Submission - Call Needed";
  var body =  "New Form Submission Details:\n\n"
            + "Submitted At: " + timestamp       + "\n"
            + "Submitter Email: " + email         + "\n"
            + "Parent Name: "      + parentFirstName + " " + parentLastName + "\n"
            + "Parent Phone: "     + parentPhone     + "\n"
            + "Parent Email: "     + parentEmail     + "\n"
            + "Student Name: "     + studentName     + "\n\n"
            + "Please follow up ASAP.";
  MailApp.sendEmail("info@trueeducationcumming.com", subject, body);

  // 2) Create Calendar Event
  createCalendarEvent(parentFirstName, parentLastName, timestamp);

  // 3) Log to a 'Submission Log' sheet
  logToSheet([timestamp, email, parentFirstName, parentLastName, parentPhone, parentEmail, studentName]);

  // 4) Push data to MyStudio via API
  pushToMyStudio({
    timestamp: timestamp,
    email: email,
    parentFirstName: parentFirstName,
    parentLastName: parentLastName,
    parentPhone: parentPhone,
    parentEmail: parentEmail,
    studentName: studentName
  });
}

function createCalendarEvent(firstName, lastName, timestamp) {
  Logger.log("Creating calendar event...");

  var calendar = CalendarApp.getDefaultCalendar();
  var startTime = new Date(timestamp);
  var endTime = new Date(startTime.getTime() + 30 * 60 * 1000);
  var title = "Call Needed: " + firstName + " " + lastName;
  var description = "Follow-up call needed for " + firstName + " " + lastName + ".";

  calendar.createEvent(title, startTime, endTime, { description: description });
  Logger.log("Calendar event created: " + title);
}

function logToSheet(rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Submission Log') || ss.insertSheet('Submission Log');
  logSheet.appendRow(rowData);
  Logger.log("Logged submission to 'Submission Log'.");
}

function pushToMyStudio(data) {
  try {
    var payload = JSON.stringify(data);
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + MYSTUDIO_API_KEY
      },
      payload: payload,
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(MYSTUDIO_ENDPOINT, options);
    Logger.log('MyStudio API response: ' + response.getResponseCode() + ' - ' + response.getContentText());
  } catch (err) {
    Logger.log('Error pushing to MyStudio: ' + err);
  }
}

function createTrigger() {
  Logger.log("Creating trigger for sendNotificationEmail...");
  ScriptApp.newTrigger('sendNotificationEmail')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  Logger.log("Trigger created.");
}
