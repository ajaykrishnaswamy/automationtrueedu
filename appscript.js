function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Scripts')
    .addItem('Create Trigger', 'createTrigger')
    .addToUi();
}

function sendNotificationEmail(e) {
  Logger.log("sendNotificationEmail triggered.");

  var formResponses = e.values;
  Logger.log("Form responses: " + JSON.stringify(formResponses));
 
  var timestamp = formResponses[0];
  var firstname = formResponses[1]; // Adjust index if needed
  var lastname = formResponses[2]; // Adjust index if needed

  Logger.log("Timestamp: " + timestamp);
  Logger.log("First Name: " + firstname);
  Logger.log("Last Name: " + lastname);
 
  var subject = "New Form Submission - Call Needed";
  var body = "New Form submission received:\n\n"
           + "First Name: " + firstname + "\n"
           + "Last Name: " + lastname + "\n"
           + "Submitted at: " + timestamp + "\n\n"
           + "Please call them ASAP.";

  Logger.log("Email Subject: " + subject);
  Logger.log("Email Body: " + body);

  MailApp.sendEmail("info@trueeducationcumming.com", subject, body);

  Logger.log("Email sent successfully to info@trueeducationcumming.com");

  // Add Google Calendar Event
  createCalendarEvent(firstname, lastname, timestamp);
}

function createCalendarEvent(firstname, lastname, timestamp) {
  Logger.log("Creating Calendar Event...");

  var calendarId = "primary"; // Use "primary" or replace with a specific Calendar ID
  var calendar = CalendarApp.getCalendarById(calendarId);

  var startTime = new Date(timestamp);
  var endTime = new Date(startTime.getTime() + 30 * 60 * 1000); // 30 minutes later

  var eventTitle = "Call Needed: " + firstname + " " + lastname;
  var eventDescription = "Follow up call needed for " + firstname + " " + lastname + ".";

  var event = calendar.createEvent(eventTitle, startTime, endTime, {
    description: eventDescription
  });

  Logger.log("Event created: " + event.getTitle() + " at " + event.getStartTime());
}

function createTrigger() {
  Logger.log("Creating trigger for sendNotificationEmail...");
  
  ScriptApp.newTrigger('sendNotificationEmail')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  Logger.log("Trigger created successfully.");
}
