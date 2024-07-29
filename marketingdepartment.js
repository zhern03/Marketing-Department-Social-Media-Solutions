// Function to validate form entries
function validateEntry() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shPostRequirement = myGoogleSheet.getSheetByName("Post Requirement");
  var ui = SpreadsheetApp.getUi();

  // Reset background colors
  shPostRequirement.getRange("C4").setBackground('#FFFFFF');
  shPostRequirement.getRange("C6").setBackground('#FFFFFF');
  shPostRequirement.getRange("C8").setBackground('#FFFFFF');
  shPostRequirement.getRange("C10").setBackground('#FFFFFF');
  shPostRequirement.getRange("C12").setBackground('#FFFFFF');
  shPostRequirement.getRange("C14").setBackground('#FFFFFF');

  // Validate event title
  if (shPostRequirement.getRange("C4").isBlank()) {
    ui.alert("Please enter an event title.");
    shPostRequirement.getRange("C4").activate().setBackground('#FF0000');
    return false;
  }

  // Validate content
  if (shPostRequirement.getRange("C6").isBlank()) {
    ui.alert("Please enter the content of the post.");
    shPostRequirement.getRange("C6").activate().setBackground('#FF0000');
    return false;
  }

  // Validate link
  if (shPostRequirement.getRange("C8").isBlank()) {
    ui.alert("Please enter the link of the post.");
    shPostRequirement.getRange("C8").activate().setBackground('#FF0000');
    return false;
  }

  // Validate time
  if (shPostRequirement.getRange("C10").isBlank()) {
    ui.alert("Please select a time that this post is to be posted.");
    shPostRequirement.getRange("C10").activate().setBackground('#FF0000');
    return false;
  }

  // Validate date
  if (shPostRequirement.getRange("C12").isBlank()) {
    ui.alert("Please choose a date that this post is to be posted.");
    shPostRequirement.getRange("C12").activate().setBackground('#FF0000');
    return false;
  }

  return true;
}

// Function to submit the data to the scheduling sheet and create a calendar event
function submitData() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shPostRequirement = myGoogleSheet.getSheetByName("Post Requirement");
  var datasheet = myGoogleSheet.getSheetByName("Scheduling");
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert("Submit", "Are you sure to submit the data?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.NO) {
    return;
  }

  if (validateEntry()) {
    var blankRow = datasheet.getLastRow() + 1;

    // Get the values from the form
    var eventTitle = shPostRequirement.getRange("C4").getValue();
    var eventContent = shPostRequirement.getRange("C6").getValue();
    var eventLink = shPostRequirement.getRange("C8").getValue();
    var eventTime = shPostRequirement.getRange("C10").getDisplayValue();
    var eventDate = shPostRequirement.getRange("C12").getValue();
    var eventInvitation = shPostRequirement.getRange("C14").getValue();

    // Combine date and time for the event
    var formattedDate = Utilities.formatDate(new Date(eventDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var startTime = new Date(formattedDate + "T" + eventTime + ":00");
    var endTime = new Date(startTime.getTime() + (1 * 60 * 60 * 1000)); // Assuming 1 hour event duration

    // Add data to the scheduling sheet
    datasheet.getRange(blankRow, 1).setValue(eventTitle);
    datasheet.getRange(blankRow, 2).setValue(eventContent);
    datasheet.getRange(blankRow, 3).setValue(eventLink);
    datasheet.getRange(blankRow, 4).setValue(eventTime);
    datasheet.getRange(blankRow, 5).setValue(eventDate);
    datasheet.getRange(blankRow, 6).setValue(eventInvitation);

    // Prepare guests list
    var guestsList = [];
    if (eventInvitation) {
      guestsList = eventInvitation.split(',').map(function(guest) {
        return guest.trim();
      });
    }

    // Create a calendar event with guests
    var userEmail = Session.getActiveUser().getEmail();
    var calendar = CalendarApp.getCalendarById(userEmail);
    var event = calendar.createEvent(eventTitle, startTime, endTime, {
      description: eventContent + "\n" + eventLink,
      guests: guestsList.join(',')
    });

    // Clear the form
    shPostRequirement.getRange("C4").clear();
    shPostRequirement.getRange("C6").clear();
    shPostRequirement.getRange("C8").clear();
    shPostRequirement.getRange("C10").clear();
    shPostRequirement.getRange("C12").clear();
    shPostRequirement.getRange("C14").clear();
  }
}

// Function to clear data from the form
function cleardata() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shPostRequirement = myGoogleSheet.getSheetByName("Post Requirement");
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Clear", "Are you sure to clear the data?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.NO) {
    return;
  }

  if (response == ui.Button.YES) {
    shPostRequirement.getRange("C4").clear();
    shPostRequirement.getRange("C6").clear();
    shPostRequirement.getRange("C8").clear();
    shPostRequirement.getRange("C10").clear();
    shPostRequirement.getRange("C12").clear();
    shPostRequirement.getRange("C14").clear();
  }
}

function sendReminders() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = myGoogleSheet.getSheetByName("Scheduling");
  var today = new Date();
  var formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "M/d/yyyy");

  // Get all data from the scheduling sheet
  var dataRange = datasheet.getDataRange();
  var data = dataRange.getValues();

  Logger.log("Today's Date: " + formattedToday);

  for (var i = 1; i < data.length; i++) {
    var eventDate = Utilities.formatDate(new Date(data[i][4]), Session.getScriptTimeZone(), "M/d/yyyy");

    Logger.log("Checking event: " + data[i][0] + " on " + eventDate);

    if (eventDate == formattedToday) {
      Logger.log("Event matches today's date.");

      var emailTo = Session.getActiveUser().getEmail();
      var guests = data[i][5];
      
      var eventTitle = data[i][0];
      var eventContent = data[i][1];
      var eventLink = data[i][2];
      var eventTime = data[i][3];

      var subject = "Reminder: " + eventTitle;
      var message = "This is a reminder for your event:\n\n" +
                    "Title: " + eventTitle + "\n" +
                    "Content: " + eventContent + "\n" +
                    "Link: " + eventLink + "\n" +
                    "Time: " + eventTime + "\n" +
                    "Date: " + eventDate + "\n\n" +
                    "Best regards,\nYour Automated System";

      // Send email to the user
      MailApp.sendEmail(emailTo, subject, message);

      Logger.log("Reminder email sent to: " + emailTo);

      // Send email to the guests
      if (guests) {
        var guestEmails = guests.split(',').map(function(guest) {
          return guest.trim();
        });
        for (var j = 0; j < guestEmails.length; j++) {
          MailApp.sendEmail(guestEmails[j], subject, message);
          Logger.log("Reminder email sent to guest: " + guestEmails[j]);
        }
      }
    }
  }
}


// Function to create a time-driven trigger for sending reminders
function createReminderTrigger() {
  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
}
