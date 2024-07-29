function sendChatInvitations() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmployeeInfo"); // get Sheet Info
  var data = sheet.getDataRange().getValues();
  var hrEmail = "wynnchan12345@gmail.com"; // Replace with the HQ email address
  
  for (var i = 1; i < data.length; i++) { // Start from 1 to skip headers
    var name = data[i][0];
    var email = data[i][1];
    var team = data[i][4];
    var subject = "Invitation to Join Your Team's Google Chat";
    var body = createMail(name, email, team, hrEmail);
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
  }

  function createMail(name, email, team, hrEmail) {
    var mailtoLink = "mailto:" + hrEmail +
                     "?subject=" + encodeURIComponent("Request to Join Google Chat Space") +
                     "&body=" + encodeURIComponent(
                       "Hello,\n\n" +
                       "I would like to request access to join the Google Chat space for " + team + ".\n\n" +
                       "Employee Name: " + name + "\n" +
                       "Employee Email: " + email + "\n\n" +
                       "Thank you.\n\n" +
                       "Best regards,\n" +
                       name
                     );
    
    return "<p>Hello " + name + ",</p>" +
           "<p>Please join your team's Google Chat by sending a request using the following link: <a href='" + mailtoLink + "'>Request to Join Chat</a></p>" +
           "<p>Best regards,<br>HR Team</p>";
  }
}


function sendTrainingReminder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmployeeInfo");
  var data = sheet.getDataRange().getValues();

  var trainingDate = new Date("2024-07-30");
  var today = new Date();
  var reminderDate = new Date(trainingDate);
  reminderDate.setDate(reminderDate.getDate() - 1);

  if (today.getDate() == reminderDate.getDate() &&
      today.getMonth() == reminderDate.getMonth() &&
      today.getFullYear() == reminderDate.getFullYear()) {

    for (var i = 1; i < data.length; i++) {
      var email = data[i][1];
      var subject = "Reminder: Upcoming Training Session";
      var body = "This is a reminder for the upcoming training session on " + trainingDate.toDateString() + ".";

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body
      });
    }
  }
}











