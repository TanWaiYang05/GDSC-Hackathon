function onOpen(){
  const menu = SpreadsheetApp.getUi().createMenu("Training");
  menu.addItem("Generate Classroom", "classroomGenerate");
  menu.addItem("Send Invite", "sendInvites");
  menu.addItem("Assign Task with End Day", "calendarTask");
  menu.addItem("Assign Task", "calendarStartTask");
  menu.addItem("Create Meeting", "createCalendarEvent");
  menu.addItem("Chat Invite", "sendChatInvitations");
  menu.addItem("Traning Reminder", "sendTrainingReminder")
  menu.addToUi();
}

function sendInvites() {
  // Get Sheet
    var formResponses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmployeeInfo');
    var formRange = formResponses.getRange(1,1,formResponses.getLastRow(),formResponses.getLastColumn());
    var formData = formRange.getDisplayValues();
  
  
    const ssEmployeeInfo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmployeeInfo");
    const employeeData = ssEmployeeInfo.getRange(2, 2, ssEmployeeInfo.getLastRow() - 1, 5).getValues();

    const ssDepartment = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Department");
    const departmentData = ssDepartment.getRange(2, 1, ssDepartment.getLastRow() - 1, 3).getValues();

    const emailIndexEmp = 0;
    const statusIndexEmp = 1;
    const departmentIndexEmp = 4;
    const courseIdIndexDept = 2;
    const departmentIndexDept = 0;
    var courseId;

  // Filter unDone Employees
    var filteredData = formData.filter(row => row[2] === "Not Send");

  // Function to modify cell value (replace with your logic)
    function modifyCellValue(row, rowIndex) {
  // Modify the value based on your needs (replace with your desired change)
      row[2] = "Sent";

  // Write the modified row back to the sheet (adjust column index as needed)
      formResponses.getRange(rowIndex + 1, 3).setValue(row[2]);
    }

  // Loop through filtered data and modify corresponding cells
    filteredData.forEach((row, i) => {
  // Find the matching row index in the original data (replace with your logic)
      const originalRowIndex = formData.findIndex(originalRow => originalRow[0] === row[0]);
      if (originalRowIndex !== -1) {
        modifyCellValue(row.slice(), originalRowIndex); // Pass a copy of the row to avoid modifying filteredRows
      }
    });

//--------------------------
  // Create Template Object
    var htmlTemplate = HtmlService.createTemplateFromFile('Invitation Email');


  // send email
   employeeData.forEach(rowEmp => {
    if (rowEmp[statusIndexEmp] == "Not Send"){
      Logger.log(rowEmp);
      departmentData.forEach(rowDept => {
        if(rowEmp[departmentIndexEmp] == rowDept[departmentIndexDept]){
          courseId = rowDept[courseIdIndexDept].toString();
        }
      });

    const invitationId = "Training";

    if (courseId) {
    Classroom.Invitations.create({
      "userId": rowEmp[emailIndexEmp],
      "courseId": courseId,
      "role": "STUDENT",
      //"invitationId": invitationId
    });

    htmlTemplate.link = `https://classroom.google.com/invitation?id=${invitationId}`;
    }


    // evaluate the template
      var htmlForEmail = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(
      rowEmp[emailIndexEmp], // Email Recipients
      'Congratulations! Welcome to the Team at CyberTerror',
      'This email contains html',
      {htmlBody: htmlForEmail}
    )
    }
  });
}

function createCalendarEvent() {
  // Get the active sheet and the range
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the named range "events"
  let range = sheet.getRange("A2:H" + sheet.getLastRow()); // Assuming headers in the first row and data starts from the second row
  let events = range.getValues();
  
  // Loop through each event and create it in the calendar if not already scheduled
  events.forEach(function(e, index) {
    if (!e[7]) { // Only schedule if the checkbox is not ticked
      CalendarApp.getCalendarById("40d7929f97af193409510337d20b27c063a409480e3ce6e2eaa9a4f2ae6cd3f1@group.calendar.google.com").createEvent(
        e[0],
        new Date(e[1]),
        new Date(e[2]),
        {guests: e[6], sendInvites: true}
      );
      // Mark the event as scheduled by ticking the checkbox
      sheet.getRange(index + 2, 8).setValue(true); // Assuming headers in the first row and checkboxes in the 8th column
    }
  });
}

function calendarTask() {
  let sheet = SpreadsheetApp.getActiveSheet();

  let dataRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("tasks");
  // if the end time exists
  let filterColumn = 2; 

  let events = dataRange.getValues().filter(row => row[filterColumn]); 

  //Logger.log(events);

  events.forEach(function(e,index){
    if(!e[7]){ // if the event existed using 
      CalendarApp.getCalendarById("40d7929f97af193409510337d20b27c063a409480e3ce6e2eaa9a4f2ae6cd3f1@group.calendar.google.com").createEvent(
        e[0],
        new Date(e[1]),
        new Date(e[2]),
        //CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.THURSDAY).until(e[2]);
        { 
          description: e[3],
          guests: e[8],
          sendInvites: true
        }
      );
      let newIndex = index+7;
      sheet.getRange(2,8,events.length,1).setValue(true)
    }
  })
}

function calendarStartTask() {
  let sheet = SpreadsheetApp.getActiveSheet();

  let dataRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("tasks");
  // if the end time exists
  let filterColumn = 2; 

  let events = dataRange.getValues().filter(row => !row[filterColumn]); 

  //Logger.log(events);

  events.forEach(function(e,index){
    if(!e[7]){ // if the event existed using checkbox
      CalendarApp.getCalendarById("40d7929f97af193409510337d20b27c063a409480e3ce6e2eaa9a4f2ae6cd3f1@group.calendar.google.com").createAllDayEvent(
        e[0],
        e[1],
        { description: e[3],
          guests: e[8],sendInvites: true}
      );
      guests = e[8].split(',').map(guest => guest.trim());
      //Logger.log(guests);
      let newIndex = index+7;
      sheet.getRange("H"+newIndex).setValue(true)
    }
  })
}

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

OWNER_EMAIL = "xinyg04@gmail.com";

function classroomData(properties){

  const crs = Classroom.newCourse();
  
  Object.keys(properties).forEach(key => {
    crs[key] = properties[key];
  })

  const createdCourse = Classroom.Courses.create(crs);

  return [createdCourse.enrollmentCode, createdCourse.id];

}

function classroomGenerate() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Department');
  const data = ss.getRange(2, 1, ss.getLastRow() - 1, 3).getValues();

  const enrollmentCodes = [];
  const courseIdList = [];
  const nameIndex = 0;
  const codeIndex = 1;
  const idIndex = 2;
  
  data.forEach(row => {
    if (row[codeIndex] == "") {

      const [eCode, courseId] = classroomData({
        name: row[nameIndex],
        ownerId: OWNER_EMAIL
      });

      enrollmentCodes.push([eCode]);
      courseIdList.push([courseId]);

    } else {

      enrollmentCodes.push([row[codeIndex]]);
      courseIdList.push([row[idIndex]]);

    }
  });

  ss.getRange(2, 2, enrollmentCodes.length, 1).setValues(enrollmentCodes);
  ss.getRange(2, 3, courseIdList.length, 1).setValues(courseIdList);

}