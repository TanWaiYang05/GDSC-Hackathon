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
      "invitationId": invitationId
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
