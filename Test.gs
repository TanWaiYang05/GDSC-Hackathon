function main() {
  // Get Sheet
    var formResponses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmployeeInfo');
    var formData = formResponses.getRange(1,1,formResponses.getLastRow(),formResponses.getLastColumn()).getDisplayValues();
  
  // Filter unDone Employees
    var filteredData = formData.filter(row => row[2] === "Not Send");

  // Get Emails
    for(i in filteredData){
      Logger.log(filteredData[i][1]);
    }

//--------------------------
  // Create Template Object
    var htmlTemplate = HtmlService.createTemplateFromFile('Invitation Email');

    htmlTemplate.link = 'https://www.youtube.com/watch?v=dQw4w9WgXcQ';

  // evaluate the template
    var htmlForEmail = htmlTemplate.evaluate().getContent();

  // send email
    GmailApp.sendEmail(
      filteredData[0][1],
      'Congratulations! Welcome to the Team at CyberTerror',
      'This email contains html',
      {htmlBody: htmlForEmail}
    )

}
