function myFunction() {

  var formResponses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmployeeInfo');
  var formData = formResponses.getRange(1,1,formResponses.getLastRow(),formResponses.getLastColumn()).getDisplayValues();

  Logger.log(formData);
  
}


