function getSheetById(sheet_id) {
  // accessing the workbook
    var wBook = SpreadsheetApp.getActiveSpreadsheet();
  
  // access all the sheets in the workbook
    var sheets = wBook.getSheets();

  // locate sheets variable
    for(i in sheets){
      if(sheets[i].getSheetById == sheet_id){
        var sheetName = sheets[i].getSheetName();
      }
    }

  // return the wb and sheet name
    return wBook.getSheetByName(sheetName);

}
