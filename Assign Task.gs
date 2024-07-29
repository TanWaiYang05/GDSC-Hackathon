function calendarTask() {
  let sheet = SpreadsheetApp.getActiveSheet();

  let dataRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("tasks");
  // if the end time exists
  let filterColumn = 2; 

  let events = dataRange.getValues().filter(row => row[filterColumn]); 

  //Logger.log(events);

  events.forEach(function(e,index){
    if(!e[7]){ // if the event existed using 
      CalendarApp.getCalendarById("0369ee4a54be7f9d1db933910cb3f4b3373f2b8076e80e32002afcbe0d3e370a@group.calendar.google.com").createEvent(
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
      CalendarApp.getCalendarById("0369ee4a54be7f9d1db933910cb3f4b3373f2b8076e80e32002afcbe0d3e370a@group.calendar.google.com").createAllDayEvent(
        e[0],
        e[1],
        { description: e[3],
          guests: e[8],sendInvites: true}
      );
      guests = e[8].split(',').map(guest => guest.trim());
      Logger.log(guests);
      let newIndex = index+7;
      sheet.getRange("H"+newIndex).setValue(true)
    }
  })
}
