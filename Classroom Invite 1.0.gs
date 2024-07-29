function classroomInvite() {

  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmployeeInfo");
  const data = ss.getRange(2, 2, ss.getLastRow() - 1, 5).getValues();

  console.log(data); //test

  const emailIndex = 0;
  const statusIndex = 1;
  const departmentIndex = 4;

  const courseId = "701577854848";

  data.forEach(row => {
    if (row[statusIndex] == "Not Send"){
      console.log(row[emailIndex]); //test
      Classroom.Courses.Students.create({
        "userId": row[emailIndex],
      }, courseId)
    }
  });
}