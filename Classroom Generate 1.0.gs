OWNER_EMAIL = "xinyg04@gmail.com";

function classroomData(properties){

  const crs = Classroom.newCourse();
  
  Object.keys(properties).forEach(key => {
    crs[key] = properties[key];
  })

  const createdCourse = Classroom.Courses.create(crs);

  return createdCourse.enrollmentCode;

}

function classroomGenerate() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Department');
  const data = ss.getRange(2, 1, ss.getLastRow() - 1, 2).getValues();

  const enrollmentCodes = [];
  const nameIndex = 0;
  const codeIndex = 1;
  
  data.forEach(row => {
    if (row[codeIndex] == "") {

      const eCode = classroomData({
      name: row[nameIndex],
      ownerId: OWNER_EMAIL
      });

      enrollmentCodes.push([eCode]);

    } else {

      enrollmentCodes.push([row[codeIndex]]);

    }
  });

  ss.getRange(2, 2, enrollmentCodes.length, 1).setValues(enrollmentCodes);

}