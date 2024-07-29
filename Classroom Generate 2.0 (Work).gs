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