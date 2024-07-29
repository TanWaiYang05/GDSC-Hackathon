function classroomInvite() {

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

  employeeData.forEach(rowEmp => {
    if (rowEmp[statusIndexEmp] == "Not Send"){

      departmentData.forEach(rowDept => {
        if(rowEmp[departmentIndexEmp] == rowDept[departmentIndexDept]){
          courseId = rowDept[courseIdIndexDept].toString();
        }
      });

      Classroom.Invitations.create({
        "userId": rowEmp[emailIndexEmp],
        "courseId": courseId,
        "role": "STUDENT"
      });
    }
  });
  
} //important must work with sending email, which mean after invite status need to be change to "Sent".