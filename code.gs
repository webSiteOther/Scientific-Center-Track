/**
 * منصة المركز العلمي - Backend API الشامل
 * تم التحديث ليشمل كافة عمليات الإدارة، المدرسين، والطلاب.
 */

const SHEET_ID = "1xAV86cXC7Ft2XkMCtDQlx4RSCmijw4mJEAV0b1B9MR0";

function doPost(e) {
  let response = { success: false, message: "حدث خطأ غير معروف" };
  try {
    let requestData = JSON.parse(e.postData.contents);
    let action = requestData.action; 

    // --- نظام توجيه العمليات (Routing) ---
    if (action === "login") {
      response = checkLogin(requestData.username, requestData.password, requestData.deptCode);
    } 
    else if (action === "addGroup") {
      response = addGroup(requestData.groupData);
    } 
    else if (action === "getGroups") {
      response = getInstructorGroups(requestData.instructorName);
    } 
    else if (action === "saveTracking") {
      response = saveSessionPerformance(requestData.sessionData);
    } 
    else if (action === "getDashboard") {
      response = getDashboardData(); 
    } 
    else if (action === "getReportData") {
      response = getStudentReportData(requestData.studentName, requestData.month);
    } 
    else if (action === "addStudent") {
      response = addStudent(requestData.instructorName, requestData.groupName, requestData.studentName);
    } 
    else if (action === "getStudents") {
      response = getGroupStudents(requestData.instructorName, requestData.groupName);
    } 
    else if (action === "deleteStudent") {
      response = deleteStudent(requestData.instructorName, requestData.groupName, requestData.studentName);
    }
    // التحديث الجديد: دعم جلب كافة بيانات لوحة الإدارة
    else if (action === "getAllAdminData") {
      response = getAllAdminData();
    }
    else {
      response = { success: false, message: "العملية المطلوبة غير مدعومة: " + action };
    }

  } catch (error) {
    response = { success: false, message: "خطأ في السيرفر: " + error.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) { 
  return ContentService.createTextOutput("API يعمل بنجاح! 🚀"); 
}

// ==========================================
// 1. دوال الإدارة والتحليل الشامل
// ==========================================

function getAllAdminData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const safeGetValues = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    return sheet ? sheet.getDataRange().getValues() : [];
  };

  const data = {
    users: safeGetValues("Users"),
    groups: safeGetValues("Groups"),
    students: safeGetValues("Students"),
    sessions: safeGetValues("SessionsData")
  };
  
  return { success: true, data: data };
}

// ==========================================
// 2. دوال إدارة الطلاب
// ==========================================

function addStudent(instructorName, groupName, studentName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Students");
  if (!sheet) throw new Error("شيت Students غير موجود. يرجى تشغيل setupDatabase");
  
  sheet.appendRow([new Date(), instructorName, groupName, studentName]);
  return { success: true, message: "تمت إضافة الطالب بنجاح" };
}

function getGroupStudents(instructorName, groupName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Students");
  if (!sheet) return { success: true, data: [] }; 
  
  const data = sheet.getDataRange().getValues();
  let students = [];
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][1] == instructorName && data[i][2] == groupName) {
      students.push(data[i][3]);
    }
  }
  return { success: true, data: students };
}

function deleteStudent(instructorName, groupName, studentName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Students");
  if (!sheet) return { success: false, message: "شيت الطلاب غير موجود" };

  const data = sheet.getDataRange().getValues();
  for(let i = data.length - 1; i >= 1; i--) {
    if(data[i][1] == instructorName && data[i][2] == groupName && data[i][3] == studentName) {
      sheet.deleteRow(i + 1); 
      return { success: true, message: "تم حذف الطالب بنجاح" };
    }
  }
  return { success: false, message: "الطالب غير موجود" };
}

// ==========================================
// 3. الدوال الأساسية (تسجيل الدخول، المجموعات، المتابعة)
// ==========================================

function checkLogin(username, password, deptCode) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] == username && data[i][4] == password && data[i][2] == deptCode) {
      return { success: true, role: data[i][5], name: data[i][0], dept: data[i][1] };
    }
  }
  return { success: false, message: "بيانات الدخول غير صحيحة." };
}

function addGroup(groupData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Groups");
  sheet.appendRow([new Date(), groupData.instructorName, groupData.groupName, groupData.days, groupData.hours]);
  return { success: true, message: "تمت إضافة المجموعة بنجاح" };
}

function getInstructorGroups(instructorName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Groups");
  const data = sheet.getDataRange().getValues();
  let groups = data.filter(row => row[1] === instructorName).map(row => ({
    groupName: row[2], days: row[3], hours: row[4]
  }));
  return { success: true, data: groups };
}

function saveSessionPerformance(performanceData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("SessionsData");
  const data = sheet.getDataRange().getValues();
  
  performanceData.forEach(student => {
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][1] == student.instructorName && data[i][2] == student.groupName && 
         data[i][3] == student.month && data[i][4] == student.sessionNumber && 
         data[i][5] == student.studentName) {
        rowIndex = i + 1; // 1-based index in sheets
        break;
      }
    }
    
    let rowData = [
      new Date(), student.instructorName, student.groupName, student.month,
      student.sessionNumber, student.studentName, student.attendance, 
      student.interaction, student.task, student.notes, student.examScore || 0,
      student.speaking || "", student.writing || "", student.listening || "", student.reading || ""
    ];
    
    if(rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 15).setValues([rowData]);
      data[rowIndex-1] = rowData; // update local cache
    } else {
      sheet.appendRow(rowData);
    }
  });
  return { success: true, message: "تم حفظ التقييمات بنجاح" };
}

function getDashboardData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("SessionsData");
  const data = sheet.getDataRange().getValues();
  return { success: true, data: data }; 
}

function getStudentReportData(studentName, month) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("SessionsData");
  const data = sheet.getDataRange().getValues();
  let studentSessions = data.filter(row => row[5] == studentName && row[3] == month);
  if (studentSessions.length === 0) return { success: false, message: "لا توجد بيانات لهذا الطالب." };
  
  const instructor = studentSessions[0][1];
  let sessionsForReport = [];
  for(let i = 1; i <= 8; i++) {
    let sessionRow = studentSessions.find(row => row[4] == i);
    if(sessionRow) {
      sessionsForReport.push({
        sessionNumber: i, attendance: sessionRow[6], interaction: sessionRow[7],
        task: sessionRow[8], notes: sessionRow[9], examScore: sessionRow[10] || "-",
        speaking: sessionRow[11] || "", writing: sessionRow[12] || "", listening: sessionRow[13] || "", reading: sessionRow[14] || ""
      });
    } else {
      sessionsForReport.push({ sessionNumber: i, attendance: "-", interaction: "-", task: "-", notes: "-", examScore: "-", speaking: "", writing: "", listening: "", reading: "" });
    }
  }
  return { success: true, studentName: studentName, month: month, department: "المركز العلمي", instructorName: instructor, sessions: sessionsForReport };
}

// ==========================================
// 4. دالة إعداد قاعدة البيانات
// ==========================================

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // شيت المستخدمين
  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(['الاسم', 'اسم القسم', 'رمز القسم', 'اسم المستخدم', 'كلمة المرور', 'الرتبة']);
    usersSheet.appendRow(['مدير النظام', 'الإدارة العامة', 'ADMIN01', 'admin', 'admin123', 'Admin']);
  }
  
  // شيت المجموعات
  let groupsSheet = ss.getSheetByName("Groups");
  if (!groupsSheet) {
    groupsSheet = ss.insertSheet("Groups");
    groupsSheet.appendRow(['التاريخ', 'اسم المدرب', 'اسم المجموعة', 'الأيام', 'الساعات']);
  }
  
  // شيت الجلسات
  let sessionsSheet = ss.getSheetByName("SessionsData");
  if (!sessionsSheet) {
    sessionsSheet = ss.insertSheet("SessionsData");
    sessionsSheet.appendRow(['التاريخ', 'اسم المدرب', 'المجموعة', 'الشهر', 'رقم السيشن', 'اسم الطالب', 'الحضور', 'التفاعل', 'التاسك', 'ملاحظات', 'درجة الامتحان', 'Speaking', 'Writing', 'Listening', 'Reading']);
  }
  
  // شيت الطلاب
  let studentsSheet = ss.getSheetByName("Students");
  if (!studentsSheet) {
    studentsSheet = ss.insertSheet("Students");
    studentsSheet.appendRow(['التاريخ', 'اسم المدرب', 'المجموعة', 'اسم الطالب']);
    studentsSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#d9ead3");
  }

  let sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1) ss.deleteSheet(sheet1);
  return "تم بناء هيكل قاعدة البيانات بنجاح!";
}