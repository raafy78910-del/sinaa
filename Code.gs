/**
 * نظام شؤون الموظفين المدرسي — الخادم (Backend)
 * يعمل فوق Google Sheets مجانًا بالكامل عبر Google Apps Script.
 *
 * طريقة الإعداد بالتفصيل موجودة في ملف SETUP_GUIDE.md
 *
 * هذا الملف يحوّل ملف Google Sheets إلى "قاعدة بيانات" بسيطة،
 * كل ورقة (Sheet) تمثل جدول بيانات، والكود بالأسفل يوفر عمليات:
 * تسجيل دخول، عرض، إضافة، تعديل، حذف — لأي ورقة، مع تحقق من الصلاحيات
 * حسب دور المستخدم (مدير / سكرتير / اطّلاع).
 */

// ============== مصفوفة الصلاحيات (يجب أن تطابق نفس المصفوفة في index.html) ==============
// full  = عرض + إضافة + تعديل + حذف
// edit  = عرض + إضافة + تعديل (بدون حذف)
// view  = عرض فقط
// none  = لا صلاحية إطلاقًا
var PERMISSIONS = {
  Admin:     { employees: 'full', formations: 'full', leaves: 'full', transport: 'full', transfers: 'full', assignments: 'full', users: 'full' },
  Secretary: { employees: 'edit', formations: 'edit', leaves: 'edit', transport: 'edit', transfers: 'edit', assignments: 'view', users: 'none' },
  Viewer:    { employees: 'view', formations: 'view', leaves: 'view', transport: 'view', transfers: 'view', assignments: 'view', users: 'none' }
};

var SHEET_TO_MODULE = {
  Employees: 'employees',
  Formations: 'formations',
  Leaves: 'leaves',
  Transport: 'transport',
  Transfers: 'transfers',
  AssignmentsViolations: 'assignments',
  Users: 'users'
};

function hasPermission(role, sheetName, action) {
  var moduleKey = SHEET_TO_MODULE[sheetName];
  if (!moduleKey) return true; // أوراق غير محمية مباشرة (مثل Logs)
  var level = (PERMISSIONS[role] || {})[moduleKey] || 'none';
  if (action === 'list') return level !== 'none';
  if (action === 'add' || action === 'update') return level === 'edit' || level === 'full';
  if (action === 'delete') return level === 'full';
  return false;
}

// ============== تشفير كلمات المرور ==============
// كلمة المرور تصل من الواجهة وهي مُشفّرة مسبقًا (SHA-256) بنفس الخوارزمية هون،
// لذلك لا تُرسل كلمة المرور الحقيقية عبر الشبكة ولا تُخزّن كنص عادي بالشيت.
function hashPassword(username, password) {
  var raw = String(password) + ':' + String(username).toLowerCase();
  var digestBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digestBytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// ============== الإعداد الأولي (نفّذها مرة واحدة فقط) ==============
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetsConfig = {
    'Users': ['id', 'username', 'password', 'fullName', 'role', 'active'],
    'Employees': ['id', 'fullName', 'nationalId', 'jobTitle', 'subject', 'classAssigned', 'phone', 'hireDate', 'status', 'notes'],
    'Formations': ['id', 'teacherName', 'type', 'classSection', 'subject', 'hours', 'semester', 'year', 'notes'],
    'Leaves': ['id', 'employeeId', 'employeeName', 'leaveType', 'startDate', 'endDate', 'status', 'notes'],
    'Transport': ['id', 'adminUnit', 'employeeName', 'nationalId', 'workplace', 'jobTitle', 'residence', 'neighborhood', 'nearLandmark', 'effectiveDate', 'requestDate', 'routesJson', 'totalPrice', 'notes'],
    'Transfers': ['id', 'employeeName', 'type', 'date', 'fromSchool', 'toSchool', 'decisionNumber', 'status', 'notes'],
    'AssignmentsViolations': ['id', 'employeeName', 'recordType', 'title', 'date', 'decisionNumber', 'status', 'notes'],
    'Logs': ['timestamp', 'username', 'action', 'details']
  };

  Object.keys(sheetsConfig).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    sh.appendRow(sheetsConfig[name]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sheetsConfig[name].length).setFontWeight('bold');
  });

  // مستخدم مدير افتراضي — كلمة المرور: admin123 (مخزّنة كهاش وليست نصًا عاديًا)
  var usersSheet = ss.getSheetByName('Users');
  if (usersSheet.getLastRow() < 2) {
    usersSheet.appendRow([1, 'admin', hashPassword('admin', 'admin123'), 'المدير العام', 'Admin', true]);
  }

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert(
    'تم إنشاء الجداول بنجاح ✅\n\nاسم المستخدم: admin\nكلمة المرور: admin123\n\nغيّرها بعد أول تسجيل دخول من شاشة إدارة المستخدمين.'
  );
}

// ============== نقطة الدخول: طلبات POST ==============
function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var sheetName = body.sheet;

    if (action !== 'login') {
      if (!hasPermission(body.role, sheetName, action)) {
        return respond({ success: false, message: 'ليس لديك صلاحية لتنفيذ هذا الإجراء' });
      }
    }

    switch (action) {
      case 'login':
        result = login(body.username, body.password);
        break;
      case 'list':
        result = listRows(sheetName);
        break;
      case 'add':
        result = addRow(sheetName, body.data);
        logAction(body.username, 'إضافة', sheetName);
        break;
      case 'update':
        result = updateRow(sheetName, body.data);
        logAction(body.username, 'تعديل', sheetName);
        break;
      case 'delete':
        result = deleteRow(sheetName, body.id);
        logAction(body.username, 'حذف', sheetName);
        break;
      default:
        result = { success: false, message: 'إجراء غير معروف' };
    }
  } catch (err) {
    result = { success: false, message: 'خطأ في الخادم: ' + err.message };
  }
  return respond(result);
}

function doGet(e) {
  return respond({ success: true, message: 'الخادم يعمل ✅' });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============== أدوات مساعدة عامة ==============
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('الورقة غير موجودة: ' + name);
  return sh;
}

function sheetToObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 1) return [];
  var headers = values[0];
  return values.slice(1)
    .filter(function (r) { return r.join('') !== ''; })
    .map(function (r) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
}

function nextId(sheet) {
  var data = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var v = Number(data[i][0]);
    if (!isNaN(v) && v > max) max = v;
  }
  return max + 1;
}

function logAction(username, action, sheetName) {
  try {
    var sh = getSheet('Logs');
    sh.appendRow([new Date(), username || '-', action, sheetName]);
  } catch (e) { /* تجاهل أي خطأ بالتسجيل حتى لا يوقف العملية الأساسية */ }
}

// ============== العمليات الأساسية (CRUD) ==============
function listRows(sheetName) {
  var sheet = getSheet(sheetName);
  var rows = sheetToObjects(sheet);
  if (sheetName === 'Users') {
    // لا نُرجع كلمة المرور (ولا حتى الهاش) ضمن قوائم العرض
    rows = rows.map(function (u) {
      return { id: u.id, username: u.username, fullName: u.fullName, role: u.role, active: u.active };
    });
  }
  return { success: true, data: rows };
}

function addRow(sheetName, data) {
  var sheet = getSheet(sheetName);
  var headers = sheet.getDataRange().getValues()[0];
  var id = nextId(sheet);
  data = data || {};
  data.id = id;
  var row = headers.map(function (h) { return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(row);
  return { success: true, id: id };
}

function updateRow(sheetName, data) {
  var sheet = getSheet(sheetName);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idCol = headers.indexOf('id');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(data.id)) {
      var row = headers.map(function (h) {
        return data[h] !== undefined ? data[h] : values[i][headers.indexOf(h)];
      });
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true };
    }
  }
  return { success: false, message: 'لم يتم العثور على السجل' };
}

function deleteRow(sheetName, id) {
  var sheet = getSheet(sheetName);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idCol = headers.indexOf('id');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'لم يتم العثور على السجل' };
}

function login(username, password) {
  var sheet = getSheet('Users');
  var users = sheetToObjects(sheet);
  // "password" الواصلة هون هي بالفعل هاش SHA-256 محسوبة بالواجهة، فالمقارنة نصّية مباشرة
  var found = users.find(function (u) {
    return String(u.username) === String(username) && String(u.password) === String(password);
  });
  if (!found) return { success: false, message: 'اسم المستخدم أو كلمة المرور غير صحيحة' };
  if (found.active === false || String(found.active).toUpperCase() === 'FALSE') {
    return { success: false, message: 'هذا الحساب غير مُفعّل' };
  }
  return {
    success: true,
    user: { id: found.id, username: found.username, fullName: found.fullName, role: found.role }
  };
}
