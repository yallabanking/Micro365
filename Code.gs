/**
 * ======================================================================
 *  كود ربط صفحة التسويق بشيت جوجل
 *  انسخ هذا الكود بالكامل داخل: Extensions > Apps Script (داخل الشيت نفسه)
 * ======================================================================
 *
 *  المطلوب من الشيت (Google Sheet) أن يحتوي على 4 تابات (Sheets) بهذه الأسماء بالضبط:
 *
 *  1) "الموظفين"   بالأعمدة:  A=رقم الأوراكل | B=اسم الموظف | C=الفئة
 *     الفئة يجب أن تكون واحدة من: تمويل | مسئول اول | مسئول جغرافي
 *
 *  2) "تمويل"        بالأعمدة:
 *     A=التاريخ والوقت | B=أوراكل الموظف | C=اسم الموظف | D=نوع العميل |
 *     E=اسم العميل | F=رقم الهاتف | G=النشاط | H=المبلغ المتوقع |
 *     I=نوع التسويق | J=إحداثيات الزيارة | K=تاريخ المتابعة القادم | L=ملاحظات
 *
 *  3) "مسئول اول"    بالأعمدة:
 *     A=التاريخ والوقت | B=أوراكل الموظف | C=اسم الموظف | D=نوع العميل |
 *     E=اسم العميل | F=رقم الهاتف | G=النشاط | H=المبلغ المنصرف |
 *     I=عنوان العميل | J=الإحداثيات | K=ملاحظات
 *
 *  4) "مسئول جغرافي" بالأعمدة:
 *     A=التاريخ والوقت | B=أوراكل الموظف (المسجّل دخوله) | C=اسم الموظف |
 *     D=أوراكل الموظف المسؤول عن العميل | E=نوع العميل | F=اسم العميل |
 *     G=رقم الهاتف | H=النشاط | I=المبلغ المنصرف | J=قسط القرض |
 *     K=عنوان العميل | L=الإحداثيات | M=ملاحظات
 *
 *  بعد لصق الكود:
 *   1. اضغط Deploy > New deployment
 *   2. اختر النوع: Web app
 *   3. Execute as: Me
 *   4. Who has access: Anyone
 *   5. اضغط Deploy وانسخ رابط "Web app URL"
 *   6. الصق الرابط داخل ملف الصفحة (index.html) في المتغير APPS_SCRIPT_URL
 * ======================================================================
 */

const EMPLOYEES_SHEET = "الموظفين";

function doGet(e) {
  const action = e.parameter.action;

  if (action === "lookupEmployee") {
    return lookupEmployee(e.parameter.oracle);
  }

  if (action === "getDashboardData") {
    return getDashboardData();
  }

  return jsonResponse({ status: "error", message: "unknown action" });
}

/**
 * Reads all three data sheets + the employees sheet and returns them
 * as arrays of objects (field-keyed) for the dashboard to consume.
 * Column order here MUST match the order the form actually writes in.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const financeFields = ["timestamp","empOracle","empName","clientType","clientName","phone","activity","expectedAmount","marketingType","location","nextFollowUp","notes"];
  const seniorFields  = ["timestamp","empOracle","empName","clientType","clientName","phone","activity","disbursedAmount","address","location","notes"];
  const geoFields     = ["timestamp","empOracle","empName","respEmpOracle","clientType","clientName","phone","activity","disbursedAmount","installment","address","location","notes"];
  const empFields     = ["oracle","name","role"];

  return jsonResponse({
    status: "success",
    generatedAt: new Date().toISOString(),
    finance: readSheetAsObjects(ss, "تمويل", financeFields),
    seniorOfficer: readSheetAsObjects(ss, "مسئول اول", seniorFields),
    geoOfficer: readSheetAsObjects(ss, "مسئول جغرافي", geoFields),
    employees: readSheetAsObjects(ss, EMPLOYEES_SHEET, empFields)
  });
}

function readSheetAsObjects(ss, sheetName, fieldKeys) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // no data rows, only (or no) header

  const numCols = fieldKeys.length;
  const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  return values
    .filter(row => row.some(cell => cell !== "" && cell !== null))
    .map(row => {
      const obj = {};
      fieldKeys.forEach((key, i) => {
        let val = row[i];
        if (val instanceof Date) val = val.toISOString();
        obj[key] = val;
      });
      return obj;
    });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(data.sheetName);

    if (!sheet) {
      return jsonResponse({ status: "error", message: "sheet not found: " + data.sheetName });
    }

    sheet.appendRow(data.row);
    return jsonResponse({ status: "success" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function lookupEmployee(oracle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EMPLOYEES_SHEET);

  if (!sheet) {
    return jsonResponse({ status: "error", message: "sheet not found: " + EMPLOYEES_SHEET });
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const rowOracle = String(values[i][0]).trim();
    if (rowOracle === String(oracle).trim() && rowOracle !== "") {
      return jsonResponse({
        status: "found",
        name: values[i][1],
        role: values[i][2]
      });
    }
  }

  return jsonResponse({ status: "not_found" });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
