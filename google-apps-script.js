// ============================================================
// Google Apps Script – Webhook לקבלת תשובות סקרים (גרסה 2)
// ============================================================
//
// הוראות עדכון:
// 1. לכו ל-Google Sheet שלכם
// 2. Extensions → Apps Script
// 3. מחקו את כל הקוד הקיים והדביקו את הקוד הזה
// 4. שמרו (Ctrl+S)
// 5. Deploy → Manage Deployments → עריכה (אייקון עיפרון)
// 6. ב-Version בחרו "New version"
// 7. לחצו Deploy
//
// ============================================================

const SHEET_MAP = {
  kids: "ילדים ובני נוער",
  students: "סטודנטים",
  professionals: "אנשי מקצוע",
  public: "קהל רחב",
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var surveyType = data.survey;
    var sheetName = SHEET_MAP[surveyType];

    if (!sheetName) {
      return _jsonResponse({ error: "Unknown survey type: " + surveyType });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // הוספת timestamp והסרת שדה survey
    data["תאריך ושעה"] = new Date().toLocaleString("he-IL", { timeZone: "Asia/Jerusalem" });
    delete data.survey;

    // סידור: תאריך קודם
    var keys = Object.keys(data);
    keys = ["תאריך ושעה"].concat(keys.filter(function(k) { return k !== "תאריך ושעה"; }));

    // אם הגליון ריק – כותרות
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(keys);
      var headerRange = sheet.getRange(1, 1, 1, keys.length);
      headerRange.setFontWeight("bold").setBackground("#E8F0EB").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // קריאת כותרות קיימות
    var lastCol = sheet.getLastColumn();
    var existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

    // הוספת עמודות חדשות אם צריך
    keys.forEach(function(key) {
      if (existingHeaders.indexOf(key) === -1) {
        var newCol = existingHeaders.length + 1;
        sheet.getRange(1, newCol).setValue(key).setFontWeight("bold").setBackground("#E8F0EB");
        existingHeaders.push(key);
      }
    });

    // בניית שורה לפי סדר הכותרות
    var row = existingHeaders.map(function(header) {
      return data[header] !== undefined ? data[header] : "";
    });

    sheet.appendRow(row);

    // רוחב אוטומטי בפעם הראשונה
    if (sheet.getLastRow() === 2) {
      try {
        for (var i = 1; i <= row.length; i++) {
          sheet.autoResizeColumn(i);
        }
      } catch (err) {}
    }

    return _jsonResponse({ success: true, sheet: sheetName });
  } catch (err) {
    return _jsonResponse({ error: err.toString() });
  }
}

function doGet(e) {
  return _jsonResponse({
    status: "active",
    message: "Survey webhook is running",
    sheets: Object.values(SHEET_MAP),
  });
}

function _jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
