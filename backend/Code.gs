/**
 * עצומת רחוב בר יהודה — קוד השרת (Google Apps Script)
 * שומר כל חתימה לגיליון Google Sheets ומחזיר את מספר החותמים בזמן אמת.
 * אין צורך לשנות דבר בקוד — רק להדביק, לשמור ולפרוס (לפי המדריך).
 */

var SHEET_NAME = 'חתימות';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var name  = String(data.name  || '').trim();
    var street= String(data.street|| '').trim();
    var house = String(data.house || '').trim();
    var phone = String(data.phone || '').trim();

    if (!name || !street || !house) {
      return reply({ ok: false, error: 'missing fields' });
    }

    var sheet = getSheet();

    // מניעת כפילויות לפי שם + רחוב + מספר בית
    var key = (name + '|' + street + '|' + house).toLowerCase();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var k = (String(values[i][1]) + '|' + String(values[i][2]) + '|' + String(values[i][3])).toLowerCase();
      if (k === key) {
        return reply({ ok: true, duplicate: true, count: getCount() });
      }
    }

    sheet.appendRow([ new Date(), name, street, house, phone ]);
    return reply({ ok: true, count: getCount() });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var cb = e && e.parameter && e.parameter.callback;
  return reply({ ok: true, count: getCount() }, cb);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['תאריך', 'שם מלא', 'רחוב', 'מספר בית', 'טלפון']);
  }
  return sh;
}

function getCount() {
  var sh = getSheet();
  return Math.max(0, sh.getLastRow() - 1);
}

function reply(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
