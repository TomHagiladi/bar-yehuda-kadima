/**
 * חתימות בתי האב על תכנית המועצה — רחוב בר יהודה חד-סטרי (אוקטובר 2026)
 * קוד שרת (Google Apps Script) שכותב לגיליון "עצומת בר יהודה — חתימות תושבים".
 *
 * כלל הזהב: חתימה אחת לכל בית אב (רחוב + מספר בית).
 * בית שאין לו מספר במאגר נבחר במפה, ואז המפתח שלו הוא המבנה במפה.
 * חתימה שנייה לאותו בית אב לא נספרת — היא נרשמת בלשונית "כפולים" לבדיקה ידנית.
 * אפשר לחתום "בעד" או "בעד עם השגה"; השגה מחייבת נימוק.
 *
 * GET  ?action=list[&callback=fn]  -> מספר בתי האב שחתמו + רשימת הכתובות (בלי שמות ובלי השגות)
 * POST {name, street, house, loc, consent, stance, topics, section, reason, hp} -> רישום חתימה
 */

var SPREADSHEET_ID = '1mTVUAgms3g_WAktwlNGoefmj87bjeONKHHi4FMHKQO4';
var SHEET_MAIN = 'תכנית המועצה 2026';
var SHEET_DUP  = 'כפולים - תכנית המועצה';
var HEADERS_MAIN = ['תאריך', 'שם החותם', 'רחוב', 'מספר בית', 'מפתח בית אב', 'עמדה',
                    'נושאי ההשגה', 'באיזה קטע', 'פירוט ונימוק', 'מבנה במפה'];
var HEADERS_DUP  = ['תאריך', 'שם החותם', 'רחוב', 'מספר בית', 'מפתח בית אב', 'עמדה',
                    'נושאי ההשגה', 'באיזה קטע', 'פירוט ונימוק', 'מבנה במפה', 'הערה'];
var COL_KEY = 5, COL_STREET = 3, COL_LOC = 10;
var CACHE_KEY = 'plan_list_v2';

function doGet(e) {
  var p = (e && e.parameter) || {};
  return reply(listPayload(), p.callback);
}

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.hp) return reply({ ok: true });                       // מלכודת לרובוטים
    var name   = clean(data.name, 80);
    var street = clean(data.street, 60);
    var house  = clean(data.house, 20);
    var loc    = clean(data.loc, 30).replace(/[^\w:.-]/g, '');
    var objection = data.stance === 'objection';
    var topics  = objection && data.topics && data.topics.join ? clean(data.topics.join(' · '), 300) : '';
    var section = objection ? clean(data.section, 80) : '';
    var reason  = objection ? clean(data.reason, 1500) : '';

    if (!name || !street || !data.consent) return reply({ ok: false, error: 'missing' });
    if (!/\d/.test(house) && !loc) return reply({ ok: false, error: 'house' });
    if (objection && (!topics || reason.length < 10)) return reply({ ok: false, error: 'reason' });

    var key = /\d/.test(house) ? householdKey(street, house) : 'מבנה|' + loc;
    var row = [new Date(), name, street, house, key, objection ? 'בעד, עם השגה' : 'בעד',
               topics, section, reason, loc];

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sh = getSheet(SHEET_MAIN, HEADERS_MAIN);
      var last = sh.getLastRow();
      if (last > 1) {
        var keys = sh.getRange(2, COL_KEY, last - 1, 1).getValues();
        for (var i = 0; i < keys.length; i++) {
          if (String(keys[i][0]) === key) {
            getSheet(SHEET_DUP, HEADERS_DUP).appendRow(row.concat(['בית האב כבר חתום']));
            var dup = listPayload();
            dup.duplicate = true;
            return reply(dup);
          }
        }
      }
      sh.appendRow(row);
      CacheService.getScriptCache().remove(CACHE_KEY);
    } finally {
      lock.releaseLock();
    }
    var res = listPayload();
    res.signedNow = true;
    return reply(res);
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

/* רשימת בתי האב שחתמו — רק רחוב, מספר ומבנה במפה. בלי שמות ובלי השגות. */
function listPayload() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);
  var sh = getSheet(SHEET_MAIN, HEADERS_MAIN);
  var last = sh.getLastRow();
  var signed = [];
  if (last > 1) {
    var rows = sh.getRange(2, 1, last - 1, HEADERS_MAIN.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r[COL_STREET - 1] === '') continue;
      signed.push([String(r[COL_STREET - 1]), String(r[COL_STREET]), String(r[COL_LOC - 1] || '')]);
    }
  }
  var payload = { ok: true, count: signed.length, signed: signed, updated: new Date().toISOString() };
  cache.put(CACHE_KEY, JSON.stringify(payload), 20);
  return payload;
}

/* אותו כלל נרמול בדיוק כמו בדף (plan/index.html) */
function householdKey(street, house) {
  var s = String(street)
    .replace(/[״“”]/g, '"').replace(/[׳‘’`]/g, "'")
    .replace(/^\s*(רחוב|רח')\s*/, '')
    .replace(/-/g, ' ')
    .replace(/\s+/g, ' ').trim();
  var h = String(house)
    .replace(/[״“”"׳‘’`'.\-\s]/g, '')
    .trim();
  return s + '|' + h;
}

function clean(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
}

function getSheet(name, headers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function reply(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb && /^[A-Za-z_$][\w$]{0,40}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* להרצה ידנית פעם אחת מהעורך — מאשר הרשאות ויוצר את הלשוניות */
function setup() {
  var main = getSheet(SHEET_MAIN, HEADERS_MAIN);
  var dup  = getSheet(SHEET_DUP, HEADERS_DUP);
  main.getRange(1, 1, 1, HEADERS_MAIN.length).setValues([HEADERS_MAIN]);   // מעדכן כותרות אם הלשונית נוצרה בגרסה קודמת
  dup.getRange(1, 1, 1, HEADERS_DUP.length).setValues([HEADERS_DUP]);
  Logger.log(JSON.stringify(listPayload()));
}
