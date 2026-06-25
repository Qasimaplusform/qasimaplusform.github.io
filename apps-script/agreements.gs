/**
 * متابعة الاتفاقيات — قسيمة بلس
 * Backend (Google Apps Script) لداشبورد tracker.html
 *
 * طريقة الإعداد:
 *  1) افتح Google Sheets جديد (سيكون قاعدة البيانات).
 *  2) من القائمة: Extensions  ->  Apps Script
 *  3) امسح أي كود موجود والصق هذا الملف بالكامل.
 *  4) احفظ، ثم: Deploy  ->  New deployment  ->  اختر النوع "Web app"
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  5) انسخ رابط الـ Web app المنتهي بـ /exec
 *  6) افتح tracker.html  ->  الإعدادات (⚙️)  ->  الصق الرابط  ->  حفظ.
 *
 * ملاحظة أمان: الرمز السري SECRET أدناه يجب أن يطابق الموجود في tracker.html
 * (القيمة الافتراضية: qplus2026). غيّرهما معاً إذا رغبت.
 */

var SECRET = 'qplus2026';
var SHEET_NAME = 'Agreements';
var HEADERS = ['id','customer','agentCode','agentName','date','phone','note','signed','signedAt','updatedAt'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
  }
  return sh;
}

function readAll_() {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]), customer: r[1], agentCode: r[2], agentName: r[3],
      date: r[4] ? formatDate_(r[4]) : '', phone: r[5] ? String(r[5]) : '',
      note: r[6], signed: (r[7] === 'نعم' || r[7] === true || r[7] === 'true'),
      signedAt: r[8] ? formatDate_(r[8]) : '', updatedAt: r[9] ? formatDate_(r[9]) : ''
    });
  }
  return out;
}

function formatDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function upsert_(rec) {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var signedVal = (rec.signed === true || rec.signed === 'نعم' || rec.signed === 'true') ? 'نعم' : '';
  var row = [rec.id, rec.customer || '', rec.agentCode || '', rec.agentName || '',
            rec.date || '', rec.phone ? "'" + rec.phone : '', rec.note || '',
            signedVal, rec.signedAt || '', rec.updatedAt || ''];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(rec.id)) {
      sh.getRange(i + 1, 1, 1, HEADERS.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

function del_(id) {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(id)) sh.deleteRow(i + 1);
  }
}

function out_(obj, callback) {
  var txt = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + txt + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback;
  if (p.key !== SECRET) return out_({ ok: false, error: 'auth' }, cb);
  var action = p.action || 'list';
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (action === 'list') return out_({ ok: true, records: readAll_() }, cb);
    if (action === 'upsert') {
      upsert_({ id: p.id, customer: p.customer, agentCode: p.agentCode, agentName: p.agentName,
                date: p.date, phone: p.phone, note: p.note, signed: p.signed, signedAt: p.signedAt, updatedAt: p.updatedAt });
      return out_({ ok: true }, cb);
    }
    if (action === 'delete') { del_(p.id); return out_({ ok: true }, cb); }
    return out_({ ok: false, error: 'unknown action' }, cb);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, cb);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.key !== SECRET) return out_({ ok: false, error: 'auth' });
    lock.waitLock(30000);
    if (data.action === 'bulk' && data.records && data.records.length) {
      for (var i = 0; i < data.records.length; i++) upsert_(data.records[i]);
      return out_({ ok: true, count: data.records.length });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
