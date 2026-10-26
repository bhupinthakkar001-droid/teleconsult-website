/** @OnlyCurrentDoc */

const TIME_ZONE = 'Asia/Kolkata';
const SLOTS_SHEET = 'Slots';
const REQUESTS_SHEET = 'Requests';

// Replace this with your WhatsApp number including country code, digits only.
// Example format: 919876543210
const DOCTOR_WHATSAPP = '91XXXXXXXXXX';

const SLOT_HEADERS = [
  'slotId', 'date', 'time', 'status',
  'patientName', 'patientPhone', 'requestedAt', 'requestId'
];
const REQUEST_HEADERS = [
  'requestId', 'slotId', 'date', 'time',
  'patientName', 'patientPhone', 'visitType', 'status', 'requestedAt', 'decisionAt'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teleconsult bookings')
    .addItem('Approve selected request', 'approveSelectedRequest')
    .addItem('Decline selected request', 'declineSelectedRequest')
    .addToUi();
}

function setupBookingSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let slots = ss.getSheetByName(SLOTS_SHEET);
  if (!slots) slots = ss.insertSheet(SLOTS_SHEET);
  let requests = ss.getSheetByName(REQUESTS_SHEET);
  if (!requests) requests = ss.insertSheet(REQUESTS_SHEET);

  if (slots.getLastRow() === 0) slots.appendRow(SLOT_HEADERS);
  if (requests.getLastRow() === 0) requests.appendRow(REQUEST_HEADERS);

  slots.setFrozenRows(1);
  requests.setFrozenRows(1);
  slots.getRange(1, 1, 1, SLOT_HEADERS.length).setFontWeight('bold');
  requests.getRange(1, 1, 1, REQUEST_HEADERS.length).setFontWeight('bold');
  slots.getRange('A:C').setNumberFormat('@');
  slots.getRange('E:F').setNumberFormat('@');
  requests.getRange('A:F').setNumberFormat('@');

  SpreadsheetApp.getUi().alert(
    'Setup complete.\n\nIn the Slots sheet, add one row per appointment slot. ' +
    'Use date format YYYY-MM-DD, times such as 10:00 AM, and status AVAILABLE. ' +
    'Keep this spreadsheet private.'
  );
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action !== 'slots') {
    return jsonOutput({ ok: true, message: 'Teleconsult booking service is running.' });
  }

  const slots = getPublicSlots_();
  const payload = JSON.stringify({ ok: true, slots: slots });
  const callback = String(params.callback || '');
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + payload + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (String(p.action || '') !== 'request') {
    return receiptPage_('Request not submitted', 'This link is for appointment requests. Please return to the website and try again.', '');
  }

  // Honeypot field: real patients should leave it empty.
  if (String(p.website || '').trim()) {
    return receiptPage_('Unable to submit', 'Please return to the website and try again.', '');
  }

  const name = cleanText_(p.patientName, 80);
  const phone = cleanPhone_(p.phone);
  const slotId = cleanText_(p.slotId, 80);
  const visitType = cleanText_(p.visitType, 40);

  if (name.length < 2 || !/^\+?[0-9 ]{7,18}$/.test(phone) || !slotId || !['First consultation', 'Follow-up consultation'].includes(visitType)) {
    return receiptPage_('Check your details', 'Please return to the website and enter a valid name and mobile number.', '');
  }
  if (!DOCTOR_WHATSAPP || DOCTOR_WHATSAPP.indexOf('X') !== -1) {
    return receiptPage_('Booking setup incomplete', 'The booking service has not yet been fully configured. Please try again later.', '');
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SLOTS_SHEET);
    const requests = ss.getSheetByName(REQUESTS_SHEET);
    if (!sheet || !requests || sheet.getLastRow() < 2) {
      return receiptPage_('Bookings are not open yet', 'No appointment slots are currently available. Please check again later.', '');
    }

    const rows = sheet.getDataRange().getDisplayValues();
    let rowIndex = -1;
    let chosen = null;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === slotId) {
        rowIndex = i + 1;
        chosen = rows[i];
        break;
      }
    }

    if (!chosen || String(chosen[3]).toUpperCase() !== 'AVAILABLE' ||
        !/^\d{4}-\d{2}-\d{2}$/.test(String(chosen[1])) ||
        String(chosen[1]) < todayString_()) {
      return receiptPage_('That slot is no longer available', 'Someone may have requested this time first. Please return to the website and choose another available slot.', '');
    }

    const requestId = Utilities.getUuid().slice(0, 8).toUpperCase();
    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(rowIndex, 4, 1, 5).setValues([[
      'PENDING', safeCell_(name), safeCell_(phone), now, requestId
    ]]);
    requests.appendRow([
      requestId, slotId, String(chosen[1]), String(chosen[2]),
      safeCell_(name), safeCell_(phone), visitType, 'PENDING', now, ''
    ]);

    const msg = 'Hello Dr. Bhupin, I have requested an online consultation. ' +
      'Name: ' + name + '. Preferred slot: ' + chosen[1] + ' at ' + chosen[2] +
      '. Request ID: ' + requestId + '. Please confirm availability.';
    const waUrl = 'https://wa.me/' + DOCTOR_WHATSAPP + '?text=' + encodeURIComponent(msg);
    return receiptPage_(
      'Request received',
      'Your requested slot is held as PENDING while the doctor reviews it. It is not confirmed until the doctor replies to you on WhatsApp.',
      waUrl
    );
  } catch (err) {
    return receiptPage_('Could not submit the request', 'Please try again later. If the problem continues, contact the doctor directly.', '');
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function getPublicSlots_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SLOTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getDataRange().getDisplayValues();
  const today = todayString_();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const id = String(r[0] || '').trim();
    const date = String(r[1] || '').trim();
    const time = String(r[2] || '').trim();
    const status = String(r[3] || '').trim().toUpperCase();
    if (id && date >= today && /^\d{4}-\d{2}-\d{2}$/.test(date) &&
        time && status === 'AVAILABLE') {
      result.push({ slotId: id, date: date, time: time });
    }
  }
  result.sort(function(a, b) {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return toMinutes_(a.time) - toMinutes_(b.time);
  });
  return result;
}

function approveSelectedRequest() {
  setSelectedRequestStatus_('CONFIRMED');
}

function declineSelectedRequest() {
  setSelectedRequestStatus_('DECLINED');
}

function setSelectedRequestStatus_(newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName(REQUESTS_SHEET);
  const active = ss.getActiveSheet();
  if (!reqSheet || active.getName() !== REQUESTS_SHEET) {
    SpreadsheetApp.getUi().alert('Open the Requests sheet and select the pending request row first.');
    return;
  }
  const row = active.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a request row, not the header.');
    return;
  }
  const req = reqSheet.getRange(row, 1, 1, REQUEST_HEADERS.length).getDisplayValues()[0];
  if (String(req[7]).toUpperCase() !== 'PENDING') {
    SpreadsheetApp.getUi().alert('This request is not pending.');
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const slots = ss.getSheetByName(SLOTS_SHEET);
    const slotRows = slots.getDataRange().getDisplayValues();
    let matchRow = -1;
    for (let i = 1; i < slotRows.length; i++) {
      if (String(slotRows[i][0]) === String(req[1]) &&
          String(slotRows[i][7]) === String(req[0])) {
        matchRow = i + 1;
        break;
      }
    }
    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    if (matchRow > 0) {
      if (newStatus === 'CONFIRMED') {
        slots.getRange(matchRow, 4).setValue('CONFIRMED');
      } else {
        slots.getRange(matchRow, 4, 1, 5).clearContent();
        slots.getRange(matchRow, 4).setValue('AVAILABLE');
      }
    }
    reqSheet.getRange(row, 8).setValue(newStatus);
    reqSheet.getRange(row, 10).setValue(now);
    SpreadsheetApp.getUi().alert(
      newStatus === 'CONFIRMED'
        ? 'Request marked CONFIRMED. Please also confirm with the patient on WhatsApp.'
        : 'Request declined and the slot has been made available again. Please notify the patient on WhatsApp.'
    );
  } finally {
    lock.releaseLock();
  }
}

function todayString_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
}

function toMinutes_(time) {
  const m = String(time).match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return 9999;
  let hour = Number(m[1]) % 12;
  if (m[3].toUpperCase() === 'PM') hour += 12;
  return hour * 60 + Number(m[2]);
}

function cleanText_(value, maxLength) {
  return String(value || '').replace(/[<>\\u0000-\\u001F]/g, '').trim().slice(0, maxLength);
}

function cleanPhone_(value) {
  return String(value || '').replace(/[^0-9 +]/g, '').trim().slice(0, 18);
}

function safeCell_(value) {
  const text = String(value || '');
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function receiptPage_(title, message, waUrl) {
  const safeTitle = escapeHtml_(title);
  const safeMessage = escapeHtml_(message);
  const waButton = waUrl
    ? '<a class="btn" href="' + escapeHtml_(waUrl) + '" target="_blank" rel="noopener">Continue to WhatsApp →</a>'
    : '';
  const html = '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + safeTitle + '</title><style>' +
    'body{margin:0;background:#f4f7f4;color:#19362d;font:16px/1.55 system-ui,sans-serif;display:grid;place-items:center;min-height:100vh;padding:20px;box-sizing:border-box}' +
    '.card{width:min(440px,100%);box-sizing:border-box;background:white;border:1px solid #dfe8e1;border-radius:24px;padding:28px}' +
    '.mark{width:44px;height:44px;border-radius:14px;background:#e4efe7;color:#21634c;display:grid;place-items:center;font-size:24px}' +
    'h1{font-size:27px;line-height:1.12;letter-spacing:-.04em;margin:18px 0 10px}p{color:#63746a;font-size:14px;margin:0 0 22px}' +
    '.btn{display:block;text-align:center;background:#21634c;color:white;padding:14px 16px;border-radius:13px;text-decoration:none;font-weight:750}' +
    '</style></head><body><main class="card"><div class="mark">✓</div><h1>' + safeTitle + '</h1><p>' + safeMessage + '</p>' +
    waButton + '<p style="margin:18px 0 0;font-size:12px">The appointment is not final until the doctor confirms it.</p></main></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(title);
}

function escapeHtml_(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
