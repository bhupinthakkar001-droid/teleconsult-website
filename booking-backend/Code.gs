/** @OnlyCurrentDoc */

const TIME_ZONE = 'Asia/Kolkata';
const SLOTS_SHEET = 'Slots';
const REQUESTS_SHEET = 'Requests';
const REQUEST_HEADERS = [
  'requestId', 'slotId', 'date', 'time',
  'patientName', 'patientPhone', 'visitType', 'status',
  'requestedAt', 'decisionAt', 'videoLink'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teleconsult bookings')
    .addItem('Approve request & create Google Meet', 'approveSelectedRequest')
    .addItem('Decline selected request', 'declineSelectedRequest')
    .addSeparator()
    .addItem('Configure doctor WhatsApp', 'configureTeleconsultation')
    .addToUi();
}

function configureTeleconsultation() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty('DOCTOR_WHATSAPP') || '';
  const response = ui.prompt(
    'Doctor WhatsApp number',
    'Enter country code + number, digits only. Example: 919876543210' +
      (current ? '\nA number is already saved; enter a new one only if you want to change it.' : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const phone = String(response.getResponseText() || '').replace(/\D/g, '');
  if (!/^\d{10,15}$/.test(phone)) {
    ui.alert('That number does not look valid. Please use digits including country code.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('DOCTOR_WHATSAPP', phone);
  ui.alert('Doctor WhatsApp number saved privately.');
}

function setupBookingSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let slots = ss.getSheetByName(SLOTS_SHEET);
  if (!slots) slots = ss.insertSheet(SLOTS_SHEET);
  let requests = ss.getSheetByName(REQUESTS_SHEET);
  if (!requests) requests = ss.insertSheet(REQUESTS_SHEET);

  const slotHeaders = ['slotId', 'date', 'time', 'status', 'patientName', 'patientPhone', 'requestedAt', 'requestId'];
  if (slots.getLastRow() === 0) slots.appendRow(slotHeaders);
  if (requests.getLastRow() === 0) {
    requests.appendRow(REQUEST_HEADERS);
  } else {
    requests.getRange(1, 11).setValue('videoLink');
  }

  slots.setFrozenRows(1);
  requests.setFrozenRows(1);
  slots.getRange(1, 1, 1, slotHeaders.length).setFontWeight('bold');
  requests.getRange(1, 1, 1, REQUEST_HEADERS.length).setFontWeight('bold');
  slots.getRange('A:C').setNumberFormat('@');
  slots.getRange('E:F').setNumberFormat('@');
  requests.getRange('A:F').setNumberFormat('@');
  SpreadsheetApp.getUi().alert('Setup complete. Your existing booking rows have been preserved.');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (String(p.action || '') !== 'slots') {
    return jsonOutput({ ok: true, message: 'Teleconsult booking service is running.' });
  }
  const payload = JSON.stringify({ ok: true, slots: getPublicSlots_() });
  const callback = String(p.callback || '');
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + payload + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput({ ok: true, slots: getPublicSlots_() });
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (String(p.action || '') !== 'request') {
    return receiptPage_('Request not submitted', 'Please return to the website and try again.', '');
  }
  if (String(p.website || '').trim()) {
    return receiptPage_('Unable to submit', 'Please return to the website and try again.', '');
  }

  const name = cleanText_(p.patientName, 80);
  const phone = cleanPhone_(p.phone);
  const phoneDigits = phone.replace(/\D/g, '');
  const slotId = cleanText_(p.slotId, 80);
  const visitType = cleanText_(p.visitType, 40);
  const doctorWhatsApp = PropertiesService.getScriptProperties().getProperty('DOCTOR_WHATSAPP') || '';

  if (name.length < 2) return receiptPage_('Check your name', 'Please enter your full name and try again.', '');
  if (phoneDigits.length < 10 || phoneDigits.length > 15) return receiptPage_('Check your mobile number', 'Please enter a valid mobile number and try again.', '');
  if (!slotId) return receiptPage_('Choose an appointment time', 'Please return to the website and choose an available time.', '');
  if (!['First consultation', 'Follow-up consultation'].includes(visitType)) return receiptPage_('Choose appointment type', 'Please choose first consultation or follow-up.', '');
  if (!/^\d{10,15}$/.test(doctorWhatsApp)) return receiptPage_('Booking setup incomplete', 'The doctor is still setting up WhatsApp confirmation. Please try again later.', '');

  const lock = LockService.getScriptLock();
  let sheet = null;
  let rowIndex = 0;
  let reserved = false;
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getSheetByName(SLOTS_SHEET);
    const requests = ss.getSheetByName(REQUESTS_SHEET);
    if (!sheet || !requests || sheet.getLastRow() < 2) {
      return receiptPage_('No appointment slots available', 'Please check again later or contact the doctor.', '');
    }

    const rows = sheet.getDataRange().getDisplayValues();
    let chosen = null;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === slotId) {
        rowIndex = i + 1;
        chosen = rows[i];
        break;
      }
    }

    if (!chosen || String(chosen[3]).trim().toUpperCase() !== 'AVAILABLE' ||
        !/^\d{4}-\d{2}-\d{2}$/.test(String(chosen[1]).trim()) ||
        String(chosen[1]).trim() < todayString_()) {
      return receiptPage_('That slot is no longer available', 'Another patient may have requested this time. Please return to the website and choose another slot.', '');
    }

    const requestId = Utilities.getUuid().slice(0, 8).toUpperCase();
    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(rowIndex, 4, 1, 5).setValues([[
      'PENDING', safeCell_(name), safeCell_(phone), now, requestId
    ]]);
    reserved = true;

    requests.appendRow([
      requestId, slotId, String(chosen[1]), String(chosen[2]),
      safeCell_(name), safeCell_(phone), visitType, 'PENDING', now, '', ''
    ]);

    const msg = [
      'Hello Dr. Bhupin, I am requesting a video consultation.',
      'Name: ' + name,
      'Mobile: ' + phone,
      'Preferred date: ' + chosen[1],
      'Preferred time: ' + chosen[2],
      'Appointment type: ' + visitType,
      'Request ID: ' + requestId,
      '',
      'Please confirm or decline this request. I understand it is not confirmed until you approve it.'
    ].join('\n');
    const waUrl = 'https://wa.me/' + doctorWhatsApp + '?text=' + encodeURIComponent(msg);
    return receiptPage_('Request received',
      'Your request has been saved and this slot is now held as PENDING, so other patients will no longer see it. Tap below and press Send in WhatsApp. Your appointment is confirmed only after the doctor approves it.',
      waUrl);
  } catch (err) {
    if (reserved && sheet && rowIndex > 1) {
      try {
        sheet.getRange(rowIndex, 4, 1, 5).clearContent();
        sheet.getRange(rowIndex, 4).setValue('AVAILABLE');
      } catch (ignored) {}
    }
    return receiptPage_('Could not submit the request', 'No confirmation was made. Please return to the website and try again, or contact the doctor.', '');
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
    if (id && date >= today && /^\d{4}-\d{2}-\d{2}$/.test(date) && time && status === 'AVAILABLE') {
      result.push({ slotId: id, date: date, time: time });
    }
  }
  result.sort((a,b) => a.date === b.date ? toMinutes_(a.time) - toMinutes_(b.time) : a.date.localeCompare(b.date));
  return result;
}

function approveSelectedRequest() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName(REQUESTS_SHEET);
  if (!reqSheet || ss.getActiveSheet().getName() !== REQUESTS_SHEET) {
    ui.alert('Open the Requests tab and select a PENDING request row first.');
    return;
  }
  const row = ss.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Select a request row, not the header.');
    return;
  }
  const req = reqSheet.getRange(row, 1, 1, Math.max(REQUEST_HEADERS.length, reqSheet.getLastColumn())).getDisplayValues()[0];
  if (String(req[7]).toUpperCase() !== 'PENDING') {
    ui.alert('This request is not pending.');
    return;
  }

  let meetLink;
  let eventId;
  try {
    const created = createGoogleMeetEvent_(req);
    meetLink = created.meetLink;
    eventId = created.eventId;
  } catch (err) {
    ui.alert('Could not create Google Meet link:\n\n' + String(err.message || err) +
      '\n\nThe appointment remains PENDING. Check Calendar API setup and try again.');
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const latest = reqSheet.getRange(row, 1, 1, Math.max(REQUEST_HEADERS.length, reqSheet.getLastColumn())).getDisplayValues()[0];
    if (String(latest[7]).toUpperCase() !== 'PENDING') {
      ui.alert('This request is no longer pending. A Google Calendar event may have been created; please check your Calendar.');
      return;
    }
    const requestId = String(latest[0] || '');
    const slotId = String(latest[1] || '');
    const slots = ss.getSheetByName(SLOTS_SHEET);
    const slotRows = slots.getDataRange().getDisplayValues();
    let matchRow = -1;
    for (let i=1; i<slotRows.length; i++) {
      if (String(slotRows[i][0]) === slotId && String(slotRows[i][7]) === requestId) {
        matchRow = i + 1;
        break;
      }
    }
    if (matchRow < 0) {
      ui.alert('Could not find the matching held slot. The request remains PENDING; check the calendar for a newly created event.');
      return;
    }

    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    slots.getRange(matchRow, 4).setValue('CONFIRMED');
    reqSheet.getRange(row, 8).setValue('CONFIRMED');
    reqSheet.getRange(row, 10).setValue(now);
    reqSheet.getRange(row, 11).setValue(meetLink);

    const patientName = String(latest[4] || 'there');
    const patientPhoneRaw = String(latest[5] || '');
    const patientPhone = normalizeWhatsAppPhone_(patientPhoneRaw);
    const message = [
      'Hello ' + patientName + ', your video consultation with Dr. Bhupin is confirmed.',
      'Date: ' + String(latest[2] || ''),
      'Time: ' + String(latest[3] || ''),
      'Appointment: ' + String(latest[6] || 'Video consultation'),
      'Join your consultation using this Google Meet link: ' + meetLink,
      '',
      'Open the link at the appointment time. Allow camera and microphone access when prompted.',
      'If you need to reschedule, please contact the doctor.'
    ].join('\\n');
    const waUrl = 'https://wa.me/' + patientPhone + '?text=' + encodeURIComponent(message);
    const smsUrl = 'sms:+' + patientPhone + '?body=' + encodeURIComponent(message);
    const safeMessage = escapeHtml_(message);
    const safePhone = escapeHtml_(patientPhoneRaw);

    const dialogHtml = '<!doctype html><html><head><base target="_blank"><meta name="viewport" content="width=device-width,initial-scale=1"><style>' +
      'body{font:14px/1.5 Arial,sans-serif;padding:18px;color:#19362d}' +
      'h2{font-size:20px;margin:0 0 8px}p{color:#596b62;margin:0 0 14px}' +
      'a,.copy{display:block;text-align:center;background:#21634c;color:#fff;padding:12px 14px;border-radius:10px;text-decoration:none;font-weight:bold;margin-bottom:10px;border:0;cursor:pointer;width:100%;box-sizing:border-box;font-size:14px}' +
      '.secondary{background:#e5eee7;color:#19362d}.tertiary{background:#f4f7f4;color:#19362d;border:1px solid #dfe8e1}' +
      '.note{font-size:12px;margin-top:12px;color:#68776f}.phone{font-weight:bold}.message{width:100%;box-sizing:border-box;border:1px solid #dfe8e1;border-radius:8px;padding:10px;font:12px/1.45 Arial,sans-serif;color:#19362d;min-height:105px;margin:4px 0 10px}' +
      '</style></head><body><h2>Appointment confirmed</h2>' +
      '<p>The unique Google Meet link is saved in the Requests tab. The slot is now CONFIRMED and hidden from patients.</p>' +
      '<p>Patient mobile: <span class="phone">' + safePhone + '</span></p>' +
      '<a href="' + escapeHtml_(meetLink) + '">Join Google Meet as doctor →</a>' +
      '<a href="' + escapeHtml_(waUrl) + '">Send via WhatsApp →</a>' +
      '<a class="secondary" href="' + escapeHtml_(smsUrl) + '">Send via SMS / Messages →</a>' +
      '<p class="note">If the patient isn't on WhatsApp, use SMS / Messages instead. On some computers, the SMS button may not open a messaging app; copy the message below and send it from your phone.</p>' +
      '<label for="smsMessage"><strong>Confirmation message</strong></label>' +
      '<textarea id="smsMessage" class="message" readonly>' + safeMessage + '</textarea>' +
      '<button class="copy tertiary" onclick="var t=document.getElementById(\'smsMessage\');t.focus();t.select();try{document.execCommand(\'copy\');this.textContent=\'Message copied\';}catch(e){this.textContent=\'Select and copy the message above\';}">Copy confirmation message</button>' +
      '<p class="note">Review the recipient and message before sending. The message is not sent automatically. Send the private Meet link only to this patient.</p>' +
      '</body></html>';
    ui.showModalDialog(HtmlService.createHtmlOutput(dialogHtml).setWidth(440).setHeight(350), 'Video consultation ready');
  } finally {
    lock.releaseLock();
  }
}

function declineSelectedRequest() {
  setSelectedRequestStatus_('DECLINED');
}

function setSelectedRequestStatus_(newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName(REQUESTS_SHEET);
  const active = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  if (!reqSheet || active.getName() !== REQUESTS_SHEET) {
    ui.alert('Open the Requests tab and select a pending request row first.');
    return;
  }
  const row = active.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Select a request row, not the header.');
    return;
  }
  const req = reqSheet.getRange(row, 1, 1, Math.max(REQUEST_HEADERS.length, reqSheet.getLastColumn())).getDisplayValues()[0];
  if (String(req[7]).toUpperCase() !== 'PENDING') {
    ui.alert('This request is not pending.');
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const slots = ss.getSheetByName(SLOTS_SHEET);
    const slotRows = slots.getDataRange().getDisplayValues();
    let matchRow = -1;
    for (let i=1; i<slotRows.length; i++) {
      if (String(slotRows[i][0]) === String(req[1]) && String(slotRows[i][7]) === String(req[0])) {
        matchRow = i + 1;
        break;
      }
    }
    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    if (matchRow > 0) {
      slots.getRange(matchRow, 4, 1, 5).clearContent();
      slots.getRange(matchRow, 4).setValue('AVAILABLE');
    }
    reqSheet.getRange(row, 8).setValue(newStatus);
    reqSheet.getRange(row, 10).setValue(now);
    ui.alert('Request declined and its slot is available again. Please notify the patient on WhatsApp.');
  } finally {
    lock.releaseLock();
  }
}

function createGoogleMeetEvent_(req) {
  if (typeof Calendar === 'undefined' || !Calendar.Events) {
    throw new Error('Google Calendar service is not enabled yet. In Apps Script, click Services + and add Google Calendar API (v3), then run again.');
  }
  const date = String(req[2] || '').trim();
  const time = String(req[3] || '').trim();
  const start = parseAppointmentDateTime_(date, time);
  const now = Date.now();
  if (start.getTime() < now - 60 * 60 * 1000) {
    throw new Error('This slot is in the past. Please choose a future appointment.');
  }
  const end = new Date(start.getTime() + 30 * 60 * 1000);
  const requestId = String(req[0] || Utilities.getUuid()).replace(/[^a-zA-Z0-9]/g, '');
  const event = {
    summary: 'Private video consultation',
    description: 'Teleconsultation appointment. Booking reference: ' + requestId,
    start: { dateTime: Utilities.formatDate(start, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'") },
    end: { dateTime: Utilities.formatDate(end, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'") },
    conferenceData: {
      createRequest: {
        requestId: requestId + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8),
        conferenceSolutionKey: { type: 'hangoutsMeet' }
      }
    }
  };

  const inserted = Calendar.Events.insert(event, 'primary', {
    conferenceDataVersion: 1,
    sendUpdates: 'none'
  });
  if (!inserted || !inserted.id) throw new Error('Google Calendar did not create the appointment event.');

  let latest = inserted;
  for (let i=0; i<8; i++) {
    const points = latest.conferenceData && latest.conferenceData.entryPoints;
    const meet = points && points.find(p => p.entryPointType === 'video' && p.uri);
    const link = meet ? meet.uri : latest.hangoutLink;
    if (link && /^https:\/\/meet\.google\.com\//.test(link)) {
      return { meetLink: link, eventId: inserted.id };
    }
    Utilities.sleep(1000);
    latest = Calendar.Events.get('primary', inserted.id);
  }
  // Avoid leaving an unlinked event behind if Meet conference creation failed.
  try { Calendar.Events.remove('primary', inserted.id); } catch (ignored) {}
  throw new Error('Google Calendar created the event but did not return a Meet link. Check that Google Meet is available for your account.');
}

function parseAppointmentDateTime_(dateValue, timeValue) {
  const date = String(dateValue || '').trim();
  const time = String(timeValue || '').trim();
  const dm = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!dm) throw new Error('Invalid date in Requests. Expected YYYY-MM-DD: ' + date);
  let hour;
  let minute;
  let tm = time.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (tm) {
    hour = Number(tm[1]) % 12;
    minute = Number(tm[2]);
    if (tm[3].toUpperCase() === 'PM') hour += 12;
  } else {
    tm = time.match(/^(\d{1,2}):(\d{2})$/);
    if (!tm) throw new Error('Invalid time in Requests: ' + time);
    hour = Number(tm[1]);
    minute = Number(tm[2]);
  }
  const year = Number(dm[1]), month = Number(dm[2]), day = Number(dm[3]);
  if (hour > 23 || minute > 59) throw new Error('Invalid time in Requests: ' + time);
  const check = new Date(Date.UTC(year, month-1, day));
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month-1 || check.getUTCDate() !== day) {
    throw new Error('Invalid date in Requests: ' + date);
  }
  return new Date(Date.UTC(year, month-1, day, hour, minute, 0) - 5.5*60*60*1000);
}

function normalizeWhatsAppPhone_(value) {
  let digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 10) digits = '91' + digits;
  else if (digits.length === 11 && digits.charAt(0) === '0') digits = '91' + digits.slice(1);
  return digits;
}

function todayString_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
}

function toMinutes_(time) {
  const m = String(time).match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) {
    const h24 = String(time).match(/^(\d{1,2}):(\d{2})$/);
    return h24 ? Number(h24[1])*60 + Number(h24[2]) : 9999;
  }
  let hour = Number(m[1]) % 12;
  if (m[3].toUpperCase() === 'PM') hour += 12;
  return hour*60 + Number(m[2]);
}

function cleanText_(value, maxLength) {
  return String(value || '').replace(/[<>\u0000-\u001F]/g, '').trim().slice(0, maxLength);
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

/**
 * Run this once from the Apps Script editor to force Calendar permission
 * authorization and verify that the Calendar API can create an event.
 * A temporary test event is deleted immediately afterward.
 */
function authorizeCalendarAccess() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/calendar.events.owned'
  ]);

  const start = new Date(Date.now() + 2 * 60 * 1000);
  const end = new Date(start.getTime() + 60 * 1000);
  const testEvent = Calendar.Events.insert({
    summary: 'Teleconsult permission test - auto deleted',
    start: { dateTime: Utilities.formatDate(start, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'") },
    end: { dateTime: Utilities.formatDate(end, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'") }
  }, 'primary');

  if (testEvent && testEvent.id) {
    Calendar.Events.remove('primary', testEvent.id);
  }
  SpreadsheetApp.getUi().alert(
    'Calendar permission is working. The temporary test event was removed. Return to Requests and retry the pending booking.'
  );
}
