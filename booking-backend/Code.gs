/** @OnlyCurrentDoc */

const TIME_ZONE = 'Asia/Kolkata';
const SLOTS_SHEET = 'Slots';
const REQUESTS_SHEET = 'Requests';
const DAILY_API_BASE = 'https://api.daily.co/v1';
const VIDEO_PAGE_BASE = 'https://bhupinthakkar001-droid.github.io/teleconsult-website/video.html';

const SLOT_HEADERS = [
  'slotId', 'date', 'time', 'status',
  'patientName', 'patientPhone', 'requestedAt', 'requestId'
];
const REQUEST_HEADERS = [
  'requestId', 'slotId', 'date', 'time',
  'patientName', 'patientPhone', 'visitType', 'status',
  'requestedAt', 'decisionAt', 'patientVideoLink', 'doctorVideoLink'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teleconsult bookings')
    .addItem('Approve request & create video room', 'approveSelectedRequest')
    .addItem('Decline selected request', 'declineSelectedRequest')
    .addSeparator()
    .addItem('Configure WhatsApp + Daily', 'configureTeleconsultation')
    .addToUi();
}

/**
 * Run once after pasting this code. Enter secrets in the private prompt;
 * they are stored in Script Properties, never in public source code.
 */
function configureTeleconsultation() {
  const ui = SpreadsheetApp.getUi();
  const phoneResponse = ui.prompt(
    'Doctor WhatsApp number',
    'Enter your number with country code, digits only (example: 919876543210). It is stored privately.',
    ui.ButtonSet.OK_CANCEL
  );
  if (phoneResponse.getSelectedButton() !== ui.Button.OK) return;

  const doctorPhone = String(phoneResponse.getResponseText() || '').replace(/\D/g, '');
  if (!/^\d{10,15}$/.test(doctorPhone)) {
    ui.alert('That number does not look valid. Run Configure WhatsApp + Daily again and enter digits with country code.');
    return;
  }

  const keyResponse = ui.prompt(
    'Daily API key',
    'Paste your Daily API key from the Daily developer dashboard. It will be stored privately in Apps Script, not in GitHub.',
    ui.ButtonSet.OK_CANCEL
  );
  if (keyResponse.getSelectedButton() !== ui.Button.OK) return;

  const apiKey = String(keyResponse.getResponseText() || '').trim();
  if (apiKey.length < 12) {
    ui.alert('The API key looks too short. Run Configure WhatsApp + Daily again and paste the full key.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('DOCTOR_WHATSAPP', doctorPhone);
  props.setProperty('DAILY_API_KEY', apiKey);
  ui.alert('Configuration saved privately. Your WhatsApp number and Daily API key were not added to the source code.');
}

function setupBookingSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let slots = ss.getSheetByName(SLOTS_SHEET);
  if (!slots) slots = ss.insertSheet(SLOTS_SHEET);
  let requests = ss.getSheetByName(REQUESTS_SHEET);
  if (!requests) requests = ss.insertSheet(REQUESTS_SHEET);

  if (slots.getLastRow() === 0) slots.appendRow(SLOT_HEADERS);
  if (requests.getLastRow() === 0) {
    requests.appendRow(REQUEST_HEADERS);
  } else {
    if (!String(requests.getRange(1, 11).getDisplayValue()).trim() ||
        String(requests.getRange(1, 11).getDisplayValue()).trim() === 'videoLink') {
      requests.getRange(1, 11).setValue('patientVideoLink');
    }
    if (!String(requests.getRange(1, 12).getDisplayValue()).trim()) {
      requests.getRange(1, 12).setValue('doctorVideoLink');
    }
  }

  slots.setFrozenRows(1);
  requests.setFrozenRows(1);
  slots.getRange(1, 1, 1, SLOT_HEADERS.length).setFontWeight('bold');
  requests.getRange(1, 1, 1, REQUEST_HEADERS.length).setFontWeight('bold');
  slots.getRange('A:C').setNumberFormat('@');
  slots.getRange('E:F').setNumberFormat('@');
  requests.getRange('A:F').setNumberFormat('@');

  SpreadsheetApp.getUi().alert(
    'Setup complete. Your Slots and Requests data has been preserved. ' +
    'Requests now supports separate patient and doctor video links in columns K and L.'
  );
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

  if (name.length < 2) {
    return receiptPage_('Check your name', 'Please enter your full name and try again.', '');
  }
  if (phoneDigits.length < 10 || phoneDigits.length > 15) {
    return receiptPage_('Check your mobile number', 'Please enter a valid mobile number and try again.', '');
  }
  if (!slotId) {
    return receiptPage_('Choose an appointment time', 'Please return to the website and choose an available time.', '');
  }
  if (!['First consultation', 'Follow-up consultation'].includes(visitType)) {
    return receiptPage_('Choose appointment type', 'Please choose first consultation or follow-up.', '');
  }
  if (!/^\d{10,15}$/.test(doctorWhatsApp)) {
    return receiptPage_('Booking setup incomplete', 'The doctor is still setting up WhatsApp confirmation. Please try again later.', '');
  }

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
      return receiptPage_(
        'That slot is no longer available',
        'Another patient may have requested this time. Please return to the website and choose another slot.',
        ''
      );
    }

    const requestId = Utilities.getUuid().slice(0, 8).toUpperCase();
    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');

    sheet.getRange(rowIndex, 4, 1, 5).setValues([[
      'PENDING', safeCell_(name), safeCell_(phone), now, requestId
    ]]);
    reserved = true;

    requests.appendRow([
      requestId, slotId, String(chosen[1]), String(chosen[2]),
      safeCell_(name), safeCell_(phone), visitType, 'PENDING', now, '', '', ''
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
    return receiptPage_(
      'Request received',
      'Your request has been saved and this slot is now held as PENDING, so other patients will no longer see it. Tap below and press Send in WhatsApp. Your appointment is confirmed only after the doctor approves it.',
      waUrl
    );
  } catch (err) {
    if (reserved && sheet && rowIndex > 1) {
      try {
        sheet.getRange(rowIndex, 4, 1, 5).clearContent();
        sheet.getRange(rowIndex, 4).setValue('AVAILABLE');
      } catch (ignored) {}
    }
    return receiptPage_(
      'Could not submit the request',
      'No confirmation was made. Please return to the website and try again, or contact the doctor.',
      ''
    );
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

  let video;
  try {
    video = createDailySession_(req);
  } catch (err) {
    ui.alert('Could not create the video room', String(err.message || err) + '\n\nNo appointment status was changed. Check your Daily API key and try again.');
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const latest = reqSheet.getRange(row, 1, 1, REQUEST_HEADERS.length).getDisplayValues()[0];
    if (String(latest[7]).toUpperCase() !== 'PENDING') {
      ui.alert('This request is no longer pending. A temporary Daily room may have been created; it will expire automatically.');
      return;
    }

    const requestId = String(latest[0] || '');
    const slotId = String(latest[1] || '');
    const slots = ss.getSheetByName(SLOTS_SHEET);
    const slotRows = slots.getDataRange().getDisplayValues();
    let matchRow = -1;
    for (let i = 1; i < slotRows.length; i++) {
      if (String(slotRows[i][0]) === slotId &&
          String(slotRows[i][7]) === requestId) {
        matchRow = i + 1;
        break;
      }
    }
    if (matchRow < 0) {
      ui.alert('Could not find the matching held slot. No appointment status was changed.');
      return;
    }

    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    slots.getRange(matchRow, 4).setValue('CONFIRMED');
    reqSheet.getRange(row, 8).setValue('CONFIRMED');
    reqSheet.getRange(row, 10).setValue(now);
    reqSheet.getRange(row, 11).setValue(video.patientLink);
    reqSheet.getRange(row, 12).setValue(video.doctorLink);

    const patientName = String(latest[4] || 'there');
    const patientPhone = normalizeWhatsAppPhone_(latest[5]);
    const date = String(latest[2] || '');
    const time = String(latest[3] || '');
    const visitType = String(latest[6] || 'Video consultation');
    const message = [
      'Hello ' + patientName + ', your video consultation with Dr. Bhupin is confirmed.',
      'Date: ' + date,
      'Time: ' + time,
      'Appointment: ' + visitType,
      'Join your consultation using this secure link: ' + video.patientLink,
      '',
      'Open the link at your appointment time. Allow camera and microphone access when prompted.',
      'If you need to reschedule, please contact the doctor.'
    ].join('\n');
    const waUrl = 'https://wa.me/' + patientPhone + '?text=' + encodeURIComponent(message);

    const dialogHtml = '<!doctype html><html><head><base target="_blank"><style>' +
      'body{font:14px/1.5 Arial,sans-serif;padding:18px;color:#19362d}' +
      'h2{font-size:20px;margin:0 0 8px}p{color:#596b62;margin:0 0 14px}' +
      'a{display:block;text-align:center;background:#21634c;color:#fff;padding:12px 14px;border-radius:10px;text-decoration:none;font-weight:bold;margin-bottom:10px}' +
      '.secondary{background:#e5eee7;color:#19362d}.note{font-size:12px;margin-top:12px;color:#68776f}' +
      '</style></head><body><h2>Video consultation ready</h2>' +
      '<p>The appointment is confirmed. The room link has been saved in Requests columns K and L.</p>' +
      '<a href="' + escapeHtml_(video.doctorLink) + '">Join as doctor →</a>' +
      '<a class="secondary" href="' + escapeHtml_(waUrl) + '">Open WhatsApp to send patient link →</a>' +
      '<p class="note">WhatsApp opens with a prefilled message. Review it and press Send; it is not sent automatically. Do not forward the doctor link.</p>' +
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
    for (let i = 1; i < slotRows.length; i++) {
      if (String(slotRows[i][0]) === String(req[1]) &&
          String(slotRows[i][7]) === String(req[0])) {
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
    if (newStatus === 'DECLINED') {
      ui.alert('Request declined and its slot is available again. Please notify the patient on WhatsApp.');
    }
  } finally {
    lock.releaseLock();
  }
}

function createDailySession_(req) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('DAILY_API_KEY') || '';
  if (apiKey.length < 12) {
    throw new Error('Daily API key is not configured. Choose Teleconsult bookings → Configure WhatsApp + Daily and enter it once.');
  }

  const requestId = String(req[0] || Utilities.getUuid().slice(0, 8)).toLowerCase().replace(/[^a-z0-9-]/g, '');
  const patientName = cleanText_(req[4] || 'Patient', 80);
  const date = String(req[2] || '').trim();
  const time = String(req[3] || '').trim();
  const startDate = Utilities.parseDate(date + ' ' + time, TIME_ZONE, 'yyyy-MM-dd hh:mm a');
  const startEpoch = Math.floor(startDate.getTime() / 1000);
  const nowEpoch = Math.floor(Date.now() / 1000);
  if (!startEpoch || startEpoch < nowEpoch - 3600) {
    throw new Error('The selected appointment time appears to be in the past. Check the date/time in Requests before approving.');
  }

  const roomName = 'consult-' + requestId + '-' + Utilities.getUuid().slice(0, 4).toLowerCase();
  const nbf = Math.max(nowEpoch - 60, startEpoch - 15 * 60);
  const exp = Math.max(startEpoch + 2 * 60 * 60, nowEpoch + 30 * 60);

  const roomResponse = dailyApiRequest_('/rooms', {
    name: roomName,
    privacy: 'private',
    properties: {
      nbf: nbf,
      exp: exp,
      max_participants: 2,
      enable_prejoin_ui: true,
      enable_knocking: false,
      eject_at_room_exp: true,
      start_video_off: false,
      start_audio_off: false
    }
  }, apiKey);

  if (!roomResponse.url) {
    throw new Error('Daily did not return a room URL. Check your Daily account and API key.');
  }
  const roomUrl = String(roomResponse.url);
  const patientToken = createDailyToken_(apiKey, roomName, patientName, false, nbf, exp);
  const doctorToken = createDailyToken_(apiKey, roomName, 'Dr. Bhupin Thakkar', true, nbf, exp);

  return {
    patientLink: buildVideoPageLink_(roomUrl, patientToken, 'patient'),
    doctorLink: buildVideoPageLink_(roomUrl, doctorToken, 'doctor')
  };
}

function createDailyToken_(apiKey, roomName, userName, isOwner, nbf, exp) {
  const result = dailyApiRequest_('/meeting-tokens', {
    properties: {
      room_name: roomName,
      is_owner: isOwner,
      user_name: userName,
      nbf: nbf,
      exp: exp,
      eject_at_token_exp: true,
      enable_prejoin_ui: true
    }
  }, apiKey);
  if (!result.token) throw new Error('Daily did not return a meeting token.');
  return String(result.token);
}

function dailyApiRequest_(path, payload, apiKey) {
  const response = UrlFetchApp.fetch(DAILY_API_BASE + path, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const body = response.getContentText();
  let data = {};
  try { data = JSON.parse(body); } catch (ignored) {}
  if (status < 200 || status >= 300) {
    const errorText = String(data.info || data.error || data.message || 'Daily API request failed');
    throw new Error('Daily API returned ' + status + ': ' + errorText.slice(0, 220));
  }
  return data;
}

function buildVideoPageLink_(roomUrl, token, role) {
  const safeRoomUrl = String(roomUrl || '');
  if (!/^https:\/\/[a-z0-9-]+\.daily\.co\/[A-Za-z0-9_-]+(?:\?.*)?$/i.test(safeRoomUrl)) {
    throw new Error('Daily returned a room URL in an unexpected format.');
  }
  // Put bearer tokens in the URL fragment, which is not sent to the hosting server in HTTP requests.
  return VIDEO_PAGE_BASE + '#room=' + encodeURIComponent(safeRoomUrl) +
    '&token=' + encodeURIComponent(token) + '&role=' + encodeURIComponent(role);
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
  if (!m) return 9999;
  let hour = Number(m[1]) % 12;
  if (m[3].toUpperCase() === 'PM') hour += 12;
  return hour * 60 + Number(m[2]);
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
