const TIME_ZONE = 'Asia/Kolkata';
const CONSENT_VERSION = 'telemedicine-privacy-v1-2026-10-30';
const SLOTS_SHEET = 'Slots';
const REQUESTS_SHEET = 'Requests';
const REQUEST_HEADERS = [
  'requestId', 'slotId', 'date', 'time',
  'patientName', 'patientPhone', 'visitType', 'status',
  'requestedAt', 'decisionAt', 'videoLink', 'calendarEventId',
  'presentingComplaint', 'medicalHistory', 'onlineConsultConsent', 'consentAt',
  'serviceLevel', 'consultationFee', 'responseTarget',
  'dataUseConsent', 'dataConsentAt', 'consentVersion'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Teleconsult bookings')
    .addItem('Approve request & create Google Meet', 'approveSelectedRequest')
    .addItem('Decline selected request', 'declineSelectedRequest')
    .addItem('Set doctor dashboard password', 'setDoctorDashboardPassword')
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
  let changes = ss.getSheetByName('ChangeRequests');
  if (!changes) changes = ss.insertSheet('ChangeRequests');
  if (changes.getLastRow() === 0) changes.appendRow(['changeId','originalRequestId','patientPhone','patientName','changeType','newSlotId','newDate','newTime','status','requestedAt','decisionAt']);

  const slotHeaders = ['slotId', 'date', 'time', 'status', 'patientName', 'patientPhone', 'requestedAt', 'requestId'];
  if (slots.getLastRow() === 0) slots.appendRow(slotHeaders);
  if (requests.getLastRow() === 0) {
    requests.appendRow(REQUEST_HEADERS);
  } else {
    requests.getRange(1, 11).setValue('videoLink');
    requests.getRange(1, 12).setValue('calendarEventId');
    ensureIntakeColumns_(requests);
    ensureServiceLevelColumns_(requests);
    ensureConsentColumns_(requests);
  }

  slots.setFrozenRows(1);
  requests.setFrozenRows(1);
  slots.getRange(1, 1, 1, slotHeaders.length).setFontWeight('bold');
  requests.getRange(1, 1, 1, REQUEST_HEADERS.length).setFontWeight('bold');
  changes.setFrozenRows(1);
  changes.getRange(1, 1, 1, 11).setFontWeight('bold');
  slots.getRange('A:C').setNumberFormat('@');
  slots.getRange('E:F').setNumberFormat('@');
  requests.getRange('A:F').setNumberFormat('@');
  requests.getRange('M:V').setNumberFormat('@');
  SpreadsheetApp.getUi().alert('Setup complete. Your existing booking rows have been preserved.');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (String(p.action || '') === 'admin') return HtmlService.createHtmlOutputFromFile('Admin').setTitle('Doctor Admin Dashboard');
  if (String(p.action || '') === 'consumePrescriptionHandoff') {
    const result = consumePrescriptionHandoff_(String(p.token || ''));
    const payload = JSON.stringify(result);
    const callback = String(p.callback || '');
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService.createTextOutput(callback + '(' + payload + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonOutput(result);
  }
  if (String(p.action || '') === 'bookingOptions') {
    const payload=JSON.stringify({ok:true,priorityEnabled:getPriorityBookingsEnabled_()});
    const callback=String(p.callback||'');
    if(callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService.createTextOutput(callback+'('+payload+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonOutput({ok:true,priorityEnabled:getPriorityBookingsEnabled_()});
  }
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
  if (String(p.action || '') === 'changeRequest') return handlePatientChangeRequest_(p);
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
  const presentingComplaint = cleanMultilineText_(p.presentingComplaint, 1000);
  const medicalHistory = cleanMultilineText_(p.medicalHistory, 1000);
  const telemedicineConsentGiven = String(p.onlineConsultConsent || '') === 'yes';
  const dataUseConsentGiven = String(p.dataUseConsent || '') === 'yes';
  const serviceLevel = String(p.serviceLevel || 'REGULAR').toUpperCase();
  const doctorWhatsApp = PropertiesService.getScriptProperties().getProperty('DOCTOR_WHATSAPP') || '';

  if (name.length < 2) return receiptPage_('Check your name', 'Please enter your full name and try again.', '');
  if (phoneDigits.length < 10 || phoneDigits.length > 15) return receiptPage_('Check your mobile number', 'Please enter a valid mobile number and try again.', '');
  if (!slotId) return receiptPage_('Choose an appointment time', 'Please return to the website and choose an available time.', '');
  if (!['First consultation', 'Follow-up consultation'].includes(visitType)) return receiptPage_('Choose appointment type', 'Please choose first consultation or follow-up.', '');
  if (presentingComplaint.length < 3) return receiptPage_('Add your presenting problem', 'Please briefly describe what you are consulting about and how long you have had it.', '');
  if (!telemedicineConsentGiven) return receiptPage_('Telemedicine consent required', 'Please confirm that you understand the limitations of online consultation and wish to proceed.', '');
  if (!dataUseConsentGiven) return receiptPage_('Data-use consent required', 'Please confirm that you agree to the use of your booking and health details for consultation and medical records. See the Privacy Notice.', '');
  if (!['REGULAR','PRIORITY'].includes(serviceLevel)) return receiptPage_('Choose consultation option', 'Please choose regular or priority consultation.', '');
  if (serviceLevel === 'PRIORITY' && !getPriorityBookingsEnabled_()) {
    return receiptPage_('Priority response option unavailable', 'Priority response bookings are temporarily unavailable. Please choose a regular consultation.', '');
  }
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
    if (requests) { ensureIntakeColumns_(requests); ensureServiceLevelColumns_(requests); ensureConsentColumns_(requests); }
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
      safeCell_(name), safeCell_(phone), visitType, 'PENDING', now, '', '', '',
      safeCell_(presentingComplaint), safeCell_(medicalHistory), 'YES', now,
      serviceLevel, serviceLevel === 'PRIORITY' ? '2000' : '500',
      serviceLevel === 'PRIORITY' ? '30-minute response' : 'Standard scheduling',
      'YES', now, CONSENT_VERSION
    ]);

    const msg = [
      'Hello Dr. Bhupin, I am requesting a video consultation.',
      'Name: ' + name,
      'Mobile: ' + phone,
      'Preferred date: ' + chosen[1],
      'Preferred time: ' + chosen[2],
      'Appointment type: ' + visitType,
      'Consultation option: ' + (serviceLevel === 'PRIORITY' ? 'PRIORITY RESPONSE' : 'REGULAR'),
      'Consultation fee: ₹' + (serviceLevel === 'PRIORITY' ? '2000' : '500'),
      'Expected response: ' + (serviceLevel === 'PRIORITY' ? 'within 30 minutes' : 'standard scheduling'),
      'Request ID: ' + requestId,
      '',
      'Please confirm or decline this request. I understand it is not confirmed until you approve it.'
    ].join('\n');
    const waUrl = 'https://wa.me/' + doctorWhatsApp + '?text=' + encodeURIComponent(msg);
    const receiptMessage = serviceLevel === 'PRIORITY'
      ? 'Your priority response request has been received. We will respond within 30 minutes. The fee is ₹2,000. Your selected slot is held as PENDING; the consultation is confirmed only after the doctor accepts it. This is not an emergency service.'
      : 'Your regular request has been received and your slot is held as PENDING. The fee is ₹500. Your appointment is not confirmed until the doctor approves it. Tap below and press Send in WhatsApp.';
    return receiptPage_('Request received', receiptMessage, waUrl, requestId);
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
    reqSheet.getRange(row, 12).setValue(eventId);

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
      '<button class="copy secondary" onclick="var t=document.getElementById(\'smsMessage\');t.focus();t.select();try{document.execCommand(\'copy\');this.textContent=\'Message copied — send it from your phone\';}catch(e){this.textContent=\'Select and copy the message below\';}">Copy message for SMS / Messages</button>' +
      '<p class="note">SMS cannot be sent automatically from this desktop dialog. Copy the message, open the Messages app on your phone, address it to the patient mobile shown above, paste the text, and send it.</p>' +
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

function cleanMultilineText_(value, maxLength) {
  return String(value || '').replace(/[<>]/g, '').replace(/\r\n?/g, '\n')
    .replace(/[\u0000-\u0009\u000B-\u001F]/g, '').trim().slice(0, maxLength);
}

function ensureIntakeColumns_(sheet) {
  sheet.getRange(1, 13, 1, 4).setValues([[
    'presentingComplaint', 'medicalHistory', 'onlineConsultConsent', 'consentAt'
  ]]);
  sheet.getRange(1, 13, 1, 4).setFontWeight('bold');
  sheet.getRange('M:P').setNumberFormat('@');
}

function ensureConsentColumns_(sheet) {
  sheet.getRange(1, 20, 1, 3).setValues([['dataUseConsent', 'dataConsentAt', 'consentVersion']]);
  sheet.getRange(1, 20, 1, 3).setFontWeight('bold');
  sheet.getRange('T:V').setNumberFormat('@');
}

function ensureServiceLevelColumns_(sheet) {
  sheet.getRange(1, 17, 1, 3).setValues([['serviceLevel', 'consultationFee', 'responseTarget']]);
  sheet.getRange(1, 17, 1, 3).setFontWeight('bold');
  sheet.getRange('Q:S').setNumberFormat('@');
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

function receiptPage_(title, message, waUrl, reference) {
  const safeTitle = escapeHtml_(title);
  const safeMessage = escapeHtml_(message);
  const safeReference = reference ? '<p style="font-size:13px;background:#f4f7f4;border-radius:10px;padding:10px"><strong>Booking reference:</strong> ' + escapeHtml_(reference) + '<br>Please save this ID to manage your appointment later.</p>' : '';
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
    safeReference + waButton + '<p style="margin:18px 0 0;font-size:12px">The appointment is not final until the doctor confirms it.</p></main></body></html>';
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


/* ---------------- MOBILE DOCTOR DASHBOARD ----------------
   Set a password once from the spreadsheet menu. The dashboard
   uses short-lived server-side sessions; no sheet access is exposed
   directly to the browser.
*/
function setDoctorDashboardPassword() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Set doctor dashboard password',
    'Choose a strong password (at least 12 characters). You will use it to sign in to the private mobile dashboard.',
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const password = String(response.getResponseText() || '');
  if (password.length < 12) {
    ui.alert('Please use at least 12 characters.');
    return;
  }
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  PropertiesService.getScriptProperties().setProperty('DOCTOR_ADMIN_PASSWORD_HASH',
    Utilities.base64Encode(digest));
  ui.alert('Dashboard password saved. Open your website URL with ?action=admin to sign in.');
}

function adminLogin(password) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('DOCTOR_ADMIN_PASSWORD_HASH');
  if (!expected) throw new Error('Doctor dashboard password is not set. Set it from the Teleconsult bookings menu in the spreadsheet first.');
  const key = 'ADMIN_LOGIN_FAILS';
  const cache = CacheService.getScriptCache();
  const failures = Number(cache.get(key) || 0);
  if (failures >= 8) throw new Error('Too many attempts. Wait 15 minutes and try again.');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(password || ''), Utilities.Charset.UTF_8);
  const actual = Utilities.base64Encode(digest);
  if (actual !== expected) {
    cache.put(key, String(failures + 1), 900);
    throw new Error('Incorrect password.');
  }
  cache.remove(key);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('ADMIN_SESSION_' + token, '1', 21600);
  return {token:token};
}

function requireAdminSession_(token) {
  if (!token || CacheService.getScriptCache().get('ADMIN_SESSION_' + String(token)) !== '1') {
    throw new Error('Your session expired. Please sign in again.');
  }
}

function adminGetRequests(token) {
  requireAdminSession_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REQUESTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {pending:[], confirmed:[]};
  ensureIntakeColumns_(sheet);
  ensureServiceLevelColumns_(sheet);
  ensureConsentColumns_(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow()-1, REQUEST_HEADERS.length).getDisplayValues();
  const pending = [], confirmed = [];
  rows.forEach((r, i) => {
    const item = {row:i+2, requestId:r[0], date:r[2], time:r[3], patientName:r[4],
      patientPhone:r[5], visitType:r[6], status:r[7], requestedAt:r[8], videoLink:r[10],
      presentingComplaint:r[12], medicalHistory:r[13], onlineConsultConsent:r[14], consentAt:r[15],
      serviceLevel:r[16] || 'REGULAR', consultationFee:r[17] || '500', responseTarget:r[18] || 'Standard scheduling',
      dataUseConsent:r[19], dataConsentAt:r[20], consentVersion:r[21]};
    if (r[7] === 'PENDING') pending.push(item);
    if (r[7] === 'CONFIRMED') confirmed.push(item);
  });
  pending.sort((a,b) => {
    const priorityOrder=(a.serviceLevel==='PRIORITY'?0:1)-(b.serviceLevel==='PRIORITY'?0:1);
    return priorityOrder || (a.date+a.time).localeCompare(b.date+b.time);
  });
  confirmed.sort((a,b) => (b.date+b.time).localeCompare(a.date+a.time));
  return {pending:pending, confirmed:confirmed};
}

function adminApproveRequest(token, requestId) {
  requireAdminSession_(token);
  return processAdminRequest_(requestId, true);
}

function adminDeclineRequest(token, requestId) {
  requireAdminSession_(token);
  return processAdminRequest_(requestId, false);
}

function processAdminRequest_(requestId, approve) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reqSheet = ss.getSheetByName(REQUESTS_SHEET);
    const slots = ss.getSheetByName(SLOTS_SHEET);
    if (!reqSheet || !slots) throw new Error('Booking sheets are missing.');
    const values = reqSheet.getDataRange().getDisplayValues();
    let row = -1, req = null;
    for (let i=1; i<values.length; i++) {
      if (String(values[i][0]) === String(requestId)) { row=i+1; req=values[i]; break; }
    }
    if (!req || String(req[7]).toUpperCase() !== 'PENDING') throw new Error('This request is no longer pending.');
    const slotRows = slots.getDataRange().getDisplayValues();
    let slotRow = -1;
    for (let i=1; i<slotRows.length; i++) {
      if (String(slotRows[i][0]) === String(req[1]) && String(slotRows[i][7]) === String(req[0])) { slotRow=i+1; break; }
    }
    if (slotRow < 0) throw new Error('Could not find the matching held slot.');
    const now = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    if (approve) {
      const created = createGoogleMeetEvent_(req);
      slots.getRange(slotRow,4).setValue('CONFIRMED');
      reqSheet.getRange(row,8).setValue('CONFIRMED');
      reqSheet.getRange(row,10).setValue(now);
      reqSheet.getRange(row,11).setValue(created.meetLink);
      reqSheet.getRange(row,12).setValue(created.eventId);
      return {ok:true, meetLink:created.meetLink, patientPhone:normalizeWhatsAppPhone_(req[5]),
        patientName:req[4], date:req[2], time:req[3]};
    }
    slots.getRange(slotRow,4,1,5).clearContent();
    slots.getRange(slotRow,4).setValue('AVAILABLE');
    reqSheet.getRange(row,8).setValue('DECLINED');
    reqSheet.getRange(row,10).setValue(now);
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}


/* ---------------- MOBILE SLOT MANAGEMENT ----------------
   Only AVAILABLE slots can be edited or removed. PENDING and
   CONFIRMED slots are locked to protect patient appointments.
*/
function adminGetSlots(token) {
  requireAdminSession_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SLOTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {slots:[]};
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getDisplayValues();
  const today = todayString_();
  const result = [];
  rows.forEach(function(r) {
    const date = String(r[1] || '').trim();
    const status = String(r[3] || '').trim().toUpperCase();
    if (r[0] && date >= today && ['AVAILABLE','PENDING','CONFIRMED','CHANGE_PENDING'].includes(status)) {
      result.push({slotId:String(r[0]), date:date, time:String(r[2] || ''), status:status});
    }
  });
  result.sort(function(a,b) {
    return a.date === b.date ? toMinutes_(a.time) - toMinutes_(b.time) : a.date.localeCompare(b.date);
  });
  return {slots:result};
}

function adminAddSlot(token, dateValue, timeValue) {
  requireAdminSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const input = validateAdminSlotInput_(dateValue, timeValue);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SLOTS_SHEET);
    if (!sheet) throw new Error('Slots sheet was not found.');
    ensureNoDuplicateSlot_(sheet, input.date, input.minutes, '');
    const slotId = 'S' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
    sheet.appendRow([slotId, input.date, input.time, 'AVAILABLE', '', '', '', '']);
    const row = sheet.getLastRow();
    sheet.getRange(row, 1, 1, 3).setNumberFormat('@');
    return {ok:true, message:'Appointment slot added.'};
  } finally {
    lock.releaseLock();
  }
}

function adminUpdateSlot(token, slotId, dateValue, timeValue) {
  requireAdminSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const input = validateAdminSlotInput_(dateValue, timeValue);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SLOTS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('No appointment slots were found.');
    const rows = sheet.getDataRange().getDisplayValues();
    let row = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(slotId)) { row = i + 1; break; }
    }
    if (row < 0) throw new Error('This slot could not be found. Refresh the dashboard.');
    const status = String(sheet.getRange(row, 4).getDisplayValue()).toUpperCase();
    if (status !== 'AVAILABLE') throw new Error('Only available slots can be edited. Booked slots are locked.');
    ensureNoDuplicateSlot_(sheet, input.date, input.minutes, String(slotId));
    sheet.getRange(row, 2, 1, 2).setNumberFormat('@');
    sheet.getRange(row, 2, 1, 2).setValues([[input.date, input.time]]);
    return {ok:true, message:'Appointment slot updated.'};
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteSlot(token, slotId) {
  requireAdminSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SLOTS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('No appointment slots were found.');
    const rows = sheet.getDataRange().getDisplayValues();
    let row = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(slotId)) { row = i + 1; break; }
    }
    if (row < 0) throw new Error('This slot could not be found. Refresh the dashboard.');
    const status = String(sheet.getRange(row, 4).getDisplayValue()).toUpperCase();
    if (status !== 'AVAILABLE') throw new Error('Only available slots can be removed. Booked slots are locked.');
    // Retain the row for audit/history; public availability ignores REMOVED slots.
    sheet.getRange(row, 4).setValue('REMOVED');
    return {ok:true, message:'Appointment slot removed.'};
  } finally {
    lock.releaseLock();
  }
}

function validateAdminSlotInput_(dateValue, timeValue) {
  const date = String(dateValue || '').trim();
  const rawTime = String(timeValue || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Please choose a valid appointment date.');
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(rawTime)) throw new Error('Please choose a valid appointment time.');
  const start = parseAppointmentDateTime_(date, rawTime);
  if (start.getTime() <= Date.now()) throw new Error('Please choose a future date and time.');
  const parts = rawTime.split(':');
  const hour = Number(parts[0]);
  const minute = parts[1];
  const suffix = hour >= 12 ? 'PM' : 'AM';
  const hour12 = hour % 12 || 12;
  return {date:date, time:hour12 + ':' + minute + ' ' + suffix, minutes:hour * 60 + Number(minute)};
}

function ensureNoDuplicateSlot_(sheet, date, minutes, exceptSlotId) {
  if (sheet.getLastRow() < 2) return;
  const rows = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '');
    const rowDate = String(rows[i][1] || '').trim();
    const status = String(rows[i][3] || '').trim().toUpperCase();
    if (id && id !== exceptSlotId && rowDate === date &&
        ['AVAILABLE','PENDING','CONFIRMED','CHANGE_PENDING'].includes(status) &&
        toMinutes_(rows[i][2]) === minutes) {
      throw new Error('A slot already exists at this date and time.');
    }
  }
}


/* ---------------- PATIENT CANCELLATION / RESCHEDULING ---------------- */
function handlePatientChangeRequest_(p) {
  if (String(p.website || '').trim()) return receiptPage_('Unable to submit', 'Please return to the website and try again.', '');
  const requestId = cleanText_(p.requestId, 40).toUpperCase();
  const phone = cleanPhone_(p.phone);
  const phoneDigits = phone.replace(/\D/g, '');
  const changeType = String(p.changeType || '').toUpperCase();
  const newSlotId = cleanText_(p.newSlotId, 80);
  if (!requestId || phoneDigits.length < 10 || phoneDigits.length > 15) {
    return receiptPage_('Check your details', 'Enter the booking reference and the mobile number used for the booking.', '');
  }
  if (!['CANCEL','RESCHEDULE'].includes(changeType)) return receiptPage_('Invalid request', 'Choose cancellation or rescheduling.', '');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requests = ss.getSheetByName(REQUESTS_SHEET);
    if (!requests || requests.getLastRow() < 2) throw new Error('No booking was found with those details.');
    const rows = requests.getDataRange().getDisplayValues();
    let req = null, reqRow = -1;
    for (let i=1;i<rows.length;i++) {
      if (String(rows[i][0]).toUpperCase() === requestId &&
          normalizeDigits_(rows[i][5]) === normalizeDigits_(phone)) {
        req=rows[i]; reqRow=i+1; break;
      }
    }
    if (!req || String(req[7]).toUpperCase() !== 'CONFIRMED') {
      throw new Error('We could not verify a confirmed appointment with those details. Check your booking ID and mobile number.');
    }
    const changes = getChangeRequestsSheet_();
    const changeRows = changes.getDataRange().getDisplayValues();
    for (let i=1;i<changeRows.length;i++) {
      if (String(changeRows[i][1])===requestId && String(changeRows[i][8]).toUpperCase()==='PENDING') {
        throw new Error('A change request is already awaiting the doctor’s decision.');
      }
    }

    let target = null, targetRow = -1, targetDate = '', targetTime = '';
    if (changeType === 'RESCHEDULE') {
      if (!newSlotId) throw new Error('Please choose a new appointment time.');
      const slots = ss.getSheetByName(SLOTS_SHEET);
      const slotRows = slots.getDataRange().getDisplayValues();
      for (let i=1;i<slotRows.length;i++) {
        if (String(slotRows[i][0])===newSlotId && String(slotRows[i][3]).toUpperCase()==='AVAILABLE') {
          target=slotRows[i]; targetRow=i+1; break;
        }
      }
      if (!target) throw new Error('That new time is no longer available. Please refresh and choose another.');
      targetDate=String(target[1]); targetTime=String(target[2]);
      const slotStart=parseAppointmentDateTime_(targetDate,targetTime);
      if (slotStart.getTime()<=Date.now()) throw new Error('Please choose a future appointment time.');
    }
    const changeId='C'+Utilities.getUuid().replace(/-/g,'').slice(0,10).toUpperCase();
    const now=Utilities.formatDate(new Date(),TIME_ZONE,'yyyy-MM-dd HH:mm:ss');
    changes.appendRow([changeId,requestId,phoneDigits,req[4],changeType,newSlotId,targetDate,targetTime,'PENDING',now,'']);
    if (changeType === 'RESCHEDULE') {
      const slots=ss.getSheetByName(SLOTS_SHEET);
      slots.getRange(targetRow,4,1,5).setValues([['CHANGE_PENDING',safeCell_(req[4]),safeCell_(phone),now,requestId]]);
    }
    return receiptPage_('Change request received',
      changeType==='CANCEL'
        ? 'Your cancellation request has been sent to the doctor. Your appointment remains confirmed until the doctor approves the cancellation.'
        : 'Your preferred new time has been held temporarily. Your existing appointment remains confirmed until the doctor approves the change.',
      '');
  } catch(err) {
    return receiptPage_('Could not submit request', String(err.message||err), '');
  } finally {
    lock.releaseLock();
  }
}

function getChangeRequestsSheet_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('ChangeRequests');
  if(!sheet) sheet=ss.insertSheet('ChangeRequests');
  if(sheet.getLastRow()===0) {
    sheet.appendRow(['changeId','originalRequestId','patientPhone','patientName','changeType','newSlotId','newDate','newTime','status','requestedAt','decisionAt']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,11).setFontWeight('bold');
  }
  return sheet;
}

function adminGetChangeRequests(token) {
  requireAdminSession_(token);
  const sheet=getChangeRequestsSheet_();
  if(sheet.getLastRow()<2) return {pending:[]};
  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,11).getDisplayValues();
  const pending=[];
  rows.forEach(function(r) {
    if(String(r[8]).toUpperCase()==='PENDING') pending.push({
      changeId:r[0],requestId:r[1],patientPhone:r[2],patientName:r[3],
      changeType:r[4],newDate:r[6],newTime:r[7],currentDate:findRequestDate_(r[1]),currentTime:findRequestTime_(r[1]),status:r[8]
    });
  });
  return {pending:pending};
}

function findRequestDate_(requestId) {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET);
  if(!sheet || sheet.getLastRow()<2) return '';
  const rows=sheet.getDataRange().getDisplayValues();
  for(let i=1;i<rows.length;i++) if(String(rows[i][0])===String(requestId)) return String(rows[i][2]||'');
  return '';
}
function findRequestTime_(requestId) {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET);
  if(!sheet || sheet.getLastRow()<2) return '';
  const rows=sheet.getDataRange().getDisplayValues();
  for(let i=1;i<rows.length;i++) if(String(rows[i][0])===String(requestId)) return String(rows[i][3]||'');
  return '';
}
function adminDecideChangeRequest(token, changeId, approve) {
  requireAdminSession_(token);
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const changes=getChangeRequestsSheet_();
    const cr=changes.getDataRange().getDisplayValues();
    let change=null,changeRow=-1;
    for(let i=1;i<cr.length;i++) if(String(cr[i][0])===String(changeId)){change=cr[i];changeRow=i+1;break;}
    if(!change || String(change[8]).toUpperCase()!=='PENDING') throw new Error('This change request is no longer pending.');
    const requests=ss.getSheetByName(REQUESTS_SHEET);
    const rr=requests.getDataRange().getDisplayValues();
    let req=null,reqRow=-1;
    for(let i=1;i<rr.length;i++) if(String(rr[i][0])===String(change[1])){req=rr[i];reqRow=i+1;break;}
    if(!req || String(req[7]).toUpperCase()!=='CONFIRMED') throw new Error('The original appointment is no longer confirmed.');
    const slots=ss.getSheetByName(SLOTS_SHEET);
    const slotRows=slots.getDataRange().getDisplayValues();
    const now=Utilities.formatDate(new Date(),TIME_ZONE,'yyyy-MM-dd HH:mm:ss');

    if(!approve) {
      if(change[5]) releaseChangeHeldSlot_(slots,change[5],change[1]);
      changes.getRange(changeRow,9).setValue('DECLINED');
      changes.getRange(changeRow,11).setValue(now);
      return {ok:true,message:'Change request declined. The original appointment remains confirmed.'};
    }

    if(change[4]==='CANCEL') {
      releaseOriginalConfirmedSlot_(slots,req[1],req[0]);
      reqsSetCancellation_(requests,reqRow,req);
      changes.getRange(changeRow,9).setValue('APPROVED');
      changes.getRange(changeRow,11).setValue(now);
      if(req[11] && typeof Calendar!=='undefined' && Calendar.Events) {
        try { Calendar.Events.remove('primary',String(req[11])); } catch(ignored) {}
      }
      return {ok:true,message:'Appointment cancelled and its slot released.'};
    }

    if(change[4]==='RESCHEDULE') {
      let targetRow=-1,target=null;
      for(let i=1;i<slotRows.length;i++) {
        if(String(slotRows[i][0])===String(change[5]) &&
           String(slotRows[i][3]).toUpperCase()==='CHANGE_PENDING' &&
           String(slotRows[i][7])===String(change[1])) {targetRow=i+1;target=slotRows[i];break;}
      }
      if(!target) throw new Error('The requested new slot is no longer held. Decline this request and ask the patient to try again.');
      const revised=req.slice();
      revised[2]=String(target[1]); revised[3]=String(target[2]);
      const created=createGoogleMeetEvent_(revised);
      releaseOriginalConfirmedSlot_(slots,req[1],req[0]);
      slots.getRange(targetRow,4,1,5).setValues([['CONFIRMED',safeCell_(req[4]),safeCell_(req[5]),now,req[0]]]);
      requests.getRange(reqRow,2).setValue(String(target[0]));
      requests.getRange(reqRow,3).setValue(String(target[1]));
      requests.getRange(reqRow,4).setValue(String(target[2]));
      requests.getRange(reqRow,10).setValue(now);
      requests.getRange(reqRow,11).setValue(created.meetLink);
      requests.getRange(reqRow,12).setValue(created.eventId);
      changes.getRange(changeRow,9).setValue('APPROVED');
      changes.getRange(changeRow,11).setValue(now);
      if(req[11] && typeof Calendar!=='undefined' && Calendar.Events) {
        try { Calendar.Events.remove('primary',String(req[11])); } catch(ignored) {}
      }
      return {ok:true,message:'Appointment rescheduled. A new Google Meet link has been created.'};
    }
    throw new Error('Unsupported change request.');
  } finally {
    lock.releaseLock();
  }
}

function releaseChangeHeldSlot_(slots,slotId,requestId) {
  const rows=slots.getDataRange().getDisplayValues();
  for(let i=1;i<rows.length;i++) {
    if(String(rows[i][0])===String(slotId) &&
       String(rows[i][3]).toUpperCase()==='CHANGE_PENDING' &&
       String(rows[i][7])===String(requestId)) {
      slots.getRange(i+1,4,1,5).clearContent();
      slots.getRange(i+1,4).setValue('AVAILABLE');
      return;
    }
  }
}
function releaseOriginalConfirmedSlot_(slots,slotId,requestId) {
  const rows=slots.getDataRange().getDisplayValues();
  for(let i=1;i<rows.length;i++) {
    if(String(rows[i][0])===String(slotId) &&
       String(rows[i][3]).toUpperCase()==='CONFIRMED' &&
       String(rows[i][7])===String(requestId)) {
      slots.getRange(i+1,4,1,5).clearContent();
      slots.getRange(i+1,4).setValue('AVAILABLE');
      return;
    }
  }
}
function reqsSetCancellation_(requests,row,req) {
  requests.getRange(row,8).setValue('CANCELLED');
  requests.getRange(row,10).setValue(Utilities.formatDate(new Date(),TIME_ZONE,'yyyy-MM-dd HH:mm:ss'));
}
function normalizeDigits_(value) {
  return normalizeWhatsAppPhone_(value);
}



/* ---------------- CONSULTATION WORKSPACE ---------------- */
function getConsultationsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Consultations');
  const headers = [
    'requestId','patientName','patientPhone','appointmentDate','appointmentTime',
    'age','sex','allergies','currentMedicines','bloodPressure','pulse','respiratoryRate',
    'temperature','spo2','weight','height','clinicalHistory','examination',
    'assessmentDiagnosis','investigations','treatmentPlan','followUpAdvice',
    'followUpDate','noteStatus','createdAt','updatedAt'
  ];
  if (!sheet) sheet = ss.insertSheet('Consultations');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.getRange('A:Z').setNumberFormat('@');
  } else {
    const firstHeader = String(sheet.getRange(1,1).getDisplayValue() || '');
    if (firstHeader !== 'requestId') {
      throw new Error('The Consultations sheet exists but has an unexpected layout. Please do not edit it manually; contact support before proceeding.');
    }
  }
  return sheet;
}

function findConfirmedBookingForConsultation_(requestId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('No booking records were found.');
  const rows = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]) === String(requestId)) {
      if (String(r[7]).toUpperCase() !== 'CONFIRMED') {
        throw new Error('A consultation record can only be opened for a confirmed appointment.');
      }
      return {
        requestId:String(r[0]), appointmentDate:String(r[2]), appointmentTime:String(r[3]),
        patientName:String(r[4]), patientPhone:String(r[5]), visitType:String(r[6]),
        presentingComplaint:String(r[12] || ''), medicalHistory:String(r[13] || ''),
        videoLink:String(r[10] || '')
      };
    }
  }
  throw new Error('The appointment could not be found. Refresh the dashboard and try again.');
}

function adminGetConsultation(token, requestId) {
  requireAdminSession_(token);
  const booking = findConfirmedBookingForConsultation_(cleanText_(requestId,80));
  const sheet = getConsultationsSheet_();
  let consultation = {};
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2,1,sheet.getLastRow()-1,26).getDisplayValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === booking.requestId) {
        const r = rows[i];
        consultation = {
          age:r[5], sex:r[6], allergies:r[7], currentMedicines:r[8],
          bloodPressure:r[9], pulse:r[10], respiratoryRate:r[11], temperature:r[12],
          spo2:r[13], weight:r[14], height:r[15], clinicalHistory:r[16],
          examination:r[17], assessmentDiagnosis:r[18], investigations:r[19],
          treatmentPlan:r[20], followUpAdvice:r[21], followUpDate:r[22],
          noteStatus:r[23], createdAt:r[24], updatedAt:r[25]
        };
        break;
      }
    }
  }
  return {ok:true, booking:booking, consultation:consultation};
}

function adminSaveConsultation(token, requestId, payload, requestedStatus) {
  requireAdminSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const booking = findConfirmedBookingForConsultation_(cleanText_(requestId,80));
    const p = payload || {};
    const status = String(requestedStatus || 'DRAFT').toUpperCase() === 'COMPLETED' ? 'COMPLETED' : 'DRAFT';
    const fields = {
      age:cleanText_(p.age,20),
      sex:cleanText_(p.sex,30),
      allergies:cleanMultilineText_(p.allergies,1000),
      currentMedicines:cleanMultilineText_(p.currentMedicines,2000),
      bloodPressure:cleanText_(p.bloodPressure,50),
      pulse:cleanText_(p.pulse,30),
      respiratoryRate:cleanText_(p.respiratoryRate,30),
      temperature:cleanText_(p.temperature,30),
      spo2:cleanText_(p.spo2,30),
      weight:cleanText_(p.weight,30),
      height:cleanText_(p.height,30),
      clinicalHistory:cleanMultilineText_(p.clinicalHistory,5000),
      examination:cleanMultilineText_(p.examination,3000),
      assessmentDiagnosis:cleanMultilineText_(p.assessmentDiagnosis,2000),
      investigations:cleanMultilineText_(p.investigations,2000),
      treatmentPlan:cleanMultilineText_(p.treatmentPlan,4000),
      followUpAdvice:cleanMultilineText_(p.followUpAdvice,2000),
      followUpDate:/^\d{4}-\d{2}-\d{2}$/.test(String(p.followUpDate || '')) ? String(p.followUpDate) : ''
    };
    if (status === 'COMPLETED' && (!fields.assessmentDiagnosis || !fields.treatmentPlan)) {
      throw new Error('Add the assessment/diagnosis and treatment plan before marking this consultation completed.');
    }
    const sheet = getConsultationsSheet_();
    const now = Utilities.formatDate(new Date(),TIME_ZONE,'yyyy-MM-dd HH:mm:ss');
    let targetRow = -1, createdAt = now;
    if (sheet.getLastRow() >= 2) {
      const ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getDisplayValues();
      for (let i=0; i<ids.length; i++) {
        if (String(ids[i][0]) === booking.requestId) {
          targetRow = i + 2;
          createdAt = String(sheet.getRange(targetRow,25).getDisplayValue() || now);
          break;
        }
      }
    }
    const row = [
      booking.requestId, safeCell_(booking.patientName), safeCell_(booking.patientPhone),
      booking.appointmentDate, booking.appointmentTime,
      fields.age, fields.sex, fields.allergies, fields.currentMedicines,
      fields.bloodPressure, fields.pulse, fields.respiratoryRate, fields.temperature,
      fields.spo2, fields.weight, fields.height, fields.clinicalHistory, fields.examination,
      fields.assessmentDiagnosis, fields.investigations, fields.treatmentPlan,
      fields.followUpAdvice, fields.followUpDate, status, createdAt, now
    ].map(function(v){ return safeCell_(v); });
    if (targetRow > 0) sheet.getRange(targetRow,1,1,26).setValues([row]);
    else sheet.appendRow(row);
    return {ok:true, status:status, updatedAt:now, message:status==='COMPLETED'?'Consultation marked completed and saved.':'Consultation draft saved.'};
  } finally {
    lock.releaseLock();
  }
}


/* ---------------- ONE-TIME PRESCRIPTION HANDOFF ----------------
   Transfers only a short-lived, single-use payload to the authenticated
   prescription app. No patient details are placed in the URL.
*/
function adminCreatePrescriptionHandoff(token, requestId, formPayload) {
  requireAdminSession_(token);
  const booking = findConfirmedBookingForConsultation_(cleanText_(requestId,80));
  if (!/\b(test|fictional|demo)\b/i.test(booking.patientName)) {
    throw new Error('The connected AI Clinic app is a test-only prototype. For safety, handoff is restricted to a clearly labelled fictional/demo booking. Real patient information cannot be transferred.');
  }
  const p = formPayload || {};
  const field = function(key, limit) {
    return cleanMultilineText_(p[key] || '', limit);
  };
  const clinicalParts = [
    booking.presentingComplaint ? 'Presenting complaint: ' + booking.presentingComplaint : '',
    booking.medicalHistory ? 'Past medical history: ' + booking.medicalHistory : '',
    field('clinicalHistory',5000) ? 'Clinical history: ' + field('clinicalHistory',5000) : '',
    field('examination',3000) ? 'Examination / video assessment: ' + field('examination',3000) : '',
    field('bloodPressure',50) ? 'Blood pressure: ' + field('bloodPressure',50) : '',
    field('pulse',30) ? 'Pulse: ' + field('pulse',30) : '',
    field('respiratoryRate',30) ? 'Respiratory rate: ' + field('respiratoryRate',30) : '',
    field('temperature',30) ? 'Temperature: ' + field('temperature',30) : '',
    field('spo2',30) ? 'SpO2: ' + field('spo2',30) : '',
    field('investigations',2000) ? 'Relevant investigations: ' + field('investigations',2000) : ''
  ].filter(Boolean);
  const now = new Date();
  const handoff = {
    patient: {
      name: cleanText_(booking.patientName,80),
      age: cleanText_(p.age,3),
      sex: cleanText_(p.sex,30),
      weight: cleanText_(p.weight,10)
    },
    clinic: 'Dr. Bhupin Thakkar | Online Consultation',
    doctor: 'Dr. Bhupin Thakkar, MBBS, DNB (Internal Medicine)',
    prescriptionDate: Utilities.formatDate(now,TIME_ZONE,'yyyy-MM-dd'),
    prescriptionTime: Utilities.formatDate(now,TIME_ZONE,'HH:mm'),
    clinical: clinicalParts.join('\n'),
    assessment: field('assessmentDiagnosis',2000),
    history: [
      field('currentMedicines',2000),
      booking.medicalHistory
    ].filter(Boolean).join('\n'),
    allergies: field('allergies',1000),
    advice: field('treatmentPlan',4000),
    investigations: field('investigations',2000),
    followup: [
      field('followUpAdvice',2000),
      field('followUpDate',20) ? 'Follow-up date: ' + field('followUpDate',20) : ''
    ].filter(Boolean).join('\n')
  };
  const cache = CacheService.getScriptCache();
  const handoffToken = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  cache.put('RX_HANDOFF_' + handoffToken, JSON.stringify(handoff), 180);
  return {ok:true, token:handoffToken, expiresInSeconds:180};
}

function consumePrescriptionHandoff_(handoffToken) {
  if (!/^[A-Fa-f0-9]{64}$/.test(String(handoffToken || ''))) {
    return {ok:false,error:'This prescription handoff link is invalid or has expired. Return to the doctor dashboard and try again.'};
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const key = 'RX_HANDOFF_' + handoffToken;
    const raw = cache.get(key);
    if (!raw) return {ok:false,error:'This one-time prescription handoff has expired or was already used. Return to the doctor dashboard and create a new one.'};
    cache.remove(key);
    try {
      return {ok:true,payload:JSON.parse(raw)};
    } catch (err) {
      return {ok:false,error:'The prescription handoff could not be read. Please create a new one.'};
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- PRIORITY BOOKING AVAILABILITY CONTROL ---------------- */
function getPriorityBookingsEnabled_() {
  const value=PropertiesService.getScriptProperties().getProperty('PRIORITY_BOOKINGS_ENABLED');
  return value === null ? true : value === 'true';
}
function adminGetPrioritySettings(token) {
  requireAdminSession_(token);
  return {priorityEnabled:getPriorityBookingsEnabled_()};
}
function adminSetPriorityBookingsEnabled(token, enabled) {
  requireAdminSession_(token);
  if (typeof enabled !== 'boolean') throw new Error('Invalid priority booking setting.');
  PropertiesService.getScriptProperties().setProperty('PRIORITY_BOOKINGS_ENABLED',enabled?'true':'false');
  return {ok:true,priorityEnabled:enabled,
    message:enabled?'Priority response bookings are now visible to patients.':'Priority response bookings have been temporarily disabled.'};
}
