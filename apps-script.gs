/**
 * RealtorOne Content Sprint: Google Apps Script receiver.
 * Bound to a Google Sheet (Extensions > Apps Script). Free, no paid tools.
 *
 * The page sends three kinds of events (JSON, text/plain, no-cors):
 *   signup          sent when the realtor finishes the questionnaire.
 *                   -> adds one row to the "Responses" sheet with every field,
 *                   -> emails the realtor a confirmation (MailApp, sent from the sheet owner's Gmail),
 *                   -> emails the owner a "new signup" notification.
 *   payment_clicked sent when they open the payment link. -> Payment status = LINK_OPENED
 *   paid_claimed    sent when they say they have paid (or return from the payment page).
 *                   -> Payment status = PAID_CLAIMED (NOT verified: check Razorpay),
 *                   -> emails the realtor the welcome email, and the owner a "verify payment" note.
 * The "Delivery status" column shows what happened with the emails.
 * See SETUP.md.
 */

/* ===== CONFIG: keep in step with index.html ===== */
var PAGE_URL        = 'https://anatomyofacloser-cyber.github.io/realtorone-content-sprint/';
var LOGO_URL        = PAGE_URL + 'logo.png';   // The Realtor One logo, served by GitHub Pages. Empty = emails have no logo
var PAYMENT_URL     = 'https://rzp.io/rzp/ZRf2Lllg';   // same Razorpay Payment Link as in index.html. Empty = no pay button in the email
var WHATSAPP_NUMBER = '917302317285';   // digits with country code. Empty = no WhatsApp line in the emails
var OWNER_EMAIL     = '';          // who gets the "new signup" emails. Empty = the Gmail account that owns this sheet
var CURRENCY        = 'AED';
var PRICE           = 400;
var REGULAR_PRICE   = 610;         // same as index.html. Empty = the regular price is not mentioned in emails
var SENDER_NAME     = 'RealtorOne Content Sprint';
/* ================================================ */

/* ===== AUTOMATION CONFIG ===== */
var TZ                    = 'Asia/Dubai';
var DAILY_HOUR            = 9;     // daily content trigger hour (Dubai). Google runs it sometime in that hour
var SPRINT_DAYS           = 30;    // days of content per person
var REMINDER_AFTER_HOURS  = 24;    // send ONE payment reminder this long after signup
var REMINDER_MAX_AGE_DAYS = 7;     // ...but never to signups older than this (stops old rows getting a late nudge)
var QUOTA_RESERVE         = 10;    // keep this many MailApp sends free for confirmations / owner notices
var SEND_CLAIM_NOTE       = true;  // short "we are confirming your payment" note when someone taps "I've completed my payment"
var SEND_OWNER_DIGEST     = true;  // each morning, email YOU the WhatsApp queue (one email) when it has messages
var SEND_PAYMENT_REMINDER = false; // OFF: the flow is exactly 3 customer messages (confirmation, welcome, daily). true = also ONE reminder 24h+ after signup
/* ============================== */

var SHEET_NAME = 'Responses';
var CONTENT_TAB = 'Content Today';
var WA_TAB = 'WhatsApp Queue';
var HEADERS = ['Submission ID', 'Timestamp (UTC)', 'Name', 'Email', 'WhatsApp', 'Instagram',
               'Areas / specialization', 'Property types', 'Deal focus', 'Developers', 'Communities',
               'Target buyer', 'Target seller', 'Years of experience', 'Languages', 'Brand goals',
               'Posting frequency', 'Inspiration 1', 'Inspiration 2', 'Inspiration 3',
               'Delivery channel', 'Consent', 'Payment status', 'Payment updated (UTC)',
               'Delivery status', 'Owner notified', 'Reminder sent', 'Welcome sent', 'Content start date',
               'Last content sent', 'Content days sent', 'Notes', 'Raw (JSON)'];
/* Column positions are read from row 1 by NAME, so an older sheet keeps working:
   any header that is missing is added at the end the next time the script runs. */
var COL = {}, NCOLS = 0;

function doGet() {
  return ContentService.createTextOutput('Content Sprint endpoint is live.');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var body = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var d = JSON.parse(body);
    var sh = getSheet();
    var id = plain(d.id, 40);
    if (!id) { return json({ ok: false, error: 'missing id' }); }
    var ev = String(d.event || 'signup');
    var result;
    if (ev === 'signup') { result = onSignup(sh, d, id); }
    else if (ev === 'payment_clicked') { result = onPayment(sh, d, id, 'LINK_OPENED'); }
    else if (ev === 'paid_claimed') { result = onPayment(sh, d, id, 'PAID_CLAIMED'); }
    else { result = { ok: false, error: 'unknown event' }; }
    return json(result);
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function loadCols(sh) {
  var last = sh.getLastColumn();
  var cur = last ? sh.getRange(1, 1, 1, last).getValues()[0].map(function (x) { return String(x); }) : [];
  var changed = !last;
  HEADERS.forEach(function (h) { if (cur.indexOf(h) < 0) { cur.push(h); changed = true; } });
  if (changed) { sh.getRange(1, 1, 1, cur.length).setValues([cur]); sh.setFrozenRows(1); }
  COL = {}; cur.forEach(function (h, i) { if (h) { COL[h] = i + 1; } });
  NCOLS = cur.length;
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  loadCols(sh);
  return sh;
}

function setCell(sh, r, name, val) {
  if (COL[name]) { sh.getRange(r, COL[name]).setValue(val); }
}
function getCell(sh, r, name) {
  return COL[name] ? sh.getRange(r, COL[name]).getValue() : '';
}
/* Appends one row, filling the columns by header name. */
function putRow(sh, obj) {
  var arr = []; for (var i = 0; i < NCOLS; i++) { arr.push(''); }
  for (var k in obj) { if (obj.hasOwnProperty(k) && COL[k]) { arr[COL[k] - 1] = safe(obj[k]); } }
  sh.appendRow(arr);
  return sh.getLastRow();
}

function findRow(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) { return 0; }
  var ids = sh.getRange(2, COL['Submission ID'], last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === id) { return i + 2; } }
  return 0;
}

/* Rows used only to test the endpoint: saved, but no email is ever sent to them. */
function isTestRow(name, email) {
  return /^test$/i.test(String(name || '').trim()) || /\.invalid$/i.test(String(email || '').trim()) ||
         /@example\.(com|org|net)$/i.test(String(email || '').trim());
}
function isPaid(status) { var s = String(status || '').trim().toUpperCase(); return s === 'PAID' || s === 'PAID_CLAIMED'; }

/* ---------- signup ---------- */
function onSignup(sh, d, id) {
  if (findRow(sh, id)) { return { ok: true, duplicate: true }; }   // same submission sent twice
  var email = plain(d.email, 200);
  var emailOk = /^[^\s@,;<>"]+@[^\s@,;<>"]+\.[^\s@,;<>"]{2,}$/.test(email);
  var s = {
    id: id, name: plain(d.name, 80), email: email, whatsapp: plain(d.whatsapp, 40),
    instagram: plain(d.instagram, 40), areas: plain(d.areas, 200), types: list(d.propertyTypes),
    deal: plain(d.dealFocus, 40), devs: plain(d.developers, 200), comms: plain(d.communities, 200),
    buyers: list(d.targetBuyer), sellers: list(d.targetSeller), years: plain(d.years, 40),
    langs: list(d.languages), goals: list(d.goals), freq: plain(d.postingFrequency, 60),
    insp: (Array.isArray(d.inspiration) ? d.inspiration : []).slice(0, 3).map(function (x) { return plain(x, 40); }),
    channel: plain(d.delivery, 20), preview: d.preview === true
  };
  var test = isTestRow(s.name, s.email);
  var status = test ? 'TEST_ROW_NO_EMAIL' : (emailOk ? sendConfirmationEmail(s) : 'EMAIL_NOT_SENT: invalid email');
  var r = putRow(sh, {
    'Submission ID': id, 'Timestamp (UTC)': plain(d.ts, 40) || new Date().toISOString(), 'Name': s.name, 'Email': s.email,
    'WhatsApp': s.whatsapp, 'Instagram': s.instagram, 'Areas / specialization': s.areas, 'Property types': s.types,
    'Deal focus': s.deal, 'Developers': s.devs, 'Communities': s.comms, 'Target buyer': s.buyers, 'Target seller': s.sellers,
    'Years of experience': s.years, 'Languages': s.langs, 'Brand goals': s.goals, 'Posting frequency': s.freq,
    'Inspiration 1': s.insp[0] || '', 'Inspiration 2': s.insp[1] || '', 'Inspiration 3': s.insp[2] || '',
    'Delivery channel': s.channel, 'Consent': d.consent === true ? 'yes' : 'no', 'Payment status': 'PENDING',
    'Delivery status': status, 'Notes': test ? 'test row: no emails are sent to it' : (s.preview ? 'sent from a preview page' : '')
  });
  setCell(sh, r, 'Raw (JSON)', JSON.stringify(s));
  if (!test) {
    setCell(sh, r, 'Owner notified', sendOwnerEmail('New Content Sprint signup: ' + s.name,
      ownerSummary(s) + '\n\nPayment status: PENDING (they have not said they paid yet).'));
  }
  return { ok: true, status: status };
}

/* ---------- payment events ---------- */
/* The page can only say "I paid" (PAID_CLAIMED). That sends a short "we are confirming" note.
   The full welcome goes out when YOU set Payment status to PAID in the sheet (see sendPendingWelcomes). */
function onPayment(sh, d, id, newStatus) {
  var r = findRow(sh, id);
  if (!r) {   // the signup row was not found: keep a trace anyway
    r = putRow(sh, { 'Submission ID': id, 'Timestamp (UTC)': plain(d.ts, 40) || new Date().toISOString(),
                     'Payment status': 'PENDING', 'Notes': 'No signup row found for this ID' });
  }
  var cur = String(getCell(sh, r, 'Payment status'));
  if (isPaid(cur)) { return { ok: true, status: cur, duplicate: true }; }   // never step backwards (or repeat the note)
  setCell(sh, r, 'Payment status', newStatus);
  setCell(sh, r, 'Payment updated (UTC)', plain(d.ts, 40) || new Date().toISOString());
  if (d.preview === true) { setCell(sh, r, 'Notes', 'payment event from a preview page'); }
  if (newStatus === 'PAID_CLAIMED') {
    var s = rowToSignup(sh, r);
    if (!s.name) { s.name = 'unknown (ID ' + id + ')'; }
    var test = isTestRow(s.name, s.email);
    if (test) {
      setCell(sh, r, 'Delivery status', 'TEST_ROW_NO_EMAIL');
    } else {
      if (SEND_CLAIM_NOTE) {
        setCell(sh, r, 'Delivery status', validEmail(s.email) ? sendClaimNote(s) : 'EMAIL_NOT_SENT: invalid email');
      }
      setCell(sh, r, 'Owner notified', sendOwnerEmail(
        'Content Sprint: payment claimed by ' + s.name,
        s.name + ' says they have paid ' + CURRENCY + ' ' + PRICE + '.\n\nCheck Razorpay for the payment. If it is there, set Payment status to PAID in the sheet: that sends their Welcome email, and the daily 9 AM content starts for them.\n\n' + ownerSummary(s)));
    }
  }
  return { ok: true, status: newStatus };
}

function rowToSignup(sh, r) {
  var v = sh.getRange(r, 1, 1, NCOLS).getValues()[0];
  function g(h) { return COL[h] ? String(v[COL[h] - 1] == null ? '' : v[COL[h] - 1]).replace(/^'(?=[=+\-@])/, '') : ''; }   // safe() may have added a leading ' to numbers like +971...
  return { id: g('Submission ID'), name: g('Name'), email: g('Email'), whatsapp: g('WhatsApp'),
           instagram: g('Instagram'), areas: g('Areas / specialization'), types: g('Property types'),
           deal: g('Deal focus'), devs: g('Developers'), comms: g('Communities'), buyers: g('Target buyer'),
           sellers: g('Target seller'), years: g('Years of experience'), langs: g('Languages'),
           goals: g('Brand goals'), freq: g('Posting frequency'),
           insp: [g('Inspiration 1'), g('Inspiration 2'), g('Inspiration 3')].filter(Boolean),
           channel: g('Delivery channel') };
}

/* ====================================================================
   AUTOMATION (run setupTriggers once; see SETUP.md)
   ==================================================================== */

var AUTO_HANDLERS = ['sendReminders', 'hourlyTasks', 'sendDailyContent', 'onSheetEdit'];
function clearAutoTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (AUTO_HANDLERS.indexOf(t.getHandlerFunction()) >= 0) { ScriptApp.deleteTrigger(t); }
  });
}

/* Run this ONCE from the editor (and again after any change to DAILY_HOUR / TZ). It (re)creates:
   - an hourly trigger (sends any pending Welcome emails, plus the optional reminder)
   - a daily trigger at DAILY_HOUR Asia/Dubai (daily content: email + WhatsApp queue)
   - an "on edit" trigger: setting Payment status to PAID sends the Welcome within seconds
   and the "Content Today" and "WhatsApp Queue" tabs. */
function setupTriggers() {
  clearAutoTriggers();
  ScriptApp.newTrigger('hourlyTasks').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendDailyContent').timeBased().atHour(DAILY_HOUR).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  var sh = getSheet();   // makes sure the columns exist
  getContentSheet();     // makes sure the "Content Today" tab exists
  getWaSheet();          // makes sure the "WhatsApp Queue" tab exists
  try {                  // dropdown for Payment status (still accepts any other typed value)
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['PENDING', 'LINK_OPENED', 'PAID_CLAIMED', 'PAID'], true).setAllowInvalid(true).build();
    sh.getRange(2, COL['Payment status'], 1000, 1).setDataValidation(rule);
  } catch (e) { Logger.log('Dropdown not added: ' + e); }
  Logger.log('Triggers created: hourlyTasks (hourly), onSheetEdit (when you edit the sheet), sendDailyContent (daily, ' + DAILY_HOUR + ':00-' + (DAILY_HOUR + 1) + ':00 ' + TZ + ').');
}
function removeTriggers() { clearAutoTriggers(); Logger.log('Automation triggers removed.'); }

/* Hourly: welcome anyone you marked PAID (safety net for the on-edit trigger), plus the optional reminder. */
function hourlyTasks() {
  sendPendingWelcomes();
  if (SEND_PAYMENT_REMINDER) { runReminders(false); }
}

/* Fires when you edit the sheet (installable trigger, so it is allowed to send email). */
function onSheetEdit(e) {
  try {
    if (!e || !e.range) { return; }
    var sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) { return; }
    loadCols(sheet);
    var c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
    if (COL['Payment status'] >= c1 && COL['Payment status'] <= c2) { sendPendingWelcomes(); }
  } catch (err) { Logger.log('onSheetEdit: ' + err); }
}

function canSend(extra) { return MailApp.getRemainingDailyQuota() > QUOTA_RESERVE + (extra || 0); }
function validEmail(e) { return /^[^\s@,;<>"]+@[^\s@,;<>"]+\.[^\s@,;<>"]{2,}$/.test(String(e || '').trim()); }
function pad2(n) { return (n < 10 ? '0' : '') + n; }
function todayYmd() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
/* A cell can hold a real date (Sheets converts typed dates) or text like 2026-10-20. Returns 'yyyy-MM-dd' or ''. */
function toYmd(v) {
  if (v === '' || v == null) { return ''; }
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  var m = String(v).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]) : '';
}
function dayDiff(fromYmd, toYmdStr) {   // whole days from -> to
  var a = fromYmd.split('-'), b = toYmdStr.split('-');
  return Math.round((Date.UTC(+b[0], +b[1] - 1, +b[2]) - Date.UTC(+a[0], +a[1] - 1, +a[2])) / 86400000);
}

/* ---------- (a) hourly: ONE payment reminder ---------- */
function sendReminders() { return runReminders(false); }
function remindersDryRun() { return runReminders(true); }

function runReminders(dry) {
  if (!SEND_PAYMENT_REMINDER) { Logger.log('Payment reminder is switched off (SEND_PAYMENT_REMINDER = false).'); return { disabled: true }; }
  var lock = LockService.getScriptLock();
  var log = { checked: 0, due: 0, sent: 0, skipped: 0, dry: !!dry };
  try {
    lock.waitLock(20000);
    var sh = getSheet(), last = sh.getLastRow();
    if (last < 2) { return log; }
    var data = sh.getRange(2, 1, last - 1, NCOLS).getValues();
    function g(row, h) { return COL[h] ? row[COL[h] - 1] : ''; }
    var now = Date.now(), activePaid = 0;
    data.forEach(function (row) { if (String(g(row, 'Payment status')).toUpperCase().trim() === 'PAID' && toYmd(g(row, 'Last content sent')) !== todayYmd()) { activePaid++; } });
    for (var i = 0; i < data.length; i++) {
      var row = data[i], r = i + 2;
      log.checked++;
      var status = String(g(row, 'Payment status')).trim();
      var email = String(g(row, 'Email')).trim(), name = String(g(row, 'Name'));
      var ts = Date.parse(String(g(row, 'Timestamp (UTC)')));
      if (isPaid(status)) { continue; }                                                // paid (or says paid): no reminder
      if (String(g(row, 'Reminder sent')) !== '') { continue; }                       // once only
      if (String(g(row, 'Consent')).toLowerCase() !== 'yes') { continue; }
      if (!validEmail(email) || isTestRow(name, email)) { continue; }
      if (isNaN(ts)) { continue; }
      var ageH = (now - ts) / 3600000;
      if (ageH < REMINDER_AFTER_HOURS || ageH > REMINDER_MAX_AGE_DAYS * 24) { continue; }
      log.due++;
      if (dry) { Logger.log('WOULD REMIND: ' + name + ' <' + email + '> (' + status + ', ' + Math.round(ageH) + 'h after signup)'); continue; }
      if (!canSend(activePaid)) { log.skipped++; setCell(sh, r, 'Delivery status', 'REMINDER_NOT_SENT: daily quota used up (will retry next hour)'); continue; }
      var res = sendReminderEmail(rowToSignup(sh, r));
      if (res === 'REMINDER_SENT') {
        setCell(sh, r, 'Reminder sent', new Date().toISOString());
        log.sent++;
      } else {
        // mark failures too, so one bad address is not retried (and re-failed) every hour
        setCell(sh, r, 'Reminder sent', 'FAILED ' + new Date().toISOString());
      }
      setCell(sh, r, 'Delivery status', res);
    }
    Logger.log('Reminders: ' + JSON.stringify(log));
    return log;
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- (b) daily 9 AM Dubai: content email for PAID rows ---------- */
function getContentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONTENT_TAB);
  if (!sh) {
    sh = ss.insertSheet(CONTENT_TAB);
    sh.appendRow(['Date (yyyy-mm-dd, blank = today)', 'For (ALL, an email, or a Submission ID)', 'Link', 'Note (optional)']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* Today's entries from the "Content Today" tab. Blank date = today. A row with no link and no note is ignored. */
function readContentEntries(today) {
  var sh = getContentSheet(), last = sh.getLastRow(), out = [];
  if (last < 2) { return out; }
  var v = sh.getRange(2, 1, last - 1, 4).getValues();
  for (var i = 0; i < v.length; i++) {
    var link = String(v[i][2] || '').trim(), note = String(v[i][3] || '').trim();
    if (!link && !note) { continue; }
    var d = toYmd(v[i][0]);
    if (d && d !== today) { continue; }
    out.push({ row: i + 2, blankDate: !d, who: String(v[i][1] || '').trim(), link: link, note: note });
  }
  return out;
}
function entriesFor(entries, s) {
  var email = String(s.email || '').trim().toLowerCase(), id = String(s.id || '').trim().toLowerCase();
  return entries.filter(function (e) {
    var w = e.who.toLowerCase();
    return w === '' || w === 'all' || w === email || w === id;
  });
}

/* ---------- Welcome email: once, when Payment status becomes PAID ---------- */
function addDays(ymd, n) {
  var p = ymd.split('-');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)).toISOString().slice(0, 10);
}
function sendPendingWelcomes() { return withLock(function () { return doWelcomes(getSheet(), false, false); }); }
function welcomesDryRun() { return withLock(function () { return doWelcomes(getSheet(), true, false); }); }
function withLock(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); return fn(); } finally { try { lock.releaseLock(); } catch (x) {} }
}

/* Runs inside a lock. startToday = the 9 AM run itself calls this, so the first content goes out in the same run. */
function doWelcomes(sh, dry, startToday) {
  var log = { sent: 0, skipped: 0, quotaStopped: 0, dry: !!dry };
  var last = sh.getLastRow();
  if (last < 2) { return log; }
  var data = sh.getRange(2, 1, last - 1, NCOLS).getValues();
  function g(row, h) { return COL[h] ? row[COL[h] - 1] : ''; }
  var today = todayYmd();
  for (var i = 0; i < data.length; i++) {
    var row = data[i], r = i + 2;
    if (String(g(row, 'Payment status')).trim().toUpperCase() !== 'PAID') { continue; }   // PAID_CLAIMED does not trigger the welcome
    if (String(g(row, 'Welcome sent')) !== '') { continue; }                                // once only
    var name = String(g(row, 'Name')), email = String(g(row, 'Email')).trim();
    if (isTestRow(name, email)) { if (!dry) { setCell(sh, r, 'Welcome sent', 'SKIPPED test row'); } continue; }
    if (!validEmail(email)) {
      log.skipped++;
      if (!dry) { setCell(sh, r, 'Welcome sent', 'SKIPPED no valid email'); setCell(sh, r, 'Delivery status', 'EMAIL_NOT_SENT: invalid email'); }
      continue;
    }
    var start = toYmd(g(row, 'Content start date'));
    if (!start) { start = (startToday || Number(Utilities.formatDate(new Date(), TZ, 'H')) < DAILY_HOUR) ? today : addDays(today, 1); }
    if (dry) { Logger.log('WOULD WELCOME: ' + name + ' <' + email + '>, content starts ' + start); log.sent++; continue; }
    if (MailApp.getRemainingDailyQuota() < 1) {
      log.quotaStopped++;
      setCell(sh, r, 'Delivery status', 'WELCOME_NOT_SENT: daily quota used up (will retry)');
      break;
    }
    if (!toYmd(g(row, 'Content start date'))) { setCell(sh, r, 'Content start date', start); }
    var res = sendWelcomeEmail(rowToSignup(sh, r), start);
    setCell(sh, r, 'Delivery status', res);
    setCell(sh, r, 'Welcome sent', res === 'WELCOME_SENT' ? new Date().toISOString() : 'FAILED ' + new Date().toISOString());
    if (res === 'WELCOME_SENT') { log.sent++; }
  }
  Logger.log('Welcomes: ' + JSON.stringify(log));
  return log;
}

/* ---------- (b) every morning: the day's content, by the channel they chose ---------- */
function getWaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(WA_TAB);
  if (!sh) {
    sh = ss.insertSheet(WA_TAB);
    sh.appendRow(['Date', 'Day', 'Name', 'WhatsApp', 'Message', 'Click to send', 'Sent? (type x when done)', 'Submission ID']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function channelOf(s) {   // 'email' | 'whatsapp' | 'both'
  var c = String(s.channel || '').toLowerCase();
  var w = c.indexOf('whats') >= 0, m = c.indexOf('mail') >= 0;
  return (c.indexOf('both') >= 0 || (w && m)) ? 'both' : (w ? 'whatsapp' : 'email');
}
function waDigits(n) { return String(n || '').replace(/\D/g, '').replace(/^0+/, ''); }

function sendDailyContent() { return runDaily(false); }
function dailyDryRun() { return runDaily(true); }

function runDaily(dry) {
  var lock = LockService.getScriptLock();
  var log = { today: '', eligible: 0, emailed: 0, queued: 0, nothingToSend: 0, notYetStarted: 0, finished: 0, failed: 0, quotaStopped: 0, welcomes: 0, dry: !!dry };
  try {
    lock.waitLock(20000);
    var today = todayYmd(); log.today = today;
    var sh = getSheet(), last = sh.getLastRow();
    if (last < 2) { return log; }
    log.welcomes = doWelcomes(sh, dry, true).sent;     // anyone marked PAID but not yet welcomed gets it first
    var entries = readContentEntries(today);
    var data = sh.getRange(2, 1, last - 1, NCOLS).getValues();
    function g(row, h) { return COL[h] ? row[COL[h] - 1] : ''; }
    var queue = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i], r = i + 2;
      if (String(g(row, 'Payment status')).trim().toUpperCase() !== 'PAID') { continue; }   // owner sets PAID after checking Razorpay
      var email = String(g(row, 'Email')).trim(), name = String(g(row, 'Name'));
      if (isTestRow(name, email)) { continue; }
      var s = rowToSignup(sh, r), ch = channelOf(s);
      var wantEmail = ch !== 'whatsapp', wantWa = ch !== 'email';
      var emailOk = validEmail(email), waNum = waDigits(s.whatsapp), waOk = waNum.length >= 8;
      if (!(wantEmail && emailOk) && !(wantWa && waOk)) { continue; }
      log.eligible++;
      var start = toYmd(g(row, 'Content start date'));
      if (!start) { start = today; if (!dry) { setCell(sh, r, 'Content start date', today); } }   // blank = starts at this run
      var day = dayDiff(start, today) + 1;
      if (day < 1) { log.notYetStarted++; continue; }
      if (day > SPRINT_DAYS) {
        log.finished++;
        if (!dry && String(g(row, 'Delivery status')).indexOf('SPRINT_COMPLETE') !== 0) { setCell(sh, r, 'Delivery status', 'SPRINT_COMPLETE (' + SPRINT_DAYS + ' days)'); }
        continue;
      }
      if (toYmd(g(row, 'Last content sent')) === today) { continue; }                       // already delivered today
      var mine = entriesFor(entries, s);
      if (!mine.length) { log.nothingToSend++; continue; }                                    // nothing for this person today: skip
      if (dry) { Logger.log('WOULD DELIVER day ' + day + ' to ' + name + ' via ' + ch + ': ' + mine.length + ' item(s)'); log.emailed += wantEmail && emailOk ? 1 : 0; log.queued += wantWa && waOk ? 1 : 0; continue; }
      var parts = [], delivered = false;
      if (wantEmail && emailOk) {
        if (!canSend(0)) {
          log.quotaStopped++; parts.push('CONTENT_NOT_SENT: daily quota used up (day ' + day + ')');
        } else {
          var res = sendContentEmail(s, day, mine);
          if (res === 'CONTENT_SENT') { log.emailed++; delivered = true; parts.push('CONTENT_SENT day ' + day + '/' + SPRINT_DAYS); }
          else { log.failed++; parts.push(res + ' (day ' + day + ')'); }
        }
      }
      if (wantWa) {
        if (waOk) {
          queue.push({ date: today, day: day, name: name, number: waNum, message: waMessage(s, day, mine), id: s.id });
          log.queued++; delivered = true; parts.push('WHATSAPP_QUEUED day ' + day + '/' + SPRINT_DAYS);
        } else { parts.push('WHATSAPP_NOT_QUEUED: invalid number'); }
      }
      setCell(sh, r, 'Delivery status', parts.join('; '));
      if (delivered) {
        setCell(sh, r, 'Last content sent', today);
        setCell(sh, r, 'Content days sent', (Number(g(row, 'Content days sent')) || 0) + 1);
      }
    }
    if (!dry && queue.length) {
      writeWaQueue(queue);
      if (SEND_OWNER_DIGEST) { sendQueueDigest(queue); }
    }
    // entries with a blank date are "for today": stamp today's date so they are not sent again tomorrow
    if (!dry && (log.emailed + log.queued) > 0) {
      var cs = getContentSheet();
      entries.forEach(function (e) { if (e.blankDate) { cs.getRange(e.row, 1).setValue(today); } });
    }
    Logger.log('Daily content: ' + JSON.stringify(log));
    return log;
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function waLink(q) { return 'https://wa.me/' + q.number + '?text=' + encodeURIComponent(q.message); }

function writeWaQueue(queue) {
  var sh = getWaSheet();
  queue.forEach(function (q) {
    sh.appendRow([q.date, q.day, safe(q.name), safe('+' + q.number), q.message, '', '', q.id]);
    var r = sh.getLastRow();
    sh.getRange(r, 6).setFormula('=HYPERLINK("' + waLink(q).replace(/"/g, '%22') + '","Send on WhatsApp")');
  });
}

function sendQueueDigest(queue) {
  try {
    var to = OWNER_EMAIL || Session.getEffectiveUser().getEmail();
    if (!to || !canSend(0)) { return; }
    var url = SpreadsheetApp.getActiveSpreadsheet().getUrl ? SpreadsheetApp.getActiveSpreadsheet().getUrl() : '';
    var subject = 'WhatsApp queue: ' + queue.length + ' message' + (queue.length === 1 ? '' : 's') + ' to send today';
    var text = ['Today\'s WhatsApp messages are ready. Open the "' + WA_TAB + '" tab' + (url ? ' (' + url + ')' : '') + ' and tap "Send on WhatsApp" on each row, or use the links below.', ''];
    var items = '';
    queue.forEach(function (q) {
      text.push(q.name + ' (+' + q.number + ', day ' + q.day + '): ' + waLink(q));
      items += '<p style="margin:0 0 10px;font-family:' + EM.sans + ';font-size:15px;line-height:22px;color:' + EM.ink + ';"><b>' + esc(q.name) + '</b> <span style="color:' + EM.muted + ';">+' + esc(q.number) + ', day ' + q.day + '</span><br><a href="' + esc(waLink(q)) + '" target="_blank" style="color:' + EM.accent + ';font-weight:600;">Send on WhatsApp</a></p>';
    });
    var rows = eRow('<h1 style="margin:0 0 8px;font-family:' + EM.serif + ';font-weight:normal;font-size:26px;line-height:32px;color:' + EM.charcoal + ';">' + queue.length + ' WhatsApp message' + (queue.length === 1 ? '' : 's') + ' to send</h1>' +
      ePara('Each link opens WhatsApp with the message ready. Tap send.') + items +
      (url ? eButton(url, 'Open the WhatsApp Queue') : ''), '28px 32px 24px');
    MailApp.sendEmail({ to: to, subject: subject, body: text.join('\n'), htmlBody: eShell('Your WhatsApp queue for today is ready.', subject, rows), name: SENDER_NAME });
  } catch (err) { Logger.log('Digest not sent: ' + err); }
}

function waMessage(s, day, entries) {
  var L = ['Good morning ' + firstName(s.name) + '! Your RealtorOne Content Sprint, day ' + day + ' of ' + SPRINT_DAYS + ':', ''];
  entries.forEach(function (e) {
    if (e.note) { L.push(String(e.note).slice(0, 500)); }
    if (e.link) { L.push(e.link); }
    L.push('');
  });
  L.push('Reply here if you need anything.');
  return L.join('\n');
}

/* ---------- emails ---------- */
function ownerSummary(s) {
  return ['Name: ' + s.name, 'Email: ' + s.email, 'WhatsApp: ' + s.whatsapp, 'Instagram: @' + s.instagram,
    'Areas: ' + s.areas, 'Property types: ' + s.types, 'Deal focus: ' + s.deal,
    'Developers: ' + (s.devs || '-'), 'Communities: ' + (s.comms || '-'),
    'Target buyer: ' + s.buyers, 'Target seller: ' + s.sellers, 'Experience: ' + s.years,
    'Languages: ' + s.langs, 'Brand goals: ' + s.goals, 'Posts now: ' + s.freq,
    'Inspiration: ' + (Array.isArray(s.insp) ? s.insp.join(', ') : s.insp), 'Send content via: ' + s.channel,
    'Submission ID: ' + s.id].join('\n');
}

function sendOwnerEmail(subject, body) {
  try {
    var to = OWNER_EMAIL || Session.getEffectiveUser().getEmail();
    if (!to) { return 'OWNER_EMAIL_NOT_SENT: no owner address'; }
    if (MailApp.getRemainingDailyQuota() < 1) { return 'OWNER_EMAIL_NOT_SENT: daily quota used up'; }
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: SENDER_NAME });
    return 'OWNER_EMAIL_SENT';
  } catch (err) { return 'OWNER_EMAIL_FAILED: ' + String(err).slice(0, 120); }
}

function sendConfirmationEmail(s) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { return 'EMAIL_NOT_SENT: daily quota used up'; }
    MailApp.sendEmail({ to: s.email, subject: 'Your RealtorOne Content Sprint: exactly what you get, ' + firstName(s.name),
      body: confirmationText(s), htmlBody: confirmationHtml(s), name: SENDER_NAME });
    return 'CONFIRMATION_SENT';
  } catch (err) { return 'EMAIL_FAILED: ' + String(err).slice(0, 120); }
}

function sendClaimNote(s) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { return 'EMAIL_NOT_SENT: daily quota used up'; }
    MailApp.sendEmail({ to: s.email, subject: 'We are confirming your payment, ' + firstName(s.name),
      body: claimText(s), htmlBody: claimHtml(s), name: SENDER_NAME });
    return 'CLAIM_NOTE_SENT';
  } catch (err) { return 'EMAIL_FAILED: ' + String(err).slice(0, 120); }
}

/* startYmd = the day their first content goes out (yyyy-MM-dd) */
function sendWelcomeEmail(s, startYmd) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { return 'EMAIL_NOT_SENT: daily quota used up'; }
    MailApp.sendEmail({ to: s.email, subject: 'Welcome to the RealtorOne Content Sprint, ' + firstName(s.name),
      body: welcomeText(s, startYmd), htmlBody: welcomeHtml(s, startYmd), name: SENDER_NAME });
    return 'WELCOME_SENT';
  } catch (err) { return 'EMAIL_FAILED: ' + String(err).slice(0, 120); }
}

function money() { return CURRENCY + ' ' + PRICE; }
function regularMoney() { return REGULAR_PRICE ? CURRENCY + ' ' + REGULAR_PRICE : ''; }
/* "Founding offer: you get it for AED 400 / 30 days (regular price AED 610)" */
function offerLine() {
  return 'Founding offer: you get it for ' + money() + ' / 30 days' + (regularMoney() ? ' (regular price ' + regularMoney() + ')' : '') + '.';
}
function waUrl() { var d = String(WHATSAPP_NUMBER || '').replace(/\D/g, ''); return d ? 'https://wa.me/' + d : ''; }
function firstName(n) { return plain(n, 40).split(' ')[0] || 'there'; }

/* ---- shared wording: taken from the landing page so the emails say exactly what the page says ---- */
var DAY_ITEMS = [['Post 1', 'Static', 'Market insight'], ['Post 2', 'Static', 'Personal-brand authority'],
                 ['Carousel 1', 'Carousel', 'Educational topic'], ['Carousel 2', 'Carousel', 'Community insight'],
                 ['Carousel 3', 'Carousel', 'Buyer / seller psychology']];
var YOU_GET = ['2 static posts a day', '3 carousel posts a day', 'Captions included', 'Dubai real-estate topics and current market trends',
               'Content relevant to your specialization', 'Ready-to-post creative', 'No content planning required'];
var QUOTE_LINE = 'You don\'t know which post will start the conversation that becomes your next client.';
var DISCLAIMER = 'Content keeps you visible. It does not guarantee leads or deals.';
var WRAP_LINE = '30 days x 5 posts/day = 150 opportunities to be seen.';

function channelPhrase(s) {
  var ch = channelOf(s), wa = s.whatsapp ? ' (' + s.whatsapp + ')' : '', em = s.email ? ' (' + s.email + ')' : '';
  return ch === 'email' ? 'by email' + em : (ch === 'whatsapp' ? 'on WhatsApp' + wa : 'by email' + em + ' and on WhatsApp' + wa);
}
/* "today" / "tomorrow, Saturday 17 October" / "on Monday 20 October" + the time */
function startPhrase(startYmd) {
  if (!startYmd) { return 'at the next 9:00 AM Dubai time'; }
  var diff = dayDiff(todayYmd(), startYmd);
  var p = startYmd.split('-');
  var nice = Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12)), 'UTC', 'EEEE d MMMM');
  return (diff <= 0 ? 'today' : (diff === 1 ? 'tomorrow, ' + nice : 'on ' + nice)) + ' at 9:00 AM Dubai time';
}
function hasContentToday(startYmd) { return startYmd && dayDiff(todayYmd(), startYmd) <= 0; }

/* 1) CONFIRMATION (on signup) */
function confirmationText(s) {
  var L = ['Hi ' + firstName(s.name) + ',', '',
    'Thanks for completing the onboarding. Here is exactly what the RealtorOne Content Sprint is, and what happens next.', '',
    'WHAT IT IS',
    'Every morning at 9 AM Dubai time, you receive ready-to-post content built around your market, your specialization and your personal brand.',
    '30 days. 150 pieces of content. One less thing to worry about.', '',
    'WHAT ONE CONTENT DAY LOOKS LIKE'];
  DAY_ITEMS.forEach(function (d) { L.push('- ' + d[0] + ' (' + d[1] + '): ' + d[2]); });
  L.push('Captions are included with every post.', '',
    'THE MATHS', '5 posts a day x 30 days = 150 pieces of content (60 static posts + 90 carousels).', '',
    'BUILT AROUND YOU',
    'Area: ' + (s.areas || '-'), 'Property: ' + (s.types || '-'), 'Focus: ' + (s.deal || '-'), 'Content sent: ' + channelPhrase(s), '',
    'FOUNDING OFFER',
    'You get it for ' + money() + ' / 30 days' + (regularMoney() ? ' (regular price ' + regularMoney() + ')' : '') + '.',
    'Only 10 founding spots, because every realtor\'s content is personalised.', '');
  L.push(PAYMENT_URL ? 'Complete payment: ' + PAYMENT_URL : 'Complete payment on the page: ' + PAGE_URL);
  L.push('', 'WHAT HAPPENS NEXT',
    '1. Complete your payment with the link above.',
    '2. We check it and send you a welcome email with your start date.',
    '3. From the next 9 AM Dubai time, your content arrives ' + channelPhrase(s) + ', every morning for 30 days.', '', DISCLAIMER);
  if (waUrl()) { L.push('', 'Questions? WhatsApp us: ' + waUrl()); }
  L.push('', '--', 'RealtorOne Content Sprint, a RealtorOne project');
  return L.join('\n');
}

/* (short note when they tap "I've completed my payment") */
function claimText(s) {
  var L = ['Hi ' + firstName(s.name) + ',', '',
    'Thanks, we have your note that you completed the payment. We are confirming it now.', '',
    'As soon as it is confirmed you get your welcome email with your start date. Nothing else for you to do.'];
  if (waUrl()) { L.push('', 'Questions? WhatsApp us: ' + waUrl()); }
  L.push('', '--', 'RealtorOne Content Sprint, a RealtorOne project');
  return L.join('\n');
}

/* 2) WELCOME (when the owner sets Payment status = PAID) */
function welcomeText(s, startYmd) {
  var L = ['Hi ' + firstName(s.name) + ',', '', 'Welcome to the RealtorOne Content Sprint. Your payment is confirmed.', '',
    'WHEN IT STARTS',
    'Your first day of content arrives ' + startPhrase(startYmd) + ', then every morning for 30 days.', '',
    'HOW IT IS DELIVERED',
    'Your content comes ' + channelPhrase(s) + '.', '',
    'WHAT YOU GET'];
  YOU_GET.forEach(function (x) { L.push('- ' + x); });
  L.push('', 'WHAT A CONTENT DAY LOOKS LIKE', '9:00 AM Dubai: your content for the day lands, ready to post.');
  DAY_ITEMS.forEach(function (d) { L.push('- ' + d[0] + ' (' + d[1] + '): ' + d[2]); });
  L.push('Captions included with every post.', '',
    'THE MATHS', '5 posts a day x 30 days = 150 pieces of content (60 static posts + 90 carousels).', '',
    'WHAT WE NEED FROM YOU',
    '- Keep ' + (channelOf(s) === 'email' ? 'an eye on your inbox' : (channelOf(s) === 'whatsapp' ? 'WhatsApp open' : 'your inbox and WhatsApp open')) + ' at 9 AM Dubai time, and post what you receive.',
    '- Reply to this email if your area, specialization or delivery channel changes.',
    '- Check spam / Promotions once, and mark us as not spam.', '',
    'AT THE END OF YOUR 30 DAYS', WRAP_LINE, '', QUOTE_LINE, DISCLAIMER);
  if (waUrl()) { L.push('', 'Questions? WhatsApp us: ' + waUrl()); }
  L.push('', '--', 'RealtorOne Content Sprint, a RealtorOne project');
  return L.join('\n');
}

function SPOTS_TEXT() { return 'only 10 founding spots'; }

/* ===== HTML email (table layout, inline CSS only, 600px max) ===== */
var EM = { charcoal: '#1c1c1e', cream: '#f6f1e7', ink: '#2a2a2c', muted: '#6e6a62', line: '#e3dccd', accent: '#b9852a',
  serif: "Georgia,'Times New Roman',Times,serif", sans: "-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif" };

function esc(x) {
  return String(x == null ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function eLabel(t) { return '<div style="font-family:' + EM.sans + ';font-size:11px;line-height:16px;letter-spacing:2px;text-transform:uppercase;font-weight:600;color:' + EM.accent + ';margin:0 0 10px;">' + esc(t) + '</div>'; }
function ePara(t) { return '<p style="margin:0 0 12px;font-family:' + EM.sans + ';font-size:16px;line-height:25px;color:' + EM.ink + ';">' + esc(t) + '</p>'; }
function eRow(inner, pad) { return '<tr><td style="background:' + EM.cream + ';padding:' + (pad || '8px 32px 18px') + ';">' + inner + '</td></tr>'; }
function eButton(href, text) {
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:6px 0 4px;"><tr><td style="background:' + EM.charcoal + ';border-radius:10px;">' +
    '<a href="' + esc(href) + '" target="_blank" style="display:block;color:' + EM.cream + ';font-family:' + EM.sans + ';font-size:16px;line-height:20px;font-weight:600;text-decoration:none;padding:15px 26px;">' + esc(text) + '</a></td></tr></table>';
}
function eShell(preheader, title, rows) {
  var h = '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light"><title>' + esc(title) + '</title></head>';
  h += '<body style="margin:0;padding:0;background:#ebe5d8;">';
  h += '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:#ebe5d8;font-size:1px;line-height:1px;">' + esc(preheader) + '</div>';
  h += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ebe5d8;"><tr><td align="center" style="padding:20px 10px;">';
  h += '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px;border-radius:12px;overflow:hidden;">';
  var logo = LOGO_URL
    ? '<td width="72" valign="middle" style="padding:0 16px 0 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td style="background:#ffffff;border-radius:10px;padding:5px 7px;"><img src="' + esc(LOGO_URL) + '" width="52" alt="The Realtor One" style="display:block;width:52px;height:auto;border:0;"></td></tr></table></td>'
    : '';
  h += '<tr><td style="background:' + EM.charcoal + ';padding:26px 32px 24px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' + logo +
    '<td valign="middle"><div style="font-family:' + EM.sans + ';font-size:12px;line-height:18px;letter-spacing:4px;text-transform:uppercase;font-weight:600;color:' + EM.cream + ';">RealtorOne Content Sprint</div>' +
    '<div style="height:2px;width:36px;background:' + EM.accent + ';font-size:1px;line-height:2px;margin:10px 0 10px;">&nbsp;</div>' +
    '<div style="font-family:' + EM.serif + ';font-style:italic;font-size:17px;line-height:23px;color:#e0b865;">Invisible Habits. Visible Results.</div></td></tr></table></td></tr>';
  h += rows;
  h += '<tr><td align="center" style="background:' + EM.cream + ';padding:14px 32px 30px;"><div style="font-family:' + EM.sans + ';font-size:12px;line-height:18px;color:' + EM.muted + ';">A RealtorOne project &middot; You&rsquo;re receiving this because you signed up for the Content Sprint.</div></td></tr>';
  return h + '</table></td></tr></table></body></html>';
}
function eDetails(pairs) {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border:1px solid ' + EM.line + ';border-radius:10px;">' +
    pairs.map(function (p, i) {
      return '<tr><td style="padding:10px 16px;font-family:' + EM.sans + ';font-size:13px;color:' + EM.muted + ';width:34%;vertical-align:top;' + (i ? 'border-top:1px solid ' + EM.line + ';' : '') + '">' + esc(p[0]) +
        '</td><td style="padding:10px 16px;font-family:' + EM.sans + ';font-size:15px;color:' + EM.ink + ';vertical-align:top;' + (i ? 'border-top:1px solid ' + EM.line + ';' : '') + '">' + esc(p[1] || '-') + '</td></tr>';
    }).join('') + '</table>';
}
function eWa() {
  return waUrl() ? '<p style="margin:8px 0 0;font-family:' + EM.sans + ';font-size:14px;line-height:21px;color:' + EM.muted + ';">Questions? <a href="' + esc(waUrl()) + '" target="_blank" style="color:' + EM.ink + ';">Message us on WhatsApp</a>.</p>' : '';
}

/* ---- small HTML blocks (inline CSS, table layout) ---- */
function eHello(s) { return '<p style="margin:0 0 6px;font-family:' + EM.serif + ';font-size:20px;line-height:28px;color:' + EM.charcoal + ';">Hi ' + esc(firstName(s.name)) + ',</p>'; }
function eH1(t, size) { return '<h1 style="margin:0 0 12px;font-family:' + EM.serif + ';font-weight:normal;font-size:' + (size || 30) + 'px;line-height:' + ((size || 30) + 6) + 'px;color:' + EM.charcoal + ';">' + esc(t) + '</h1>'; }
function eStep(n, t) {
  return '<tr><td width="40" valign="top" style="padding:0 0 12px;"><div style="width:28px;height:28px;border-radius:14px;background:' + EM.charcoal + ';color:' + EM.cream + ';font-family:' + EM.serif + ';font-size:15px;line-height:28px;text-align:center;">' + n + '</div></td>' +
    '<td valign="top" style="padding:3px 0 12px;font-family:' + EM.sans + ';font-size:15px;line-height:23px;color:' + EM.ink + ';">' + t + '</td></tr>';
}
function eSteps(items) { return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' + items.map(function (t, i) { return eStep(i + 1, t); }).join('') + '</table>'; }
function eChecks(items) {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' + items.map(function (t) {
    return '<tr><td width="26" valign="top" style="padding:0 0 8px;font-family:' + EM.sans + ';font-size:15px;line-height:23px;color:' + EM.accent + ';font-weight:700;">&#10003;</td>' +
      '<td valign="top" style="padding:0 0 8px;font-family:' + EM.sans + ';font-size:15px;line-height:23px;color:' + EM.ink + ';">' + t + '</td></tr>';
  }).join('') + '</table>';
}
function eDay() {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border:1px solid ' + EM.line + ';border-radius:10px;">' +
    '<tr><td colspan="2" style="padding:11px 16px;font-family:' + EM.sans + ';font-size:13px;line-height:19px;font-weight:700;color:' + EM.accent + ';letter-spacing:1px;">9:00 AM &middot; DUBAI</td></tr>' +
    DAY_ITEMS.map(function (d) {
      return '<tr><td style="padding:10px 16px;border-top:1px solid ' + EM.line + ';font-family:' + EM.sans + ';font-size:15px;line-height:22px;color:' + EM.ink + ';"><b>' + esc(d[0]) + '</b> <span style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:' + EM.muted + ';">&nbsp;' + esc(d[1]) + '</span></td>' +
        '<td align="right" style="padding:10px 16px;border-top:1px solid ' + EM.line + ';font-family:' + EM.sans + ';font-size:14px;line-height:21px;color:' + EM.muted + ';">' + esc(d[2]) + '</td></tr>';
    }).join('') + '</table>' +
    '<p style="margin:8px 0 0;font-family:' + EM.sans + ';font-size:14px;line-height:21px;color:' + EM.muted + ';">Captions included with every post.</p>';
}
function eMaths() {
  function cell(b, t) { return '<td width="33%" align="center" style="padding:14px 4px;background:#fff;border:1px solid ' + EM.line + ';border-radius:10px;"><div style="font-family:' + EM.serif + ';font-size:28px;line-height:32px;color:' + EM.charcoal + ';">' + b + '</div><div style="font-family:' + EM.sans + ';font-size:12px;line-height:17px;color:' + EM.muted + ';">' + t + '</div></td>'; }
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="6" border="0" style="width:100%;"><tr>' + cell('5', 'posts a day') + cell('30', 'days') + cell('150', 'pieces of content') + '</tr></table>' +
    '<p style="margin:8px 0 0;font-family:' + EM.sans + ';font-size:14px;line-height:21px;color:' + EM.ink + ';"><b>60 static posts + 90 carousels.</b> The proof of the work, not the point of it.</p>';
}
function ePrice() {
  return '<div style="background:#fff;border:1px solid ' + EM.accent + ';border-radius:12px;padding:16px 18px;">' +
    '<div style="font-family:' + EM.sans + ';font-size:11px;line-height:16px;letter-spacing:2px;text-transform:uppercase;font-weight:700;color:' + EM.accent + ';">Founding offer</div>' +
    (regularMoney() ? '<div style="margin-top:6px;font-family:' + EM.sans + ';font-size:14px;line-height:20px;color:' + EM.muted + ';">Regular price <span style="text-decoration:line-through;">' + esc(regularMoney()) + '</span></div>' : '') +
    '<div style="font-family:' + EM.serif + ';font-size:34px;line-height:40px;color:' + EM.charcoal + ';">' + esc(money()) + ' <span style="font-family:' + EM.sans + ';font-size:14px;color:' + EM.muted + ';">/ 30 days</span></div>' +
    '<div style="font-family:' + EM.sans + ';font-size:14px;line-height:21px;color:' + EM.ink + ';">You get it for ' + esc(money()) + '. <b>Only 10 founding spots</b>, because every realtor&rsquo;s content is personalised.</div></div>';
}
function eQuote(top) { return '<p style="margin:' + (top === undefined ? 16 : top) + 'px 0 0;font-family:' + EM.serif + ';font-style:italic;font-size:17px;line-height:25px;color:' + EM.ink + ';">' + esc(QUOTE_LINE.replace("don't", 'don\u2019t')) + '</p>' +
  '<p style="margin:8px 0 0;font-family:' + EM.sans + ';font-size:12px;line-height:18px;color:' + EM.muted + ';">' + esc(DISCLAIMER) + '</p>'; }

/* 1) CONFIRMATION */
function confirmationHtml(s) {
  var rows = eRow(eHello(s) + eH1('Here is exactly what you get', 30) +
      ePara('Thanks for completing the onboarding. This is what the RealtorOne Content Sprint is, and what happens next.'), '32px 32px 8px') +
    eRow(eLabel('What it is') +
      '<p style="margin:0 0 10px;font-family:' + EM.serif + ';font-size:20px;line-height:28px;color:' + EM.charcoal + ';">Every morning at 9 AM Dubai time, receive ready-to-post content built around your market, your specialization and your personal brand.</p>' +
      ePara('30 days. 150 pieces of content. One less thing to worry about.'), '8px 32px 14px') +
    eRow(eLabel('What one content day looks like') + eDay(), '8px 32px 18px') +
    eRow(eLabel('The maths') + eMaths(), '8px 32px 18px') +
    eRow(eLabel('Built around you') + eDetails([['Area', s.areas], ['Property', s.types], ['Focus', s.deal], ['Content sent', channelPhrase(s)]]), '8px 32px 18px') +
    eRow(ePrice() + '<div style="height:14px;line-height:14px;font-size:1px;">&nbsp;</div>' +
      (PAYMENT_URL ? eButton(PAYMENT_URL, 'Complete payment \u00b7 ' + money()) : eButton(PAGE_URL, 'Back to the page to pay')), '8px 32px 10px') +
    eRow(eLabel('What happens next') + eSteps([
      '<b>Complete your payment</b> with the button above.',
      '<b>We check it</b> and send you a welcome email with your start date.',
      '<b>From the next 9 AM Dubai time</b> your content arrives ' + esc(channelPhrase(s)) + ', every morning for 30 days.']), '8px 32px 6px') +
    eRow(eQuote(0) + eWa(), '4px 32px 26px');
  return eShell('Exactly what you get, and one step left: complete your payment.', 'Your RealtorOne Content Sprint', rows);
}

/* short claim note */
function claimHtml(s) {
  var rows = eRow(eHello(s) + eH1('We are confirming your payment', 28) +
      ePara('Thanks, we have your note that you completed the payment. We are confirming it now.') +
      ePara('As soon as it is confirmed you get your welcome email with your start date. Nothing else for you to do.') + eWa(), '32px 32px 26px');
  return eShell('Thanks. We are confirming your payment.', 'We are confirming your payment', rows);
}

/* 2) WELCOME */
function welcomeHtml(s, startYmd) {
  var ch = channelOf(s);
  var rows = eRow(eHello(s) + eH1('Welcome to the RealtorOne Content Sprint', 32) +
      ePara('Your payment is confirmed. Here is what to expect.'), '32px 32px 8px') +
    eRow(eLabel('When it starts') +
      '<p style="margin:0 0 6px;font-family:' + EM.serif + ';font-size:22px;line-height:30px;color:' + EM.charcoal + ';">Your first day of content arrives ' + esc(startPhrase(startYmd)) + '.</p>' +
      ePara('Then every morning for 30 days.'), '8px 32px 12px') +
    eRow(eLabel('How it is delivered') + ePara('Your content comes ' + channelPhrase(s) + '.' +
      (ch === 'email' ? '' : ' We message you on WhatsApp each morning.')), '8px 32px 12px') +
    eRow(eLabel('What you get') + eChecks(YOU_GET.map(esc)), '8px 32px 12px') +
    eRow(eLabel('What a content day looks like') + eDay(), '8px 32px 18px') +
    eRow(eLabel('The maths') + eMaths(), '8px 32px 18px') +
    eRow(eLabel('What we need from you') + eChecks([
      'Keep ' + (ch === 'email' ? 'an eye on your inbox' : (ch === 'whatsapp' ? 'WhatsApp open' : 'your inbox and WhatsApp open')) + ' at 9 AM Dubai time, and post what you receive.',
      'Reply to this email if your area, specialization or delivery channel changes.',
      'Check spam / Promotions once, and mark us as not spam.']), '8px 32px 18px') +
    eRow('<div style="border:1px solid ' + EM.accent + ';border-radius:12px;padding:16px;text-align:center;font-family:' + EM.serif + ';font-size:21px;line-height:28px;color:' + EM.charcoal + ';">' +
      '<div style="font-family:' + EM.sans + ';font-size:11px;letter-spacing:2px;text-transform:uppercase;color:' + EM.accent + ';font-weight:700;margin-bottom:6px;">At the end of your 30 days</div>30 days &times; 5 posts/day = 150 opportunities to be seen.</div>' +
      eQuote() + eWa(), '8px 32px 26px');
  return eShell('Welcome. Your 30 days of ready-to-post content are about to begin.', 'Welcome to the RealtorOne Content Sprint', rows);
}

/* ---------- reminder + daily content emails ---------- */
function sendReminderEmail(s) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { return 'REMINDER_NOT_SENT: daily quota used up'; }
    MailApp.sendEmail({ to: s.email, subject: 'Your Content Sprint details are saved, ' + firstName(s.name),
      body: reminderText(s), htmlBody: reminderHtml(s), name: SENDER_NAME });
    return 'REMINDER_SENT';
  } catch (err) { return 'REMINDER_FAILED: ' + String(err).slice(0, 120); }
}
function reminderText(s) {
  var L = ['Hi ' + firstName(s.name) + ',', '',
    'A quick note: your Content Sprint onboarding details are saved, but your founding spot is not secured until payment is made.', '',
    offerLine(), 'Only 10 founding spots, because each realtor\'s content is personalised around their market, specialization and personal brand.', ''];
  L.push(PAYMENT_URL ? 'Secure your spot: ' + PAYMENT_URL : 'Secure your spot on the page: ' + PAGE_URL);
  L.push('', 'Already paid? Just reply to this email' + (waUrl() ? ' or message us on WhatsApp (' + waUrl() + ')' : '') + ' and we will match it.',
    'This is the only reminder we will send.', '', '--', 'RealtorOne Content Sprint, a RealtorOne project');
  return L.join('\n');
}
function reminderHtml(s) {
  var rows = eRow('<p style="margin:0 0 6px;font-family:' + EM.serif + ';font-size:20px;line-height:28px;color:' + EM.charcoal + ';">Hi ' + esc(firstName(s.name)) + ',</p>' +
      ePara('A quick note: your Content Sprint onboarding details are saved, but your founding spot is not secured until payment is made.'), '32px 32px 8px') +
    eRow(eLabel('Founding offer') +
      '<p style="margin:0 0 12px;font-family:' + EM.sans + ';font-size:16px;line-height:25px;color:' + EM.ink + ';">You get it for <b>' + esc(money()) + '</b> / 30 days' + (regularMoney() ? ' <span style="color:' + EM.muted + ';">(regular price <span style="text-decoration:line-through;">' + esc(regularMoney()) + '</span>)</span>' : '') + '.</p>' +
      ePara('Only 10 founding spots, because each realtor\'s content is personalised around their market, specialization and personal brand.') +
      (PAYMENT_URL ? eButton(PAYMENT_URL, 'Secure my spot for ' + money()) : eButton(PAGE_URL, 'Back to the page')) +
      '<p style="margin:12px 0 0;font-family:' + EM.sans + ';font-size:14px;line-height:21px;color:' + EM.muted + ';">Already paid? Just reply to this email and we will match it. This is the only reminder we will send.</p>' + eWa(), '8px 32px 26px');
  return eShell('Your details are saved. One step left to secure your founding spot.', 'Your Content Sprint details are saved', rows);
}

function safeLink(u) { return /^https?:\/\/\S+$/i.test(String(u || '').trim()) ? String(u).trim() : ''; }
function sendContentEmail(s, day, entries) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { return 'CONTENT_NOT_SENT: daily quota used up'; }
    MailApp.sendEmail({ to: s.email, subject: 'Your Content Sprint: day ' + day + ' of ' + SPRINT_DAYS,
      body: contentText(s, day, entries), htmlBody: contentHtml(s, day, entries), name: SENDER_NAME });
    return 'CONTENT_SENT';
  } catch (err) { return 'CONTENT_FAILED: ' + String(err).slice(0, 120); }
}
function contentText(s, day, entries) {
  var L = ['Good morning ' + firstName(s.name) + ',', '', 'Your ready-to-post content for day ' + day + ' of ' + SPRINT_DAYS + ':', ''];
  entries.forEach(function (e) {
    if (e.note) { L.push(e.note); }
    if (e.link) { L.push(e.link); }
    L.push('');
  });
  L.push('Day ' + day + ' of ' + SPRINT_DAYS + (day === SPRINT_DAYS ? ': this is your last day.' : '.'));
  if (waUrl()) { L.push('Questions? WhatsApp us: ' + waUrl()); }
  L.push('', '--', 'RealtorOne Content Sprint, a RealtorOne project');
  return L.join('\n');
}
function contentHtml(s, day, entries) {
  var items = entries.map(function (e) {
    var link = safeLink(e.link);
    return '<div style="margin:0 0 16px;">' +
      (e.note ? '<p style="margin:0 0 8px;font-family:' + EM.sans + ';font-size:16px;line-height:25px;color:' + EM.ink + ';white-space:pre-line;">' + esc(e.note) + '</p>' : '') +
      (link ? eButton(link, 'Open today\'s content') : (e.link ? '<p style="margin:0;font-family:' + EM.sans + ';font-size:15px;color:' + EM.ink + ';">' + esc(e.link) + '</p>' : '')) + '</div>';
  }).join('');
  var rows = eRow('<p style="margin:0 0 6px;font-family:' + EM.serif + ';font-size:20px;line-height:28px;color:' + EM.charcoal + ';">Good morning ' + esc(firstName(s.name)) + ',</p>' +
      '<h1 style="margin:0 0 8px;font-family:' + EM.serif + ';font-weight:normal;font-size:30px;line-height:36px;color:' + EM.charcoal + ';">Day ' + day + ' of ' + SPRINT_DAYS + '</h1>' +
      ePara('Your ready-to-post content for today:'), '32px 32px 8px') +
    eRow(items, '8px 32px 10px') +
    eRow('<p style="margin:0;font-family:' + EM.sans + ';font-size:13px;line-height:19px;color:' + EM.muted + ';">' + (day === SPRINT_DAYS ? 'This is your last day of the 30-day Sprint. ' : '') + 'Post it today, and keep the streak going.</p>' + eWa(), '8px 32px 26px');
  return eShell('Your ready-to-post content for day ' + day + ' is here.', 'Your Content Sprint: day ' + day, rows);
}

/* ---------- helpers ---------- */
function plain(s, max) { return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').slice(0, max || 60).trim(); }
function list(a) { return (Array.isArray(a) ? a : [a]).filter(function (x) { return x; }).map(function (x) { return plain(x, 60); }).slice(0, 12).join(', '); }
// Stops spreadsheet formula injection (a value starting with = + - @ is stored as text).
function safe(v) {
  var s = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/* Optional: run from the editor to get the confirmation, claim note and welcome in your own inbox (sent to YOU only). */
function testEmailsToMe() {
  var me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  var s = { id: 'TEST', name: 'Test Realtor', email: me, whatsapp: '+971500000000', instagram: 'test.realtor', areas: 'Dubai Marina',
            types: 'Apartments', deal: 'Both', devs: '', comms: '', buyers: 'Investors', sellers: 'Landlords', years: '3–5 years',
            langs: 'English', goals: 'Get more enquiries', freq: 'Daily', insp: ['@example'], channel: 'Both' };
  Logger.log(sendConfirmationEmail(s) + ' / ' + sendClaimNote(s) + ' / ' + sendWelcomeEmail(s, addDays(todayYmd(), 1)) + ' to ' + me);
}
