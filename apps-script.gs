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
var PAYMENT_URL     = 'https://rzp.io/rzp/ZRf2Lllg';   // same Razorpay Payment Link as in index.html. Empty = no pay button in the email
var WHATSAPP_NUMBER = '917302317285';   // digits with country code. Empty = no WhatsApp line in the emails
var OWNER_EMAIL     = '';          // who gets the "new signup" emails. Empty = the Gmail account that owns this sheet
var CURRENCY        = 'AED';
var PRICE           = 400;
var REGULAR_PRICE   = 610;         // same as index.html. Empty = the regular price is not mentioned in emails
var SENDER_NAME     = 'RealtorOne Content Sprint';
/* ================================================ */

var SHEET_NAME = 'Responses';
var HEADERS = ['Submission ID', 'Timestamp (UTC)', 'Name', 'Email', 'WhatsApp', 'Instagram',
               'Areas / specialization', 'Property types', 'Deal focus', 'Developers', 'Communities',
               'Target buyer', 'Target seller', 'Years of experience', 'Languages', 'Brand goals',
               'Posting frequency', 'Inspiration 1', 'Inspiration 2', 'Inspiration 3',
               'Delivery channel', 'Consent', 'Payment status', 'Payment updated (UTC)',
               'Delivery status', 'Owner notified', 'Notes', 'Raw (JSON)'];
var COL = {}; HEADERS.forEach(function (h, i) { COL[h] = i + 1; });

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

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}

function findRow(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) { return 0; }
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === id) { return i + 2; } }
  return 0;
}

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
  var status = emailOk ? sendConfirmationEmail(s) : 'EMAIL_NOT_SENT: invalid email';
  var row = [
    id, plain(d.ts, 40) || new Date().toISOString(), s.name, s.email, s.whatsapp, s.instagram,
    s.areas, s.types, s.deal, s.devs, s.comms, s.buyers, s.sellers, s.years, s.langs, s.goals,
    s.freq, s.insp[0] || '', s.insp[1] || '', s.insp[2] || '', s.channel, d.consent === true ? 'yes' : 'no',
    'PENDING', '', status, '', s.preview ? 'sent from a preview page' : '', ''
  ].map(safe);
  row[COL['Raw (JSON)'] - 1] = JSON.stringify(s);
  sh.appendRow(row);
  var r = sh.getLastRow();
  sh.getRange(r, COL['Owner notified']).setValue(sendOwnerEmail(
    'New Content Sprint signup: ' + s.name,
    ownerSummary(s) + '\n\nPayment status: PENDING (they have not said they paid yet).'));
  return { ok: true, status: status };
}

/* ---------- payment events ---------- */
function onPayment(sh, d, id, newStatus) {
  var r = findRow(sh, id);
  if (!r) {   // the signup row was not found: keep a trace anyway
    var row = []; for (var i = 0; i < HEADERS.length; i++) { row.push(''); }
    row[0] = id; row[1] = plain(d.ts, 40) || new Date().toISOString();
    row[COL['Notes'] - 1] = 'No signup row found for this ID';
    sh.appendRow(row); r = sh.getLastRow();
  }
  var cur = String(sh.getRange(r, COL['Payment status']).getValue());
  if (cur === 'PAID_CLAIMED' && newStatus === 'LINK_OPENED') { return { ok: true, status: cur }; }   // never step backwards
  if (cur === 'PAID_CLAIMED' && newStatus === 'PAID_CLAIMED') { return { ok: true, duplicate: true }; }
  sh.getRange(r, COL['Payment status']).setValue(newStatus);
  sh.getRange(r, COL['Payment updated (UTC)']).setValue(plain(d.ts, 40) || new Date().toISOString());
  if (d.preview === true) { sh.getRange(r, COL['Notes']).setValue('payment event from a preview page'); }
  if (newStatus === 'PAID_CLAIMED') {
    var s = rowToSignup(sh, r);
    if (!s.name) { s.name = 'unknown (ID ' + id + ')'; }
    var ok = /^[^\s@,;<>"]+@[^\s@,;<>"]+\.[^\s@,;<>"]{2,}$/.test(s.email);
    var st = ok ? sendWelcomeEmail(s) : 'EMAIL_NOT_SENT: invalid email';
    sh.getRange(r, COL['Delivery status']).setValue(st);
    sh.getRange(r, COL['Owner notified']).setValue(sendOwnerEmail(
      'Content Sprint: payment claimed by ' + s.name,
      s.name + ' says they have paid ' + CURRENCY + ' ' + PRICE + '.\n\nCheck Razorpay for the payment, then start their content at the next 9 AM Dubai time.\n\n' + ownerSummary(s)));
  }
  return { ok: true, status: newStatus };
}

function rowToSignup(sh, r) {
  var v = sh.getRange(r, 1, 1, HEADERS.length).getValues()[0];
  function g(h) { return String(v[COL[h] - 1] || ''); }
  return { id: g('Submission ID'), name: g('Name'), email: g('Email'), whatsapp: g('WhatsApp'),
           instagram: g('Instagram'), areas: g('Areas / specialization'), types: g('Property types'),
           deal: g('Deal focus'), devs: g('Developers'), comms: g('Communities'), buyers: g('Target buyer'),
           sellers: g('Target seller'), years: g('Years of experience'), langs: g('Languages'),
           goals: g('Brand goals'), freq: g('Posting frequency'),
           insp: [g('Inspiration 1'), g('Inspiration 2'), g('Inspiration 3')].filter(Boolean),
           channel: g('Delivery channel') };
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
    MailApp.sendEmail({ to: s.email, subject: 'Your Content Sprint spot request, ' + firstName(s.name),
      body: confirmationText(s), htmlBody: confirmationHtml(s), name: SENDER_NAME });
    return 'CONFIRMATION_SENT';
  } catch (err) { return 'EMAIL_FAILED: ' + String(err).slice(0, 120); }
}

function sendWelcomeEmail(s) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { return 'EMAIL_NOT_SENT: daily quota used up'; }
    MailApp.sendEmail({ to: s.email, subject: 'Welcome to the RealtorOne Content Sprint',
      body: welcomeText(s), htmlBody: welcomeHtml(s), name: SENDER_NAME });
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

function confirmationText(s) {
  var L = ['Hi ' + firstName(s.name) + ',', '',
    'Thanks for completing the RealtorOne Content Sprint onboarding. We have your details.', '',
    'WHAT WE NOTED',
    'Area: ' + s.areas, 'Property: ' + s.types, 'Focus: ' + s.deal, 'Content sent via: ' + s.channel, '',
    'NEXT STEP',
    'Secure your founding spot (' + SPOTS_TEXT() + ').', offerLine()];
  if (PAYMENT_URL) { L.push('Pay here: ' + PAYMENT_URL); } else { L.push('Go back to the page and tap "Secure my spot": ' + PAGE_URL); }
  L.push('', 'WHAT YOU GET', 'Every morning at 9 AM Dubai time: 2 static posts + 3 carousels, captions included.',
    '30 days x 5 posts/day = 150 pieces of content.', '',
    'Content keeps you visible. It does not guarantee leads or deals.');
  if (waUrl()) { L.push('', 'Questions? WhatsApp us: ' + waUrl()); }
  L.push('', '--', 'RealtorOne Content Sprint, a RealtorOne project');
  return L.join('\n');
}

function welcomeText(s) {
  var L = ['Hi ' + firstName(s.name) + ',', '', 'Welcome to the RealtorOne Content Sprint.', '',
    'We have noted your payment and will confirm it shortly. Your content starts at the next 9:00 AM Dubai time after we confirm, then every morning for 30 days.', '',
    'WHAT HAPPENS NEXT',
    '1. We confirm your payment and message you on ' + (s.channel || 'your chosen channel') + '.',
    '2. Next 9 AM Dubai time: your first day of ready-to-post content arrives.',
    '3. 30 days: 2 static posts and 3 carousels every morning, captions included.', '',
    '30 days x 5 posts/day = 150 opportunities to be seen.', '',
    'You don\'t know which post will start the conversation that becomes your next client.'];
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
  h += '<tr><td style="background:' + EM.charcoal + ';padding:30px 32px 28px;">' +
    '<div style="font-family:' + EM.sans + ';font-size:12px;line-height:18px;letter-spacing:4px;text-transform:uppercase;font-weight:600;color:' + EM.cream + ';">RealtorOne Content Sprint</div>' +
    '<div style="height:2px;width:36px;background:' + EM.accent + ';font-size:1px;line-height:2px;margin:14px 0 14px;">&nbsp;</div>' +
    '<div style="font-family:' + EM.serif + ';font-style:italic;font-size:17px;line-height:25px;color:#cfc8b8;">Never wonder what to post again.</div></td></tr>';
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

function confirmationHtml(s) {
  var rows = eRow('<p style="margin:0 0 6px;font-family:' + EM.serif + ';font-size:20px;line-height:28px;color:' + EM.charcoal + ';">Hi ' + esc(firstName(s.name)) + ',</p>' +
      ePara('Thanks for completing the Content Sprint onboarding. We have your details and will build your content around them.'), '32px 32px 8px') +
    eRow(eLabel('What we noted') + eDetails([['Area', s.areas], ['Property', s.types], ['Focus', s.deal], ['Content sent via', s.channel]]), '8px 32px 20px') +
    eRow(eLabel('Next step') +
      ePara('Secure your founding spot (' + SPOTS_TEXT() + ').') +
      '<p style="margin:0 0 12px;font-family:' + EM.sans + ';font-size:16px;line-height:25px;color:' + EM.ink + ';">Founding offer: you get it for <b>' + esc(money()) + '</b> / 30 days' + (regularMoney() ? ' <span style="color:' + EM.muted + ';">(regular price <span style="text-decoration:line-through;">' + esc(regularMoney()) + '</span>)</span>' : '') + '.</p>' +
      (PAYMENT_URL ? eButton(PAYMENT_URL, 'Secure my spot for ' + money()) : eButton(PAGE_URL, 'Back to the page')) +
      '<p style="margin:12px 0 0;font-family:' + EM.sans + ';font-size:14px;line-height:21px;color:' + EM.ink + ';">Every morning at 9 AM Dubai time: 2 static posts + 3 carousels, captions included. 30 days &times; 5 posts/day = 150 pieces of content.</p>' +
      '<p style="margin:10px 0 0;font-family:' + EM.sans + ';font-size:12px;line-height:18px;color:' + EM.muted + ';">Content keeps you visible. It does not guarantee leads or deals.</p>' + eWa(), '8px 32px 26px');
  return eShell('We have your Content Sprint details. One step left: secure your spot.', 'Your Content Sprint spot request', rows);
}

function welcomeHtml(s) {
  function step(n, t) {
    return '<tr><td width="40" valign="top" style="padding:0 0 12px;"><div style="width:28px;height:28px;border-radius:14px;background:' + EM.charcoal + ';color:' + EM.cream + ';font-family:' + EM.serif + ';font-size:15px;line-height:28px;text-align:center;">' + n + '</div></td>' +
      '<td valign="top" style="padding:3px 0 12px;font-family:' + EM.sans + ';font-size:15px;line-height:23px;color:' + EM.ink + ';">' + t + '</td></tr>';
  }
  var rows = eRow('<p style="margin:0 0 6px;font-family:' + EM.serif + ';font-size:20px;line-height:28px;color:' + EM.charcoal + ';">Hi ' + esc(firstName(s.name)) + ',</p>' +
      '<h1 style="margin:0 0 12px;font-family:' + EM.serif + ';font-weight:normal;font-size:32px;line-height:38px;color:' + EM.charcoal + ';">Welcome to the RealtorOne Content Sprint</h1>' +
      ePara('We have noted your payment and will confirm it shortly. Your content starts at the next 9:00 AM Dubai time after we confirm, then every morning for 30 days.'), '32px 32px 8px') +
    eRow(eLabel('What happens next') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' +
      step(1, '<b>We confirm your payment</b> and message you on ' + esc(s.channel || 'your chosen channel') + '.') +
      step(2, '<b>Next 9 AM Dubai time:</b> your first day of ready-to-post content arrives.') +
      step(3, '<b>30 days:</b> 2 static posts and 3 carousels every morning, captions included.') + '</table>', '8px 32px 8px') +
    eRow('<div style="border:1px solid ' + EM.accent + ';border-radius:12px;padding:16px;text-align:center;font-family:' + EM.serif + ';font-size:21px;line-height:28px;color:' + EM.charcoal + ';">30 days &times; 5 posts/day = 150 opportunities to be seen.</div>' +
      '<p style="margin:16px 0 0;font-family:' + EM.serif + ';font-style:italic;font-size:17px;line-height:25px;color:' + EM.ink + ';">You don&rsquo;t know which post will start the conversation that becomes your next client.</p>' + eWa(), '10px 32px 26px');
  return eShell('Welcome. Your 30 days of ready-to-post content are about to begin.', 'Welcome to the RealtorOne Content Sprint', rows);
}

/* ---------- helpers ---------- */
function plain(s, max) { return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').slice(0, max || 60).trim(); }
function list(a) { return (Array.isArray(a) ? a : [a]).filter(function (x) { return x; }).map(function (x) { return plain(x, 60); }).slice(0, 12).join(', '); }
// Stops spreadsheet formula injection (a value starting with = + - @ is stored as text).
function safe(v) {
  var s = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/* Optional: run from the editor to get both emails in your own inbox before going live. */
function testEmailsToMe() {
  var me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  var s = { id: 'TEST', name: 'Test Realtor', email: me, whatsapp: '+971500000000', instagram: 'test.realtor', areas: 'Dubai Marina',
            types: 'Apartments', deal: 'Both', devs: '', comms: '', buyers: 'Investors', sellers: 'Landlords', years: '3–5 years',
            langs: 'English', goals: 'Get more enquiries', freq: 'Daily', insp: ['@example'], channel: 'WhatsApp' };
  Logger.log(sendConfirmationEmail(s) + ' / ' + sendWelcomeEmail(s) + ' to ' + me);
}
