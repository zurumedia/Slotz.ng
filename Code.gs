/* SLOTZ ticket server (Google Apps Script)
   Stores tickets in the Google Sheet it is attached to, confirms payments with Paystack,
   and checks tickets at the gate. Setup steps are in SETUP.md. */

/* ===== EDIT THIS WHEN YOU ADD AN EVENT. Names and prices must match the website exactly. ===== */
const EVENTS = {
  "Slotz Test Night": { fee: 0,  types: { "Regular": { p: 200 }, "VIP": { p: 500 } } },
  "Mara Group":       { fee: 10, types: { "Regular": { p: 20 },  "VIP": { p: 50 } } }
  /* Group tickets look like: "Group of 5": { p: 6000, g: 5 }  (g = people admitted) */
};
const WHATSAPP = "+234 810 240 7071";
/* ===== END OF EDIT BLOCK ===== */

const HEAD = ['Time','Ticket no','Code','Reference','Event','Ticket','Admits','Name','Phone','Email','Amount paid','Status','Used at'];
const ALPHA = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';

function props_() { return PropertiesService.getScriptProperties(); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('Tickets');
  if (!s) {
    s = ss.insertSheet('Tickets');
    s.appendRow(HEAD);
    s.setFrozenRows(1);
    s.getRange('I:I').setNumberFormat('@');
  }
  return s;
}

function setup() { sheet_(); }

function doGet() { return out_({ ok: true, service: 'slotz' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'issue') return out_(issue_(String(b.reference || '')));
    if (b.action === 'scan') return out_(scan_(String(b.code || ''), String(b.key || '')));
    return out_({ ok: false, error: 'Unknown request.' });
  } catch (err) {
    return out_({ ok: false, error: 'Server error. Try again.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function verify_(ref) {
  const r = UrlFetchApp.fetch('https://api.paystack.co/transaction/verify/' + encodeURIComponent(ref), {
    headers: { Authorization: 'Bearer ' + props_().getProperty('PAYSTACK_SECRET_KEY') },
    muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) return null;
  const d = JSON.parse(r.getContentText());
  return (d.status && d.data && d.data.status === 'success' && d.data.currency === 'NGN') ? d.data : null;
}

function pick_(len, rows, col) {
  const used = {};
  for (let i = 1; i < rows.length; i++) used[String(rows[i][col])] = 1;
  for (;;) {
    let s = '';
    for (let i = 0; i < len; i++) s += ALPHA.charAt(Math.floor(Math.random() * ALPHA.length));
    if (!used[s]) return s;
  }
}

function pub_(r) { return { no: r[1], code: r[2], event: r[4], type: r[5], admits: r[6], name: r[7] }; }

function issue_(ref) {
  if (!/^[A-Za-z0-9_-]{6,64}$/.test(ref)) return { ok: false, error: 'Invalid payment reference.' };
  const s = sheet_(), rows = s.getDataRange().getValues();
  let at = -1;
  for (let i = 1; i < rows.length; i++) if (String(rows[i][3]) === ref) { at = i; break; }
  if (at > 0 && rows[at][11] === 'VALID') return { ok: true, ticket: pub_(rows[at]) };

  const tx = verify_(ref);
  if (!tx) return { ok: false, error: 'Payment not confirmed yet.' };

  const f = {};
  ((tx.metadata && tx.metadata.custom_fields) || []).forEach(function (c) { f[c.variable_name] = String(c.value); });
  const name = (f.name || '').slice(0, 80), phone = (f.phone || '').slice(0, 20);
  const event = f.event || '', type = f.ticket || '', admits = parseInt(f.admits, 10);
  const ev = EVENTS[event], t = ev && ev.types[type];

  let status = 'REVIEW';
  if (ev && t && admits >= 1 && admits <= 10 && (!t.g || admits === t.g)) {
    const expected = t.g ? t.p + ev.fee * t.g : (t.p + ev.fee) * admits;
    if (Math.round(expected * 100) === tx.amount) status = 'VALID';
  }

  const no = at > 0 ? rows[at][1] : 'SLZ-' + pick_(6, rows, 1);
  const code = status === 'VALID' ? pick_(16, rows, 2) : '';
  const row = [new Date(), no, code, ref, event, type, admits || '', name, phone,
               (tx.customer && tx.customer.email) || '', tx.amount / 100, status, ''];
  if (at > 0) s.getRange(at + 1, 1, 1, 13).setValues([row]); else s.appendRow(row);

  if (status !== 'VALID')
    return { ok: false, error: 'Your payment was received, but your ticket needs a quick check by Slotz. Message us on WhatsApp ' + WHATSAPP + '.' };
  return { ok: true, ticket: pub_(row) };
}

function scan_(q, key) {
  const gk = props_().getProperty('GATE_KEY');
  if (!gk || key !== gk) return { ok: true, result: 'denied' };
  q = q.trim().toUpperCase();
  if (!q) return { ok: true, result: 'invalid' };
  const s = sheet_(), rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const hit = (r[2] !== '' && String(r[2]).toUpperCase() === q) || String(r[1]).toUpperCase() === q;
    if (!hit) continue;
    if (r[11] !== 'VALID') return { ok: true, result: 'invalid' };
    const d = { no: r[1], event: r[4], type: r[5], admits: r[6], name: r[7],
                phone: String(r[8]), email: r[9], amount: r[10] };
    if (r[12]) {
      return { ok: true, result: 'used', ticket: d,
               at: Utilities.formatDate(new Date(r[12]), Session.getScriptTimeZone(), 'd MMM yyyy, h:mm a') };
    }
    s.getRange(i + 1, 13).setValue(new Date());
    return { ok: true, result: 'valid', ticket: d };
  }
  return { ok: true, result: 'invalid' };
}
