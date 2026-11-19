/**
 * INVO BUILD – Google Sheet backend (Apps Script) with login / signup
 *
 * DO ALAG GOOGLE SHEETS:
 *   A) CONTROL sheet  (ye wali jis mein aap script paste karte hain): sirf "Users" tab.
 *      Yahan aap clients ko approve (Active = TRUE), block (FALSE) aur Expiry set karte hain.
 *   B) DATA sheet     ("INVO BUILD - Data", setup khud banata hai): "Invoices" aur "Settings" tabs.
 *      Clients ki invoices aur karobar ki details yahan save hoti hain.
 *
 * Setup:
 * 1) Nayi Google Sheet banayen (ye CONTROL sheet hogi) > Extensions > Apps Script > ye code paste karein
 * 2) Function "setup" ek baar run karein (permissions allow karein). DATA sheet khud ban jati hai,
 *    uska link Users sheet mein I2 par likha milta hai.
 * 3) Deploy > New deployment > Web app > Execute as: Me | Who has access: Anyone > Deploy
 * 4) /exec URL index.html ke top par API_URL mein paste karein (sirf aap, ek dafa)
 * 5) App mein apna account signup karein, phir Users sheet mein apni row ka Active = TRUE kar dein
 */
const USERS_H = ['Email', 'Name', 'Hash', 'Salt', 'Active', 'Expiry', 'Created'];
const INV_H = ['ID', 'Key', 'Type', 'InvNo', 'Date', 'Buyer', 'Value', 'Tax', 'Total', 'Updated', 'JSON'];
const SET_H = ['Key', 'JSON', 'Updated'];

function head(s, h) { s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold'); s.setFrozenRows(1); }

function setup() {
  const ctrl = SpreadsheetApp.getActive(), p = PropertiesService.getScriptProperties();
  head(ctrl.getSheetByName('Users') || ctrl.insertSheet('Users'), USERS_H);
  let id = p.getProperty('DATA_ID'), d;
  if (id) d = SpreadsheetApp.openById(id);
  else {
    d = SpreadsheetApp.create('INVO BUILD - Data');
    p.setProperty('DATA_ID', d.getId());
    d.getSheets()[0].setName('Invoices');
    d.insertSheet('Settings');
  }
  head(d.getSheetByName('Invoices'), INV_H);
  d.getSheetByName('Invoices').getRange('D:E').setNumberFormat('@');
  head(d.getSheetByName('Settings'), SET_H);
  ctrl.getSheetByName('Users').getRange('I1:I2').setValues([['Data sheet:'], [d.getUrl()]]);
  sec();
  Logger.log('DATA sheet: ' + d.getUrl());
}

function sec() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('SECRET');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', k); }
  return k;
}
const hmac = s => Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, sec()));
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const users = () => SpreadsheetApp.getActive().getSheetByName('Users').getDataRange().getValues();
const dataSS = () => SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DATA_ID'));

function active(u) {
  if (String(u[4]).toUpperCase() !== 'TRUE')
    throw Error('Aapka account abhi approve nahi hua ya band hai. Meherbani karke humse rabta karein');
  if (u[5] && new Date(u[5]) < new Date(new Date().setHours(0, 0, 0, 0)))
    throw Error('Aapki subscription khatam ho gayi hai, meherbani karke renew karein');
}

function signup(q) {
  const em = String(q.email || '').trim().toLowerCase(), pw = String(q.pass || '');
  const nm = String(q.name || '').trim().replace(/^[=+\-@\s]+/, '').slice(0, 60);
  if (!/^[^@\s|]+@[^@\s|]+\.[^@\s|]+$/.test(em)) throw Error('Sahi email likhein');
  if (pw.length < 6) throw Error('Password kam az kam 6 huroof ka rakhein');
  if (!nm) throw Error('Apna naam likhein');
  if (users().slice(1).some(r => String(r[0]).toLowerCase() === em)) throw Error('Ye email pehle se registered hai');
  const salt = Utilities.getUuid();
  SpreadsheetApp.getActive().getSheetByName('Users').appendRow([em, nm, hmac(salt + '|' + pw), salt, false, '', new Date()]);
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'INVO BUILD: naya signup',
      em + ' (' + nm + ') ne account banaya hai. Approve karne ke liye Users sheet mein Active ko TRUE karein.');
  } catch (e) {}
  return { ok: 1, msg: 'Account ban gaya hai. Approve hone ka intezar karein, phir login karein.' };
}

function login(q) {
  const em = String(q.email || '').trim().toLowerCase();
  const u = users().slice(1).find(r => String(r[0]).toLowerCase() === em);
  if (!u || u[2] !== hmac(u[3] + '|' + q.pass)) throw Error('Email ya password galat hai');
  active(u);
  const exp = Date.now() + 30 * 864e5; // 30 din ka session
  return { email: em, name: u[1], token: em + '|' + exp + '|' + hmac(em + '|' + exp) };
}

function auth(q) {
  const t = String(q.token || '').split('|');
  if (t.length !== 3 || t[2] !== hmac(t[0] + '|' + t[1]) || Date.now() > +t[1])
    throw Error('LOGIN:Session khatam ho gayi, dobara login karein');
  const u = users().slice(1).find(r => String(r[0]).toLowerCase() === t[0]);
  if (!u) throw Error('LOGIN:Account nahi mila');
  active(u);
  return u;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const q = JSON.parse(e.postData.contents);
    if (q.action === 'signup') return out(signup(q));
    if (q.action === 'login') return out(login(q));
    const u = auth(q), key = String(u[0]).toLowerCase();
    if (q.action === 'ping') return out({ ok: 1, name: u[1] });

    const ds = dataSS(), sh = ds.getSheetByName('Invoices'), st = ds.getSheetByName('Settings');
    const d = sh.getDataRange().getValues(), sd = st.getDataRange().getValues();
    const mine = r => String(r[1]).toLowerCase() === key;
    const at = id => d.findIndex((r, n) => n && r[0] === id && mine(r)) + 1; // 1-based, 0 = not found

    if (q.action === 'list') {
      const s = sd.slice(1).find(r => String(r[0]).toLowerCase() === key);
      return out({ rows: d.slice(1).filter(mine).map(r => JSON.parse(r[10])), set: s ? JSON.parse(s[1]) : null });
    }
    if (q.action === 'saveset') {
      const row = [key, JSON.stringify(q.set), new Date()];
      const n = sd.findIndex((r, i) => i && String(r[0]).toLowerCase() === key) + 1;
      n ? st.getRange(n, 1, 1, 3).setValues([row]) : st.appendRow(row);
      return out({ ok: 1 });
    }
    if (q.action === 'save') {
      const i = q.inv, t = i.totals;
      const row = [i.id, key, i.type, i.no, i.date, i.buyer.name, t.v, t.g, t.t, new Date(), JSON.stringify(i)];
      const n = at(i.id);
      n ? sh.getRange(n, 1, 1, row.length).setValues([row]) : sh.appendRow(row);
      return out({ ok: 1 });
    }
    if (q.action === 'del') {
      const n = at(q.id);
      if (n) sh.deleteRow(n);
      return out({ ok: 1 });
    }
    throw Error('Unknown action');
  } catch (err) {
    return out({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}
