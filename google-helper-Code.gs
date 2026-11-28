/* =====================================================================
   Family Hub — Google helper (Apps Script)
   Lives inside the "Family Hub Data" Google Sheet on the family account.
   - Reads and writes the family Google Calendar
   - Stores members, chores, rewards, stars log and shared settings in this Sheet
   Family Hub talks to it through the Web app link + key.

   One-off setup (see "Family Hub setup notes.md"):
     1. Services (+) → Google Calendar API → Add
     2. Run  setup  once (approve the permissions) → copy the KEY from the log
     3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone

   Security:
   - Every request must carry the key. Lost a QR code / link?  Run  newKey  —
     every device then needs reconnecting with a fresh QR code.
   - The helper only ever touches ONE calendar: the one set in CALENDAR_ID below.
   ===================================================================== */

/* The only calendar Family Hub can see or change.
   Leave as '' for the family account's main calendar, or paste a calendar ID
   from the list that  setup  prints (e.g. the "Family" calendar).
   After changing it: Deploy → Manage deployments → ✏️ → New version → Deploy. */
var CALENDAR_ID = '';

var TABLES = ['members', 'chores', 'rewards', 'log'];

/* ---------- Run this once from the editor ---------- */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  props.setProperty('SSID', ss.getId());
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  if (!props.getProperty('REV')) props.setProperty('REV', String(Date.now()));
  TABLES.concat(['settings']).forEach(function (t) { sheet_(ss, t); });
  var first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' && ss.getSheets().length > 1) ss.deleteSheet(first);
  Logger.log('Calendars on this account (to use one other than the main one, paste its ID into CALENDAR_ID):');
  Calendar.CalendarList.list().items.forEach(function (c) {
    Logger.log('   ' + (c.summaryOverride || c.summary) + (c.primary ? '  (main)' : '') + '   →   ' + c.id);
  });
  Logger.log('Family Hub will use: ' + cal_());
  Logger.log('All set. Your Family Hub key is:  ' + props.getProperty('KEY'));
  Logger.log('Next: Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).');
}

/* Shows the key again if you need it */
function showKey() { Logger.log(PropertiesService.getScriptProperties().getProperty('KEY')); }

/* Cancels the old key (e.g. a QR code or link got out). All devices stop syncing
   until you reconnect one with the new key and scan a fresh QR code on the rest. */
function newKey() {
  var k = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  PropertiesService.getScriptProperties().setProperty('KEY', k);
  Logger.log('Old key cancelled. New key:  ' + k);
}

function cal_() { return CALENDAR_ID || 'primary'; }

/* Bumps the version when someone edits the Sheet by hand, so devices pick it up */
function onEdit(e) { try { bump_(); } catch (x) {} }

/* ---------- Web app ---------- */
function doGet(e) { return out_({ ok: true, app: 'Family Hub helper' }); }
function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'Bad request' }); }
  return out_(handle_(p));
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(p) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('KEY');
    if (!key) return { ok: false, error: 'Run setup in the script first' };
    if (p.k !== key) { Utilities.sleep(1000); return { ok: false, error: 'Wrong key', code: 'key' }; }
    var cal = cal_(); // fixed here — whatever calendar the app asks for is ignored
    switch (p.a) {
      case 'ping': return { ok: true, email: Session.getEffectiveUser().getEmail() };
      case 'cals': return { ok: true, items: cals_() };
      case 'cal.list': return { ok: true, items: calList_(cal, p.from, p.to) };
      case 'cal.get': return { ok: true, item: calGet_(cal, p.id) };
      case 'cal.save': return { ok: true, item: calSave_(cal, p.id, p.ev, p.keepRec) };
      case 'cal.del': calDel_(cal, p.id); return { ok: true };
      case 'sync': return { ok: true, data: sync_(p) };
    }
    return { ok: false, error: 'Unknown action' };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/* =================== Calendar =================== */
var TZ_ = null;
function tz_() { return TZ_ || (TZ_ = CalendarApp.getTimeZone() || Session.getScriptTimeZone()); }
function fmt_(d, f) { return Utilities.formatDate(d, tz_(), f); }
function addDay_(s, n) {
  var p = s.split('-'), d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.toISOString().slice(0, 10);
}

function cals_() {
  var c = Calendar.CalendarList.get(cal_());
  return [{ id: c.id, name: c.summaryOverride || c.summary, primary: !!c.primary }];
}

function map_(ev) {
  if (!ev || ev.status === 'cancelled') return null;
  var o = {
    id: ev.id, rid: ev.recurringEventId || '', title: ev.summary || '(No title)',
    notes: ev.description || '', loc: ev.location || '',
    ext: (ev.extendedProperties && ev.extendedProperties['private']) || {}
  };
  if (ev.start.date) {
    o.allDay = true; o.date = ev.start.date; o.start = ''; o.end = '';
    var last = addDay_(ev.end.date, -1); o.endDate = last > o.date ? last : '';
  } else {
    var s = new Date(ev.start.dateTime), en = new Date(ev.end.dateTime);
    o.allDay = false; o.date = fmt_(s, 'yyyy-MM-dd'); o.start = fmt_(s, 'HH:mm'); o.end = fmt_(en, 'HH:mm');
    var ed = fmt_(en, 'yyyy-MM-dd');
    o.endDate = ed > o.date ? ed : '';
    if (o.endDate === addDay_(o.date, 1) && o.end === '00:00') { o.endDate = ''; o.end = '23:59'; }
  }
  if (ev.recurrence) { var g = guess_(ev.recurrence); o.repeat = g.repeat; o.until = g.until; }
  return o;
}

function calList_(cal, from, to) {
  var out = [], tok = null;
  do {
    var r = Calendar.Events.list(cal, { timeMin: from, timeMax: to, singleEvents: true, orderBy: 'startTime', maxResults: 2500, pageToken: tok || undefined });
    (r.items || []).forEach(function (ev) { var m = map_(ev); if (m) out.push(m); });
    tok = r.nextPageToken;
  } while (tok && out.length < 5000);
  return out;
}

function calGet_(cal, id) { return map_(Calendar.Events.get(cal, id)); }

function calDel_(cal, id) {
  try { Calendar.Events.remove(cal, id); }
  catch (e) { if (!/deleted|not found|410|404/i.test(String(e))) throw e; }
}

/* Works out which Family Hub "Repeats" option a Google repeat rule matches */
function guess_(rec) {
  var r = '';
  (rec || []).forEach(function (x) { if (x.indexOf('RRULE:') === 0) r = x; });
  if (!r) return { repeat: 'none', until: '' };
  var f = (r.match(/FREQ=(\w+)/) || [])[1], i = +((r.match(/INTERVAL=(\d+)/) || [])[1] || 1),
      bd = (r.match(/BYDAY=([^;]+)/) || [])[1] || '', u = (r.match(/UNTIL=(\d{8})/) || [])[1],
      odd = /COUNT=|BYSETPOS|BYMONTH=/.test(r) || rec.length > 1;
  var rep = 'custom';
  if (!odd) {
    if (f === 'DAILY' && i === 1) rep = 'daily';
    else if (f === 'WEEKLY' && bd === 'MO,TU,WE,TH,FR' && i === 1) rep = 'weekdays';
    else if (f === 'WEEKLY' && bd.indexOf(',') < 0 && i === 1) rep = 'weekly';
    else if (f === 'WEEKLY' && bd.indexOf(',') < 0 && i === 2) rep = 'fortnightly';
    else if (f === 'MONTHLY' && !bd && i === 1) rep = 'monthly';
    else if (f === 'YEARLY' && i === 1) rep = 'yearly';
  }
  return { repeat: rep, until: u ? u.slice(0, 4) + '-' + u.slice(4, 6) + '-' + u.slice(6, 8) : '' };
}

function rrule_(ev) {
  var m = { daily: 'FREQ=DAILY', weekdays: 'FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR', weekly: 'FREQ=WEEKLY',
            fortnightly: 'FREQ=WEEKLY;INTERVAL=2', monthly: 'FREQ=MONTHLY', yearly: 'FREQ=YEARLY' }[ev.repeat];
  if (!m) return '';
  if (ev.until) {
    if (ev.allDay) m += ';UNTIL=' + ev.until.replace(/-/g, '');
    else {
      var d = Utilities.parseDate(ev.until + ' 23:59:59', tz_(), 'yyyy-MM-dd HH:mm:ss');
      m += ';UNTIL=' + Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
    }
  }
  return 'RRULE:' + m;
}

function plusHour_(t) {
  var h = +t.slice(0, 2), m = t.slice(3, 5);
  return h >= 23 ? '23:59' : ((h + 1 < 10 ? '0' : '') + (h + 1) + ':' + m);
}

function calSave_(cal, id, ev, keepRec) {
  if (!ev || !ev.title || !ev.date) throw new Error('Missing title or date');
  if (ev.bday) { ev.allDay = true; ev.repeat = 'yearly'; ev.endDate = ''; }
  var body = {
    summary: ev.title, description: ev.notes || '',
    extendedProperties: { 'private': {
      fhWho: (ev.who || []).join(','), fhBday: ev.bday ? '1' : '', fhByear: ev.byear ? String(ev.byear) : ''
    } }
  };
  var last = ev.endDate && ev.endDate > ev.date ? ev.endDate : ev.date;
  if (ev.allDay) {
    body.start = { date: ev.date, dateTime: null };
    body.end = { date: addDay_(last, 1), dateTime: null };
  } else {
    var st = ev.start || '09:00', en = ev.end;
    if (!en || (last === ev.date && en <= st)) en = plusHour_(st);
    body.start = { dateTime: ev.date + 'T' + st + ':00', timeZone: tz_(), date: null };
    body.end = { dateTime: last + 'T' + en + ':00', timeZone: tz_(), date: null };
  }
  if (!keepRec) { var r = rrule_(ev); body.recurrence = r ? [r] : []; }
  var res = id ? Calendar.Events.patch(body, cal, id) : Calendar.Events.insert(body, cal);
  return map_(res);
}

/* =================== Chores, stars & rewards (this Sheet) =================== */
function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SSID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function cap_(t) { return t.charAt(0).toUpperCase() + t.slice(1); }
function sheet_(ss, t) {
  var sh = ss.getSheetByName(cap_(t));
  if (!sh) { sh = ss.insertSheet(cap_(t)); sh.setFrozenRows(1); }
  return sh;
}
function enc_(v) { return v === undefined || v === '' ? '' : JSON.stringify(v); }
function dec_(v) {
  if (typeof v !== 'string') return v;           // numbers / true / false typed straight into the Sheet
  if (v === '') return undefined;
  try { return JSON.parse(v); } catch (e) { return v; } // plain text typed into the Sheet
}
function bump_() { var r = String(Date.now()); PropertiesService.getScriptProperties().setProperty('REV', r); return r; }

function readTable_(ss, t) {
  var v = sheet_(ss, t).getDataRange().getValues();
  if (v.length < 2) return [];
  var head = v[0], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = {}, any = false;
    for (var j = 0; j < head.length; j++) {
      if (!head[j]) continue;
      var x = dec_(v[i][j]);
      if (x !== undefined) { o[head[j]] = x; any = true; }
    }
    if (any && o.id !== undefined && o.id !== '') out.push(o);
  }
  return out;
}

function writeTable_(ss, t, rows) {
  var sh = sheet_(ss, t), keys = ['id'];
  rows.forEach(function (r) { for (var k in r) if (keys.indexOf(k) < 0) keys.push(k); });
  sh.clearContents();
  var data = [keys].concat(rows.map(function (r) { return keys.map(function (k) { return enc_(r[k]); }); }));
  sh.getRange(1, 1, data.length, keys.length).setValues(data);
}

/* Adds and removes stars-log rows without rewriting the whole log */
function logOps_(ss, add, del) {
  var sh = sheet_(ss, 'log'), v = sh.getDataRange().getValues();
  if (v.length < 1 || !v[0][0]) { writeTable_(ss, 'log', add); return; }
  var head = v[0].slice(), idCol = head.indexOf('id'), have = {}, gone = {};
  (del || []).forEach(function (id) { gone[id] = 1; });
  for (var i = v.length - 1; i >= 1; i--) {
    var id = dec_(v[i][idCol]);
    if (gone[id]) sh.deleteRow(i + 1); else have[id] = 1;
  }
  var rows = (add || []).filter(function (l) { return !have[l.id] && !gone[l.id]; });
  if (!rows.length) return;
  rows.forEach(function (r) { for (var k in r) if (head.indexOf(k) < 0) head.push(k); });
  if (head.length > v[0].length) sh.getRange(1, 1, 1, head.length).setValues([head]);
  var data = rows.map(function (r) { return head.map(function (k) { return enc_(r[k]); }); });
  sh.getRange(sh.getLastRow() + 1, 1, data.length, head.length).setValues(data);
}

function readKV_(ss) {
  var v = sheet_(ss, 'settings').getDataRange().getValues(), o = {};
  for (var i = 1; i < v.length; i++) if (v[i][0]) { var x = dec_(v[i][1]); if (x !== undefined) o[v[i][0]] = x; }
  return o;
}
function writeKV_(ss, obj) {
  var cur = readKV_(ss);
  for (var k in obj) cur[k] = obj[k];
  var data = [['key', 'value']];
  for (var k2 in cur) data.push([k2, enc_(cur[k2])]);
  var sh = sheet_(ss, 'settings');
  sh.clearContents();
  sh.getRange(1, 1, data.length, 2).setValues(data);
}

function sync_(p) {
  var props = PropertiesService.getScriptProperties(), ch = p.changes || {};
  var hasCh = false;
  for (var k in ch) { hasCh = true; break; }
  if (!hasCh && p.rev && p.rev === props.getProperty('REV')) return { same: true, rev: p.rev };

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var ss = ss_();
    if (hasCh) {
      ['members', 'chores', 'rewards'].forEach(function (t) { if (ch[t]) writeTable_(ss, t, ch[t]); });
      if (ch.settings) writeKV_(ss, ch.settings);
      if (ch.logReplace) writeTable_(ss, 'log', ch.logReplace);
      else if ((ch.logAdd && ch.logAdd.length) || (ch.logDel && ch.logDel.length)) logOps_(ss, ch.logAdd, ch.logDel);
      bump_();
    }
    var d = { rev: props.getProperty('REV'), settings: readKV_(ss) };
    TABLES.forEach(function (t) { d[t] = readTable_(ss, t); });
    d.empty = !d.members.length && !d.chores.length && !d.log.length;
    return d;
  } finally {
    lock.releaseLock();
  }
}
