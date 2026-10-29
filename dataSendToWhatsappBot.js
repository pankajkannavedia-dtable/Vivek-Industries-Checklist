
/*** WHATSAPP RESPONSE SYNC — Whatsapp_Response tab -> ActLog (notify PC) + Attachments (proof) ***/
const WA_SHEET = 'Whatsapp_Response';
// 0-based column positions in Whatsapp_Response (row 1 = headers, data from row 2)
const WA_COL = { doer: 1, uid: 2, btn: 10, replyTime: 11, file: 12, msgId: 13, proofTime: 14 };

// Adds any missing header cells (e.g. waMsgId) to an existing sheet, at the SCHEMA position
function _ensureSchemaCols(name) {
  const sh = _sheet(name), want = SCHEMA[name];
  const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  for (let i = 0; i < want.length; i++) {
    if (!have[i]) sh.getRange(1, i + 1).setValue(want[i]);
  }
}

// Parses dd/MM/yyyy [HH:mm[:ss]] strings or Date objects; returns Date or null
function _waDate(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v || '').trim(); if (!s) return null;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!m) { const d = new Date(s); return isNaN(d.getTime()) ? null : d; }
  return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
}

// Appends many records to a sheet in one write, with sequential readable ids (AL001, AT001...)
function _appendObjs(sheetName, objs) {
  if (!objs.length) return;
  const hdr = SCHEMA[sheetName], sh = _sheet(sheetName);
  const m = /^([A-Za-z]+)(\d+)$/.exec(_nextId(sheetName));
  let n = parseInt(m[2], 10);
  const out = objs.map(function (o) {
    const rec = Object.assign({}, o, { id: m[1] + String(n++).padStart(m[2].length, '0'), clientId: CLAP_CLIENT_ID });
    return hdr.map(function (h) { return rec[h] !== undefined ? rec[h] : ''; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, out.length, hdr.length).setValues(out);
}

function syncWhatsappResponses() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 'busy';
  try {
    const src = _ss().getSheetByName(WA_SHEET);
    if (!src) throw new Error('Sheet "' + WA_SHEET + '" not found');
    const vals = src.getDataRange().getValues();
    if (vals.length < 2) return 'no data';

    _ensureSchemaCols('ActLog'); _ensureSchemaCols('Attachments');

    const taskById = {}; _readAll('Tasks').forEach(function (t) { taskById[String(t.id)] = t; });
    const seen = {};
    _readAll('ActLog').forEach(function (a) { if (a.waMsgId) seen[String(a.taskId) + '|' + String(a.waMsgId)] = true; });
    _readAll('Attachments').forEach(function (a) { if (a.waMsgId) seen[String(a.parentId) + '|' + String(a.waMsgId)] = true; });

    const logs = [], atts = [];
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      const uid = String(r[WA_COL.uid] || '').trim(), msgId = String(r[WA_COL.msgId] || '').trim();
      if (!uid || !msgId) continue;
      const task = taskById[uid]; if (!task) continue;

      const doer = String(r[WA_COL.doer] || task.doerName || '').trim();
      const btn = String(r[WA_COL.btn] || '').trim().toLowerCase();
      const replyDt = _waDate(r[WA_COL.replyTime]);
      const proofDt = (r[WA_COL.proofTime] instanceof Date) ? r[WA_COL.proofTime] : null;
      const evtIso = (replyDt || proofDt || new Date()).toISOString();

      // 1) Doer said "Yes" -> notify PC/Admin (one notification per task message). Task stays pending.
      if (btn === 'yes' && task.status !== 'done' && task.status !== 'missed') {
        const key = msgId + '|Y';
        if (!seen[uid + '|' + key]) {
          seen[uid + '|' + key] = true;
          logs.push({
            type: 'task_request',
            msg: doer + ' completed "' + task.name + '" — please verify and mark done',
            taskId: uid, taskName: task.name,
            requestedBy: doer + ' (WhatsApp)', requestedAt: (replyDt || new Date()).toISOString(),
            resolved: '', ts: (replyDt || new Date()).toISOString(), waMsgId: key
          });
        }
      }

      // 2) Proof file(s) -> Attachments (works with or without a Yes)
      String(r[WA_COL.file] || '').split(/\s+/).filter(function (u) { return /^https?:\/\//i.test(u); }).forEach(function (url) {
        const fm = url.match(/\/d\/([\w-]+)/);
        const key = msgId + '|' + (fm ? fm[1] : url);
        if (seen[uid + '|' + key]) return;
        seen[uid + '|' + key] = true;
        atts.push({
          parentId: uid, name: 'WhatsApp Proof', url: url,
          uploadedBy: doer + ' (WhatsApp)', createdAt: evtIso, waMsgId: key
        });
      });
    }

    _appendObjs('ActLog', logs);
    _appendObjs('Attachments', atts);
    Logger.log('WhatsApp sync: ' + logs.length + ' notification(s), ' + atts.length + ' attachment(s)');
    return { notifications: logs.length, attachments: atts.length };
  } finally {
    lock.releaseLock();
  }
}

/*** RUN ONCE — installs the 5-minute WhatsApp response sync ***/
function installWhatsappSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncWhatsappResponses') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncWhatsappResponses').timeBased().everyMinutes(5).create();
  Logger.log('WhatsApp sync trigger installed (every 5 min).');
}

/*** DIAGNOSTIC — run from editor, read the Execution log: why is each Whatsapp_Response row (not) syncing? ***/
function debugWhatsappSync() {
  const vals = _ss().getSheetByName(WA_SHEET).getDataRange().getValues();
  const taskById = {}; _readAll('Tasks').forEach(function (t) { taskById[String(t.id)] = t; });
  const seen = {};
  _readAll('ActLog').forEach(function (a) { if (a.waMsgId) seen[String(a.taskId) + '|' + String(a.waMsgId)] = true; });
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    const uid = String(r[WA_COL.uid] || '').trim(), msgId = String(r[WA_COL.msgId] || '').trim();
    const btn = String(r[WA_COL.btn] || '').trim();
    const t = taskById[uid];
    let why;
    if (!uid) why = 'SKIP: Unique ID empty';
    else if (!msgId) why = 'SKIP: Msg ID (col N) is empty';
    else if (!t) why = 'SKIP: Task ' + uid + ' not found in Tasks sheet';
    else if (btn.toLowerCase() !== 'yes') why = 'No Yes notification: button reply is "' + btn + '"';
    else if (t.status === 'done' || t.status === 'missed') why = 'SKIP: task status is "' + t.status + '" (priority ' + t.priority + ', planned ' + _dstr(t.plannedDate) + ')';
    else if (seen[uid + '|' + msgId + '|Y']) why = 'Already synced earlier';
    else why = 'OK: will create notification';
    Logger.log('Row ' + (i + 1) + ' | ' + uid + ' | msgId=' + (msgId || '(empty)') + ' | ' + why);
  }
}


function sendDataToWhatappSheet() {
  var srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  var destSS = SpreadsheetApp.openById('1oqohY_LDkgTSJEl8va9VyPf5X7oo39A88rUjdWEUZJ8');
  var destSheet = destSS.getSheetByName('Send Messages');
  var tz = Session.getScriptTimeZone();

  var colA = destSheet.getRange('A1:A' + destSheet.getMaxRows()).getValues();
  var lastRow = 1;

  for (var k = colA.length - 1; k >= 0; k--) {
    if (colA[k][0] !== '') {
      lastRow = k + 1;
      break;
    }
  }

  var existing = {};

  if (lastRow > 1) {
    var old = destSheet.getRange(2, 1, lastRow - 1, 5).getValues();

    old.forEach(function (o) {
      existing[o[2] + '|' + normDT(o[4], tz)] = true;
    });
  }

  var data = srcSheet.getDataRange().getValues();
  var outMain = [];
  var outQR = [];
  var outP = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];

    if (!r[0]) continue;
    if (String(r[12]).trim().toLowerCase() !== 'pending') continue;
    if (String(r[29]).trim() === '') continue; // Skip if AD (mobile number) is empty

    var dt = combineDateTime(r[9], r[10]);
    var key = r[0] + '|' + normDT(dt, tz);

    if (existing[key]) continue;
    existing[key] = true;

    outMain.push([r[29], r[6], r[0], r[3], dt, r[4], r[7]]);
    outP.push([r[8]]);           // Source column I -> Destination column P
    outQR.push([r[12], r[26]]);  // Source columns M and AA -> Destination columns Q and R
  }

  if (outMain.length === 0) return;

  var startRow = lastRow + 1;

  destSheet.getRange(startRow, 1, outMain.length, 7).setValues(outMain);
  destSheet.getRange(startRow, 16, outP.length, 1).setValues(outP);
  destSheet.getRange(startRow, 17, outQR.length, 2).setValues(outQR);
}


function normDT(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm:ss');
  return String(v).trim();
}

function combineDateTime(d, t) {
  if (!(d instanceof Date)) return d + ' ' + t;
  var dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var timeStr = (t instanceof Date)
    ? Utilities.formatDate(t, Session.getScriptTimeZone(), 'HH:mm:ss')
    : t;
  return dateStr + ' ' + timeStr;
}
