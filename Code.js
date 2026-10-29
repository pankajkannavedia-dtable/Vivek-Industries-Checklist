/*** CONFIG ***/
const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const DRIVE_FOLDER_ID = '1qdpXf5OjGukGPyv8J1Hd_eoOcW3oRBo-'; // for logo/photo/attachments

/*** CENTRALISED AUTH SHEET — shared doer list + roles across all systems ***/
const AUTH_SS_ID = '1i6HBTqGBdfNYPJUTQdYXxwOYnpq-9rNzzTR-bSHo2Kk';
const CLAP_SYSTEM_ID = 'CLAP';
const PRIORITY_WEIGHT = { H: 1.0, M: 0.5, L: 0.25 }; // impact on score if missed — High hurts most, Low least
const CLAP_CLIENT_ID = 'VIVEK'; // single-tenant now — was dynamic per-client before

function _authSS() { return SpreadsheetApp.openById(AUTH_SS_ID); }
function _authSheet(name) { return _authSS().getSheetByName(name); }
function _authReadAll(name) {
  const sh = _authSheet(name);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const hdr = vals[0], out = [];
  for (let i = 1; i < vals.length; i++) {
    const o = { _row: i + 1 };
    hdr.forEach((h, j) => { if (h) o[h] = vals[i][j]; });
    out.push(o);
  }
  return out;
}

/*** Build the doer list for CLAP from AUTH_USERS + AUTH_USER_ACCESS + AUTH_ROLES ***/
function _readDoersFromAuth() {
  const users = _authReadAll('AUTH_USERS');
  const access = _authReadAll('AUTH_USER_ACCESS').filter(a =>
    a.SYSTEM_ID === CLAP_SYSTEM_ID && String(a.STATUS).toUpperCase() === 'ACTIVE');
  const roles = _authReadAll('AUTH_ROLES');
  const roleById = {}; roles.forEach(r => { if (r.ROLE_ID) roleById[r.ROLE_ID] = r; });

  return access.map(function (a) {
    const u = users.find(x => String(x.USER_ID) === String(a.USER_ID));
    if (!u) return null;
    const role = roleById[a.ROLE_ID];
    const roleCode = role ? String(role.ROLE_CODE || '').toLowerCase() : 'staff'; // admin/pc/staff
    return {
      id: u.USER_ID,
      clientId: CLAP_CLIENT_ID,
      name: u.FULL_NAME || '',
      dept: u.DEPARTMENT || '',
      designation: u.DESIGNATION || '',
      email: u.EMAIL || '',
      username: u.EMAIL || '',
      phone: u.MOBILE || '',
      role: roleCode,
      isCoordinator: (roleCode === 'pc' || roleCode === 'admin'),
      active: String(u.STATUS).toUpperCase() === 'ACTIVE',
      dob: '',
      photoUrl: u.AVATAR_URL || '',
      createdAt: u.CREATED_ON || ''
    };
  }).filter(Boolean);
}

const SCHEMA = {
  Clients: ['id', 'companyName', 'adminName', 'adminEmail', 'adminPassword', 'phone', 'plan', 'accessExpiry', 'isLocked', 'welcomeMsg', 'createdAt'],
  Employees: ['id', 'clientId', 'name', 'dept', 'designation', 'email', 'username', 'password', 'role', 'isCoordinator', 'active', 'dob', 'photoUrl', 'createdAt'],
  Templates: ['id', 'clientId', 'name', 'dept', 'doerId', 'doerName', 'freq', 'priority', 'plannedTime', 'scheduleConfig', 'photoRequired', 'skipHolidays', 'isActive', 'createdAt', 'rotationEnabled', 'rotationDoer1Id', 'rotationDoer1Name', 'rotationDoer2Id', 'rotationDoer2Name', 'rotationDays'],
  Tasks: ['id', 'clientId', 'templateId', 'name', 'dept', 'doerId', 'doerName', 'freq', 'priority', 'plannedDate', 'plannedTime', 'originDate', 'status', 'doneBy', 'actualTime', 'completedAt', 'notes', 'photoRequired', 'proofPhotoUrl', 'isActive', 'requiresApproval', 'approvalStatus', 'swapId', 'originalDoerId', 'originalDoerName', 'lockedNotDone', 'createdAt'],
  Holidays: ['id', 'clientId', 'name', 'date', 'type', 'shiftDir', 'preloaded'],
  Settings: ['id', 'clientId', 'docId', 'data'],
  ActLog: ['id', 'clientId', 'type', 'msg', 'taskId', 'taskName', 'requestedBy', 'requestedAt', 'resolved', 'ts', 'waMsgId'],
  Swaps: ['id', 'clientId', 'fromDoerId', 'fromDoerName', 'toDoerId', 'toDoerName', 'startDate', 'endDate', 'reason', 'status', 'createdBy', 'createdAt'],
  AuditLog: ['id', 'clientId', 'action', 'type', 'by', 'ts'],
  Announcements: ['id', 'clientId', 'msg', 'fromAutomation', 'createdAt'],
  Automations: ['id', 'clientId', 'name', 'trigger', 'threshold', 'action', 'active', 'createdAt'],
  Subtasks: ['id', 'clientId', 'parentId', 'name', 'done', 'createdAt'],
  Comments: ['id', 'clientId', 'parentId', 'text', 'author', 'createdAt'],
  Attachments: ['id', 'clientId', 'parentId', 'name', 'url', 'uploadedBy', 'createdAt', 'waMsgId'],
  Sessions: ['id', 'token', 'role', 'clientId', 'userId', 'createdAt'],
  SystemConfig: ['_id', 'email', 'password']
};

/*** SUPER ADMINS — stored in SystemConfig (uses _id, so read it directly) ***/
function _readSAs() {
  const sh = _sheet('SystemConfig'), vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const hdr = vals[0], out = [];
  for (let i = 1; i < vals.length; i++) {
    const o = { _row: i + 1 };
    hdr.forEach((h, j) => { if (h) o[h] = vals[i][j]; });
    if (o.email) out.push(o);
  }
  return out;
}

/*** RUN ONCE — forces Google to request every scope this app needs ***/
function authorizeAll() {
  const log = [];

  // Drive — folder access + create + share (photos, attachments, logo)
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    log.push('Drive folder OK: ' + folder.getName());
    log.push('Folder access level: ' + folder.getSharingAccess() + ' / ' + folder.getSharingPermission());
    let probe = null;
    try {
      probe = folder.createFile(
        Utilities.newBlob('clap-auth-probe', 'text/plain', '__clap_auth_probe.txt'));
      log.push('Drive CREATE OK: ' + probe.getId());
    } catch (e) { log.push('Drive CREATE FAILED: ' + e.message); }

    if (probe) {
      try {
        probe.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        log.push('Drive SHARE anyone-with-link OK');
      } catch (e) {
        log.push('Drive SHARE anyone FAILED: ' + e.message);
        try {
          probe.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
          log.push('Drive SHARE domain-with-link OK (use DOMAIN_WITH_LINK)');
        } catch (e2) { log.push('Drive SHARE domain FAILED: ' + e2.message); }
      }
      probe.setTrashed(true);
      log.push('Drive delete OK');
    }
  } catch (e) { log.push('DRIVE FAILED: ' + e.message); }

  // Spreadsheets — the database
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    log.push('Sheets OK: ' + ss.getName() + ' (' + ss.getSheets().length + ' tabs)');
  } catch (e) { log.push('SHEETS FAILED: ' + e.message); }

  // Script properties — SA credentials
  try {
    PropertiesService.getScriptProperties().getProperty('SA_EMAIL');
    log.push('Properties OK');
  } catch (e) { log.push('PROPERTIES FAILED: ' + e.message); }

  // Mail — password reset emails
  try {
    log.push('Mail OK, quota remaining: ' + MailApp.getRemainingDailyQuota());
  } catch (e) { log.push('MAIL FAILED: ' + e.message); }

  // Session / user identity
  try {
    log.push('User: ' + Session.getEffectiveUser().getEmail());
  } catch (e) { log.push('SESSION FAILED: ' + e.message); }

  // Triggers — daily task generator
  try {
    log.push('Triggers OK, count: ' + ScriptApp.getProjectTriggers().length);
  } catch (e) { log.push('TRIGGERS FAILED: ' + e.message); }

  // External fetch — only if you ever call an outside API
  try {
    const code = UrlFetchApp.fetch('https://www.google.com/generate_204',
      { muteHttpExceptions: true }).getResponseCode();
    log.push('UrlFetch OK: ' + code);
  } catch (e) { log.push('URLFETCH FAILED: ' + e.message); }

  Logger.log(log.join('\n'));
  return log.join('\n');
}

// /*** SERVE ***/
// function doGet() {
//   return HtmlService.createTemplateFromFile('Index')
//     .evaluate()
//     .setTitle('Checklist')
//     .addMetaTag('viewport', 'width=device-width, initial-scale=1')
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
// }

/*** SERVE ***/
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.ssoBootstrap = '{}';
  try {
    const token = e && e.parameter && e.parameter.sso;
    if (token) {
      const r = OPS.consumeSSOToken({ token: token });
      if (r) t.ssoBootstrap = JSON.stringify(r);
    }
  } catch (err) { /* fall through to normal login screen */ }
  return t.evaluate()
    .setTitle('Checklist')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }

/*** SHEET HELPERS ***/
function _ss() { return SpreadsheetApp.openById(SS_ID); }
function _sheet(name) {
  const ss = _ss(); let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SCHEMA[name]); }
  return sh;
}
function _readAll(name) {
  const sh = _sheet(name), vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const hdr = vals[0], out = [];
  for (let i = 1; i < vals.length; i++) {
    const o = { _row: i + 1 };
    hdr.forEach((h, j) => { if (h) o[h] = vals[i][j]; });
    if (o.id) out.push(o);
  }
  return out;
}

const DEFAULT_PLATFORM_LOGO='https://vivekindustries.biz/wp-content/uploads/2025/02/vivaan-Logo-2.png';

function _uid(){ return Utilities.getUuid().replace(/-/g,'').slice(0,20); }

/*** READABLE SEQUENTIAL IDS — C001, E001, T001, K001 … ***/
const ID_PREFIX = {
  Clients: 'C', Employees: 'E', Templates: 'T', Tasks: 'K', Holidays: 'H',
  Settings: 'S', Subtasks: 'SUB', Comments: 'CM', Attachments: 'AT',
  ActLog: 'AL', AuditLog: 'AU', Announcements: 'AN', Automations: 'AM',
  Sessions: 'SE', SystemConfig: 'sa', Swaps: 'SW'
};

function _nextId(sheetName) {
  const pfx = ID_PREFIX[sheetName];
  if (!pfx) return _uid();                       // unknown sheet → fall back to hex
  const sh = _sheet(sheetName);
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0] || [];
  // SystemConfig uses _id; every other sheet uses id
  let col = hdr.indexOf('id');
  if (col < 0) col = hdr.indexOf('_id');
  if (col < 0) return _uid();

  const re = new RegExp('^' + pfx + '(\\d+)$');
  let max = 0;
  for (let i = 1; i < vals.length; i++) {
    const m = re.exec(String(vals[i][col] || '').trim());
    if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
  }
  const width = (pfx === 'sa') ? 1 : 3;           // sa1, sa2 … but C001, E001 …
  return pfx + String(max + 1).padStart(width, '0');
}

function _now() { return new Date().toISOString(); }
function _coerce(v) {
  if (v === 'TRUE' || v === true) return true;
  if (v === 'FALSE' || v === false) return false;
  return v;
}

/*** PATH RESOLVER: 'clients/X/tasks/T/subtasks' -> {sheet,clientId,parentId} ***/
const COLL_MAP = {
  clients: 'Clients', employees: 'Employees', templates: 'Templates', tasks: 'Tasks',
  holidays: 'Holidays', settings: 'Settings', actLog: 'ActLog', auditLog: 'AuditLog',
  announcements: 'Announcements', automations: 'Automations',
  subtasks: 'Subtasks', comments: 'Comments', attachments: 'Attachments'
};
function _resolve(path) {
  const p = path.split('/').filter(Boolean);
  // clients
  if (p.length === 1) return { sheet: 'Clients' };
  const clientId = p[1];
  // Support 4-segment paths like clients/X/settings/weeklyPromises where docId is p[3]
  if (p.length === 4) return { sheet: COLL_MAP[p[2]], clientId, docId: p[3] };
  if (p.length === 3) return { sheet: COLL_MAP[p[2]], clientId };
  if (p.length === 5) return { sheet: COLL_MAP[p[4]], clientId, parentId: p[3] };
  throw new Error('Bad path: ' + path);
}

/*** PUBLIC API — single entry point ***/
function api(op, payload) {
  try {
    const fn = OPS[op];
    if (!fn) throw new Error('Unknown op: ' + op);
    return { ok: true, data: fn(payload || {}) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
// Batched calls from client
function apiBatch(calls) {
  return calls.map(c => api(c.op, c.payload));
}

const OPS = {

  query: function (p) {
    const r = _resolve(p.path);
    let rows = r.sheet === 'Employees' ? _readDoersFromAuth() : _readAll(r.sheet);
    if (r.clientId) rows = rows.filter(x => String(x.clientId) === String(r.clientId));
    if (r.parentId) rows = rows.filter(x => String(x.parentId) === String(r.parentId));

    const norm = function (v) {
      if (v instanceof Date) {
        return v.getFullYear() < 1901
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm')
          : Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return v;
    };

    (p.where || []).forEach(w => {
      rows = rows.filter(x => {
        const v = norm(x[w[0]]), t = norm(w[2]);
        switch (w[1]) {
          case '==': return String(v) === String(t) || v === t;
          case '>=': return String(v) >= String(t);
          case '<=': return String(v) <= String(t);
          case '>': return String(v) > String(t);
          case '<': return String(v) < String(t);
        }
        return true;
      });
    });

    if (p.orderBy) {
      const f = p.orderBy[0], dir = p.orderBy[1];
      rows.sort((a, b) => { const A = norm(a[f]) || '', B = norm(b[f]) || ''; return dir === 'desc' ? (A < B ? 1 : A > B ? -1 : 0) : (A > B ? 1 : A < B ? -1 : 0); });
    }
    if (p.limit) rows = rows.slice(0, p.limit);
    return rows.map(x => { const { _row, ...rest } = x; Object.keys(rest).forEach(k => rest[k] = _coerce(norm(rest[k]))); return rest; });
  },

  getDoc: function (p) {
    const r = _resolve(p.path);
    // Special handler for settings docs stored with docId (like weeklyPromises)
    const _docId = r.docId || (r.sheet === 'Settings' ? p.id : null);   // ← ADD THIS
    if (r.sheet === 'Settings' && _docId) {                              // ← was r.docId
      const sh = _sheet('Settings');
      const vals = sh.getDataRange().getValues();
      if (vals.length < 2) return null;
      const hdr = vals[0];
      const iId = hdr.indexOf('id'), iCid = hdr.indexOf('clientId'), iDoc = hdr.indexOf('docId'), iData = hdr.indexOf('data');
      
      for (let i = 1; i < vals.length; i++) {
        const rowId = String(vals[i][iId] || vals[i][0]).trim();
        const rowDoc = String(vals[i][iDoc] || vals[i][2]).trim();
        const rowCid = String(vals[i][iCid] || vals[i][1]).trim();
        
        if ((rowId === String(_docId) || rowDoc === String(_docId)) && rowCid === String(r.clientId)) {  // ← use _docId
          try {
            const rawData = vals[i][iData !== -1 ? iData : 3] || '{}';
            const parsedData = JSON.parse(rawData);
            return OPS._clean(parsedData);
          } catch(e) { return {}; }
        }
      }
      return null;
    }
    const source = r.sheet === 'Employees' ? _readDoersFromAuth() : _readAll(r.sheet);
    const rows = source.filter(x => String(x.id) === String(p.id));
    if (!rows.length) return null;
    const { _row, ...rest } = rows[0];
    Object.keys(rest).forEach(k => rest[k] = _coerce(rest[k]));
    return OPS._clean(rest);
  },

  // Flatten Dates so any returned row can cross the sandbox boundary
  // Recurses into plain nested objects/arrays (e.g. {text,by,at} promise entries) instead of
  // stringifying them into the useless literal "[object Object]".
  _clean: function (o) {
      if (o instanceof Date) {
          return o.getFullYear() < 1901
              ? Utilities.formatDate(o, Session.getScriptTimeZone(), 'HH:mm')
              : Utilities.formatDate(o, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      if (o === undefined || o === null) return '';
      if (Array.isArray(o)) return o.map(function (v) { return OPS._clean(v); });
      if (typeof o === 'object') {
          const out = {};
          Object.keys(o).forEach(function (k) { out[k] = OPS._clean(o[k]); });
          return out;
      }
      return o; // primitives (string, number, boolean) pass through untouched
  },

  add: function (p) {
    const r = _resolve(p.path);
    const id = p.id || _nextId(r.sheet);
    const rec = Object.assign({}, p.data, { id });
    if (r.clientId) rec.clientId = r.clientId;
    if (r.parentId) rec.parentId = r.parentId;
    if (!rec.createdAt) rec.createdAt = _now();
    const sh = _sheet(r.sheet), hdr = SCHEMA[r.sheet];
    sh.appendRow(hdr.map(h => rec[h] !== undefined ? rec[h] : ''));
    return OPS._clean(rec);
  },

  set: function (p) { // merge upsert
    const r = _resolve(p.path);
    const _docId = r.docId || (r.sheet === 'Settings' ? p.id : null);   // ← ADD THIS

    // Special handler for settings docs stored with docId (like weeklyPromises)
    if (r.sheet === 'Settings' && _docId) {                              // ← was r.docId
      const sh = _sheet('Settings');
      const vals = sh.getDataRange().getValues();
      const hdr = vals[0] || [];
      
      const iId = hdr.indexOf('id');
      const iCid = hdr.indexOf('clientId');
      const iDoc = hdr.indexOf('docId');
      const iData = hdr.indexOf('data');
      
      let hitRow = -1;
      let currentObj = {};
      
      for (let i = 1; i < vals.length; i++) {
        const rowCid = String(vals[i][iCid !== -1 ? iCid : 1] || '').trim();
        const rowDoc = String(vals[i][iDoc !== -1 ? iDoc : 2] || '').trim();
        
        if (rowDoc === String(_docId) && rowCid === String(r.clientId)) {   // ← use _docId
          hitRow = i + 1;
          try { 
            const rawData = vals[i][iData !== -1 ? iData : 3] || '{}';
            currentObj = JSON.parse(rawData); 
          } catch(e){}
          break;
        }
      }

      const finalData = p.merge ? Object.assign({}, currentObj, p.data) : p.data;
      const dataStr = JSON.stringify(finalData);
      
      if (hitRow > 0) {
        sh.getRange(hitRow, iData !== -1 ? iData + 1 : 4).setValue(dataStr);
      } else {
        const newId = _nextId('Settings');
        const rowArray = ['', '', '', ''];
        rowArray[iId !== -1 ? iId : 0] = newId;
        rowArray[iCid !== -1 ? iCid : 1] = r.clientId;
        rowArray[iDoc !== -1 ? iDoc : 2] = _docId;                         // ← use _docId
        rowArray[iData !== -1 ? iData : 3] = dataStr;
        sh.appendRow(rowArray);
      }
      return OPS._clean(finalData);
    }

    const sh = _sheet(r.sheet), hdr = SCHEMA[r.sheet];
    const rows = _readAll(r.sheet);
    const hit = rows.find(x => String(x.id) === String(p.id));
    if (hit) {
      const merged = p.merge ? Object.assign({}, hit, p.data) : Object.assign({ id: p.id }, p.data);
      if (r.clientId) merged.clientId = r.clientId;
      if (r.parentId) merged.parentId = r.parentId;
      merged.id = p.id;
      sh.getRange(hit._row, 1, 1, hdr.length)
      .setValues([hdr.map(h => merged[h] !== undefined ? merged[h] : '')]);
      return OPS._clean(merged);
    }
    return OPS.add({ path: p.path, id: p.id, data: p.data });
  },

  update: function (p) { return OPS.set({ path: p.path, id: p.id, data: p.data, merge: true }); },

  del: function (p) {
    const r = _resolve(p.path);
    const sh = _sheet(r.sheet);
    const hit = _readAll(r.sheet).find(x => String(x.id) === String(p.id));
    if (hit) {
      try {
        sh.deleteRow(hit._row);
      } catch (e) {
        // Sheets refuses to delete the last non-frozen row (or last row overall).
        // Clear the row's contents instead — functionally equivalent, since _readAll()
        // already skips any row with no id.
        const hdr = SCHEMA[r.sheet] || [];
        if (hdr.length) sh.getRange(hit._row, 1, 1, hdr.length).clearContent();
        else sh.getRange(hit._row, 1, 1, sh.getLastColumn()).clearContent();
      }
    }
    return true;
  },

  bulkAdd: function (p) {
    const r = _resolve(p.path);
    const sh = _sheet(r.sheet), hdr = SCHEMA[r.sheet];
    let seq = null;
    const out = (p.items || []).map(it => {
      if (it.id) { var newId = it.id; }
      else {
        if (seq === null) {                                  // seed once, then increment
          const first = _nextId(r.sheet);
          const m = /^([A-Za-z]+)(\d+)$/.exec(first);
          seq = m ? { pfx: m[1], n: parseInt(m[2], 10), w: m[2].length } : null;
          newId = first;
        } else {
          seq.n++;
          newId = seq.pfx + String(seq.n).padStart(seq.w, '0');
        }
        if (seq && newId === undefined) newId = _nextId(r.sheet);
      }
      const rec = Object.assign({}, it, { id: newId });
      if (r.clientId) rec.clientId = r.clientId;
      if (!rec.createdAt) rec.createdAt = _now();
      return rec;
    });
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, hdr.length)
      .setValues(out.map(rec => hdr.map(h => rec[h] !== undefined ? rec[h] : '')));
    return out.length;
  },

  bulkSaveTaskTemplates: function (p) {
    OPS._assertWithinOfficeHours();
    const sh = _sheet('Templates'), hdr = SCHEMA.Templates;
    let seq = null;
    const out = (p.templates || []).map(function (tmpl) {
      let newId;
      if (seq === null) {
        const first = _nextId('Templates');
        const m = /^([A-Za-z]+)(\d+)$/.exec(first);
        seq = m ? { pfx: m[1], n: parseInt(m[2], 10), w: m[2].length } : null;
        newId = first;
      } else {
        seq.n++;
        newId = seq.pfx + String(seq.n).padStart(seq.w, '0');
      }
      const rec = Object.assign({}, tmpl, { id: newId, clientId: p.clientId, isActive: true, createdAt: _now() });
      return rec;
    });
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, hdr.length)
      .setValues(out.map(function (rec) { return hdr.map(function (h) { return rec[h] !== undefined ? rec[h] : ''; }); }));
    return { count: out.length, templateIds: out.map(function (r) { return r.id; }) };
  },

  // Runs the generator immediately (used right after a bulk import so today's occurrence shows
  // up without waiting for the next scheduled trigger tick).
  runGeneratorNow: function () {
    return { generated: generateDailyTasks() };
  },

  batch: function (p) {
    (p.ops || []).forEach(o => {
      if (o.type === 'set') OPS.set(o);
      else if (o.type === 'update') OPS.update(o);
      else if (o.type === 'delete') OPS.del(o);
    });
    return true;
  },

  /*** SSO — consumes a one-time token issued by the Hub, logs the user straight in ***/
  consumeSSOToken: function (p) {
    const token = String((p && p.token) || '');
    if (!token) return { ok: false, error: 'Missing token' };
    const sh = _authSS().getSheetByName('AUTH_SSO_TOKENS');
    if (!sh) return { ok: false, error: 'SSO not configured' };
    const vals = sh.getDataRange().getValues();
    const hdr = vals[0];
    const tCol = hdr.indexOf('TOKEN'), eCol = hdr.indexOf('EMAIL'), sCol = hdr.indexOf('SYSTEM_ID'),
          expCol = hdr.indexOf('EXPIRES_AT'), usedCol = hdr.indexOf('USED');
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][tCol]) !== token) continue;
      if (vals[i][usedCol] === true) return { ok: false, error: 'Token already used' };
      const exp = vals[i][expCol];
      if (exp instanceof Date && new Date() > exp) return { ok: false, error: 'Token expired' };
      if (String(vals[i][sCol]) !== CLAP_SYSTEM_ID) return { ok: false, error: 'Wrong system' };
      sh.getRange(i + 1, usedCol + 1).setValue(true); // single-use — burn it immediately
      return OPS._authenticateByEmail(String(vals[i][eCol]));
    }
    return { ok: false, error: 'Invalid token' };
  },

  // From a list of AUTH_USERS rows, keep only ACTIVE users with ACTIVE CLAP access
  _clapCandidates: function (users) {
    if (!users.length) return { error: { ok: false, error: 'Invalid credentials' } };
    const active = users.filter(x => String(x.STATUS || '').toUpperCase() === 'ACTIVE');
    if (!active.length) return { error: { ok: false, error: 'Account inactive. Contact admin.' } };
    const acc = _authReadAll('AUTH_USER_ACCESS').filter(a =>
      a.SYSTEM_ID === CLAP_SYSTEM_ID && String(a.STATUS).toUpperCase() === 'ACTIVE');
    const list = active.filter(x => acc.some(a => String(a.USER_ID) === String(x.USER_ID)));
    if (!list.length) return { error: { ok: false, error: 'You do not have access to CLAP. Contact admin.' } };
    return { list: list };
  },

  // Shared session-building logic — used by SSO. Shared emails cannot be auto-resolved to one person.
  _authenticateByEmail: function (email) {
    const u = String(email || '').trim().toLowerCase();
    const cands = OPS._clapCandidates(_authReadAll('AUTH_USERS').filter(x => String(x.EMAIL || '').trim().toLowerCase() === u));
    if (cands.error) return cands.error;
    if (cands.list.length > 1) return { ok: false, needLogin: true, email: u, error: 'This email is shared by multiple employees. Please sign in with your own User ID and password.' };
    const user = cands.list[0];
    const access = _authReadAll('AUTH_USER_ACCESS').find(a =>
      String(a.USER_ID) === String(user.USER_ID) &&
      a.SYSTEM_ID === CLAP_SYSTEM_ID &&
      String(a.STATUS).toUpperCase() === 'ACTIVE');
    if (!access) return { ok: false, error: 'You do not have access to CLAP. Contact admin.' };
    const roles = _authReadAll('AUTH_ROLES');
    const roleRow = roles.find(r => r.ROLE_ID === access.ROLE_ID);
    let roleCode = roleRow ? String(roleRow.ROLE_CODE || '').toLowerCase() : '';
    if (!roleCode) {
      const m = String(access.ROLE_ID || '').match(/_(ADMIN|PC|STAFF)$/i);
      roleCode = m ? m[1].toLowerCase() : 'staff';
    }
    return {
      ok: true, role: roleCode, clientId: CLAP_CLIENT_ID, userId: String(user.USER_ID),
      user: {
        name: String(user.FULL_NAME || ''), adminName: String(user.FULL_NAME || ''),
        dept: String(user.DEPARTMENT || ''), designation: String(user.DESIGNATION || ''),
        email: String(user.EMAIL || ''), phone: String(user.MOBILE || ''),
        companyName: 'Vivek Industries'
      },
      sessionToken: _mkSession(roleCode, CLAP_CLIENT_ID, user.USER_ID)
    };
  },

  /*** AUTH — centralised Auth sheet only (no superadmin, no per-client login) ***/
  authenticate: function (p) {
    const u = String(p.username || '').trim().toLowerCase(), pw = String(p.password || '');

    // Match by USER_ID or EMAIL, then by password — one email can belong to several employees
    const matched = _authReadAll('AUTH_USERS').filter(x =>
      (String(x.USER_ID || '').trim().toLowerCase() === u || String(x.EMAIL || '').trim().toLowerCase() === u) &&
      String(x.PASSWORD || '') === pw);
    const cands = OPS._clapCandidates(matched);
    if (cands.error) return cands.error;
    if (cands.list.length > 1) return { ok: false, error: 'This email is shared by multiple employees. Please sign in with your User ID instead of email (ask admin for your ID).' };
    const user = cands.list[0];

    const access = _authReadAll('AUTH_USER_ACCESS').find(a =>
      String(a.USER_ID) === String(user.USER_ID) &&
      a.SYSTEM_ID === CLAP_SYSTEM_ID &&
      String(a.STATUS).toUpperCase() === 'ACTIVE');
    if (!access) return { ok: false, error: 'You do not have access to CLAP. Contact admin.' };

    const roles = _authReadAll('AUTH_ROLES');
    const roleRow = roles.find(r => r.ROLE_ID === access.ROLE_ID);
    let roleCode = roleRow ? String(roleRow.ROLE_CODE || '').toLowerCase() : '';
    if (!roleCode) {
      // AUTH_ROLES row missing — derive from ROLE_ID suffix (e.g. CLAP_ADMIN -> admin) instead of defaulting to staff
      const m = String(access.ROLE_ID || '').match(/_(ADMIN|PC|STAFF)$/i);
      roleCode = m ? m[1].toLowerCase() : 'staff';
    }

    return {
      ok: true, role: roleCode, clientId: CLAP_CLIENT_ID, userId: String(user.USER_ID),
      user: {
        name: String(user.FULL_NAME || ''), adminName: String(user.FULL_NAME || ''),
        dept: String(user.DEPARTMENT || ''), designation: String(user.DESIGNATION || ''),
        email: String(user.EMAIL || ''), phone: String(user.MOBILE || ''),
        companyName: 'Vivek Industries'
      },
      sessionToken: _mkSession(roleCode, CLAP_CLIENT_ID, user.USER_ID)
    };
  },

  /*** OFFICE HOURS ENFORCEMENT — reads Settings > company > officeStart/officeEnd and checks
       the current time against it, always in India Standard Time regardless of who's asking or
       what timezone their device/session is set to. Applies with no role exception whatsoever. ***/
  _officeHoursWindow: function () {
    try {
      const sh = _sheet('Settings');
      const vals = sh.getDataRange().getValues();
      const hdr = vals[0] || [];
      const iDoc = hdr.indexOf('docId'), iData = hdr.indexOf('data');
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][iDoc]) === 'company') {
          const obj = JSON.parse(vals[i][iData] || '{}');
          if (obj.officeStart && obj.officeEnd) return { start: obj.officeStart, end: obj.officeEnd };
          break;
        }
      }
    } catch (e) { /* fall through */ }
    return null; // not configured -> no restriction
  },
  _isWithinOfficeHours: function () {
    const win = OPS._officeHoursWindow();
    if (!win) return true; // not configured -> never restrict
    const toMin = function (s) { const p = s.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10); };
    const nowIST = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'HH:mm');
    const nowM = toMin(nowIST), startM = toMin(win.start), endM = toMin(win.end);
    return startM <= endM ? (nowM >= startM && nowM <= endM) : (nowM >= startM || nowM <= endM); // handles an overnight window too, just in case
  },
  _assertWithinOfficeHours: function () {
    const win = OPS._officeHoursWindow();
    if (!win) return;
    if (!OPS._isWithinOfficeHours()) {
      throw new Error('New tasks can only be added during office hours (' + win.start + '–' + win.end + ' IST). Please contact admin.');
    }
  },
  checkOfficeHours: function () {
    const win = OPS._officeHoursWindow();
    if (!win) return { ok: true };
    return { ok: OPS._isWithinOfficeHours(), start: win.start, end: win.end };
  },

  /*** CLOUD-FUNCTION EQUIVALENTS ***/
  getClients: function () { throw new Error('Super Admin panel is disabled — CLAP is single-tenant (Vivek Industries).'); },
  saveClient: function () { throw new Error('Super Admin panel is disabled — CLAP is single-tenant (Vivek Industries).'); },
  deleteClient: function () { throw new Error('Super Admin panel is disabled — CLAP is single-tenant (Vivek Industries).'); },
  saveEmployee: function () {
    throw new Error('Employees are managed centrally in the Auth sheet. Use grantClapAccess() there.');
  },
  deleteEmployee: function () {
    throw new Error('Employees are managed centrally in the Auth sheet. Revoke access in AUTH_USER_ACCESS.');
  },
  saveTaskTemplate: function (p) {
    if (!p.templateId) OPS._assertWithinOfficeHours(); // only NEW task creation is restricted — editing an existing template is not
    const id = p.templateId || _nextId('Templates');
    const rotationOn = !!(p.rotationEnabled && p.rotationDoer1Id && p.rotationDoer2Id);
    const data = {
      name: p.name, dept: p.dept, freq: p.freq,
      priority: p.priority, plannedTime: p.plannedTime, scheduleConfig: p.scheduleConfig || '',
      photoRequired: !!p.photoRequired, skipHolidays: !!p.skipHolidays, isActive: true,
      rotationEnabled: rotationOn,
      rotationDoer1Id: rotationOn ? p.rotationDoer1Id : '',
      rotationDoer1Name: rotationOn ? p.rotationDoer1Name : '',
      rotationDoer2Id: rotationOn ? p.rotationDoer2Id : '',
      rotationDoer2Name: rotationOn ? p.rotationDoer2Name : '',
      rotationDays: rotationOn ? (parseInt(p.rotationDays, 10) || 7) : ''
    };
    if (rotationOn) {
      // Fallback display fields — generation always recomputes the true current doer live,
      // this is only so existing UI reading t.doerName has something sensible to show.
      const createdAt = p.templateId ? null : _now(); // existing template keeps its original createdAt
      const anchorSrc = createdAt || p._existingCreatedAt || _now();
      const today = _dstr(new Date());
      const current = _resolveRotationDoer({ createdAt: anchorSrc, rotationDays: data.rotationDays, rotationDoer1Id: data.rotationDoer1Id, rotationDoer1Name: data.rotationDoer1Name, rotationDoer2Id: data.rotationDoer2Id, rotationDoer2Name: data.rotationDoer2Name }, today);
      data.doerId = current.id; data.doerName = current.name;
    } else {
      data.doerId = p.doerId; data.doerName = p.doerName;
    }
    OPS.set({ path: 'clients/' + p.clientId + '/templates', id, merge: !!p.templateId, data: data });
    return { ok: true, templateId: id };
  },
  deleteTaskTemplate: function (p) {
    // SOFT delete — keep the row so _nextId never recycles this id. A recycled templateId could
    // later get matched by propagateTemplateEditToInstances() against an unrelated old template's
    // completed history and silently corrupt it.
    OPS.update({ path: 'clients/' + p.clientId + '/templates', id: p.templateId, data: { isActive: false } });
    return { ok: true };
  },
  bulkImportTasks: function (p) {
    const items = (p.tasks || []).map(t => ({
      name: t.name, dept: t.dept, doerId: t.doerId || '', doerName: t.doerName || '',
      freq: t.freq || 'D', plannedDate: t.plannedDate, plannedTime: t.plannedTime || '00:00',
      priority: t.priority || 'M', status: 'pending', isActive: true
    }));
    const n = OPS.bulkAdd({ path: 'clients/' + p.clientId + '/tasks', items });
    return { ok: true, count: n };
  },
  broadcastAnnouncement: function (p) {
    _readAll('Clients').forEach(c =>
      OPS.add({ path: 'clients/' + c.id + '/announcements', data: { msg: p.msg } }));
    return { ok: true };
  },
  getSuperAdmin: function () { throw new Error('Super Admin is disabled — manage users in the Auth sheet.'); },
  updateSuperAdmin: function () { throw new Error('Super Admin is disabled — manage users in the Auth sheet.'); },
  deleteSuperAdmin: function () { throw new Error('Super Admin is disabled — manage users in the Auth sheet.'); },
  verify2FAOTP: function () { return { ok: true }; },
  sendPasswordResetEmail: function (p) {
    MailApp.sendEmail(p.email, 'CLAP password reset', 'Contact your administrator to reset your password.');
    return { ok: true };
  },

  /*** PLATFORM LOGO — shown to Super Admin (no clientId of their own) ***/
  // Platform logo lives in SystemConfig as a row with _id = 'platform_logo'
  getPlatformLogo: function(){
    const sh=_sheet('SystemConfig'), vals=sh.getDataRange().getValues();
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][0]).trim()==='platform_logo')
        return { logoUrl: String(vals[i][1]||'') || DEFAULT_PLATFORM_LOGO };
    }
    return { logoUrl: DEFAULT_PLATFORM_LOGO };
  },

  setPlatformLogo: function(p){
    const url=String(p.logoUrl||'').trim();
    if(url && !/^https?:\/\//i.test(url))
      throw new Error('Logo must be a full https:// URL');
    if(url.length>2000)
      throw new Error('URL too long — upload the image and paste its link instead');

    const sh=_sheet('SystemConfig'), vals=sh.getDataRange().getValues();
    let row=-1;
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][0]).trim()==='platform_logo'){ row=i+1; break; }
    }
    if(row>0) sh.getRange(row,1,1,3).setValues([['platform_logo', url, '']]);
    else sh.appendRow(['platform_logo', url, '']);

    return { ok:true, logoUrl: url || DEFAULT_PLATFORM_LOGO };
  },

  /*** SESSION PERSISTENCE — disabled server-side.
       With "Execute as: Me", UserProperties is ONE store shared by every visitor, so it restored
       whoever logged in last. Sessions now live only in the user's own browser (localStorage) or Hub SSO. ***/
  saveSession: function () { return { ok: true }; },

  restoreSession: function () {
    try { PropertiesService.getUserProperties().deleteProperty('CLAP_SESS'); } catch (e) { } // purge old shared session
    return null;
  },

  clearSession: function () { return { ok: true }; },

  // Save a client logo: upload bytes to Drive, store only the short URL
  saveClientLogo: function(p){
    if(!p.clientId) throw new Error('clientId required');
    const safe=String(p.name||'logo').replace(/[^\w.\-]/g,'_');
    const nm='logo_'+p.clientId+'_'+Date.now()+'_'+safe;
    const blob=Utilities.newBlob(Utilities.base64Decode(p.base64), p.mime, nm);
    const file=DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);
    try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch(e){ /* policy-blocked — parent folder sharing applies */ }
    const url='https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w400';

    const r=_resolve('clients/'+p.clientId+'/settings');
    const sh=_sheet(r.sheet), vals=sh.getDataRange().getValues();
    const hdr=vals[0];
    const iDoc=hdr.indexOf('docId'), iData=hdr.indexOf('data'), iCid=hdr.indexOf('clientId');
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][iCid])===String(p.clientId) && String(vals[i][iDoc])==='company'){
        let obj={};
        try{ obj=JSON.parse(vals[i][iData]||'{}'); }catch(e){}
        obj.logoUrl=url;
        sh.getRange(i+1, iData+1).setValue(JSON.stringify(obj));
        return { ok:true, url:url };
      }
    }
    sh.appendRow([_nextId('Settings'), p.clientId, 'company', JSON.stringify({logoUrl:url})]);
    return { ok:true, url:url };
  },

  /*** TASK SWAP — temporary reassignment for a date range, template ownership untouched ***/
  saveSwap: function (p) {
    if (!p.fromDoerId || !p.toDoerId || !p.startDate || !p.endDate) throw new Error('fromDoerId, toDoerId, startDate, endDate are required');
    const emps = _readDoersFromAuth();
    const from = emps.find(e => e.id === p.fromDoerId), to = emps.find(e => e.id === p.toDoerId);
    if (!from || !to) throw new Error('Employee not found');
    const id = _nextId('Swaps');
    const rec = {
      clientId: p.clientId, fromDoerId: from.id, fromDoerName: from.name, toDoerId: to.id, toDoerName: to.name,
      startDate: p.startDate, endDate: p.endDate, reason: p.reason || '', status: 'active',
      createdBy: p.createdBy || '', createdAt: _now()
    };
    _sheet('Swaps').appendRow([id, rec.clientId, rec.fromDoerId, rec.fromDoerName, rec.toDoerId, rec.toDoerName, rec.startDate, rec.endDate, rec.reason, rec.status, rec.createdBy, rec.createdAt]);

    // Reassign already-generated instances in BOTH directions — this is a mutual exchange, not one-way.
    // Rotation-enabled templates are excluded — their doer for any date is fixed by the rotation
    // cycle, not swappable, per spec.
    const sh = _sheet('Tasks'), hdr = SCHEMA.Tasks;
    const inRange = t => _dstr(t.plannedDate) >= _dstr(p.startDate) && _dstr(t.plannedDate) <= _dstr(p.endDate) && t.status !== 'done';
    const rotationTemplateIds = {};
    _readAll('Templates').forEach(function (t) { if (_isRotationTemplate(t)) rotationTemplateIds[t.id] = true; });
    const notRotation = t => !rotationTemplateIds[t.templateId];
    const fromTasks = _readAll('Tasks').filter(t => String(t.clientId) === String(p.clientId) && String(t.doerId) === String(from.id) && inRange(t) && notRotation(t));
    const toTasks = _readAll('Tasks').filter(t => String(t.clientId) === String(p.clientId) && String(t.doerId) === String(to.id) && inRange(t) && notRotation(t));
    fromTasks.forEach(t => {
      const merged = Object.assign({}, t, { doerId: to.id, doerName: to.name, swapId: id, originalDoerId: from.id, originalDoerName: from.name });
      sh.getRange(t._row, 1, 1, hdr.length).setValues([hdr.map(h => merged[h] !== undefined ? merged[h] : '')]);
    });
    toTasks.forEach(t => {
      const merged = Object.assign({}, t, { doerId: from.id, doerName: from.name, swapId: id, originalDoerId: to.id, originalDoerName: to.name });
      sh.getRange(t._row, 1, 1, hdr.length).setValues([hdr.map(h => merged[h] !== undefined ? merged[h] : '')]);
    });
    return { ok: true, swapId: id, reassignedCount: fromTasks.length + toTasks.length };
  },

  listSwaps: function (p) {
    return _readAll('Swaps').filter(s => String(s.clientId) === String(p.clientId))
      .map(s => { const { _row, ...rest } = s; return OPS._clean(rest); })
      .sort((a, b) => (b.startDate || '').localeCompare(a.startDate || ''));
  },

  cancelSwap: function (p) {
    const sh = _sheet('Swaps');
    const hit = _readAll('Swaps').find(s => String(s.id) === String(p.swapId));
    if (!hit) throw new Error('Swap not found');
    const hdr = SCHEMA.Swaps;
    const merged = Object.assign({}, hit, { status: 'cancelled' });
    sh.getRange(hit._row, 1, 1, hdr.length).setValues([hdr.map(h => merged[h] !== undefined ? merged[h] : '')]);

    // Revert any not-yet-done tasks that were reassigned by this swap back to the original doer
    const tasks = _readAll('Tasks').filter(t => String(t.swapId) === String(p.swapId) && t.status !== 'done');
    const tsh = _sheet('Tasks'), thdr = SCHEMA.Tasks;
    tasks.forEach(t => {
      const merged2 = Object.assign({}, t, { doerId: t.originalDoerId, doerName: t.originalDoerName, swapId: '', originalDoerId: '', originalDoerName: '' });
      tsh.getRange(t._row, 1, 1, thdr.length).setValues([thdr.map(h => merged2[h] !== undefined ? merged2[h] : '')]);
    });
    return { ok: true, revertedCount: tasks.length };
  },

  /*** FILE UPLOAD (replaces firebase.storage) ***/
  uploadFile: function (p) {
    const safe = String(p.name || 'file').replace(/[^\w.\-]/g, '_');
    const nm = (p.prefix ? p.prefix + '_' : '') + Date.now() + '_' + safe;
    const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mime, nm);
    const file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      console.log(err);
    }
    const id = file.getId();
    return {
      url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400',
      fullUrl: 'https://drive.google.com/file/d/' + id + '/view',
      fileId: id
    };
  }
};

/*** ROLES ***/
const VALID_ROLES = ['admin', 'pc', 'staff'];   // employee-assignable roles
function _empRole(e) {
  const r = String(e.role || '').toLowerCase();
  if (VALID_ROLES.indexOf(r) > -1) return r;
  return (e.isCoordinator === true || e.isCoordinator === 'TRUE') ? 'pc' : 'staff';
}

function _mkSession(role, clientId, userId) {
  const token = _uid();
  _sheet('Sessions').appendRow([token, token, role, clientId || '', userId || '', _now()]);
  return token;
}

/*** DAILY TASK GENERATOR — install as time-driven trigger, 12:05 AM ***/
function _dstr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}
function _tstr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return String(v || '');
}

/*** Reads the configured company Working Days (from Settings > Company); defaults to Mon–Fri ***/
function _readWorkDays() {
  try {
    const sh = _sheet('Settings');
    const vals = sh.getDataRange().getValues();
    const hdr = vals[0] || [];
    const iDoc = hdr.indexOf('docId'), iData = hdr.indexOf('data');
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][iDoc]) === 'company') {
        const obj = JSON.parse(vals[i][iData] || '{}');
        if (obj.workDays && obj.workDays.length) return obj.workDays;
        break;
      }
    }
  } catch (e) { /* fall through to default */ }
  return [1, 2, 3, 4, 5, 6]; // Mon–Sat (Sunday off)
}

/*** Resolves which of the 2 rotation doers is "on duty" for a given date, based on a fixed
     N-day cycle anchored to the template's createdAt. Calendar-day based — Sundays/holidays
     still count against the current person's turn, per spec. ***/
function _resolveRotationDoer(t, dateStr) {
  const days = parseInt(t.rotationDays, 10) || 7;
  const anchor = new Date(String(t.createdAt || '').slice(0, 10) + 'T00:00:00');
  const target = new Date(dateStr + 'T00:00:00');
  let daysSince = Math.floor((target - anchor) / 86400000);
  if (daysSince < 0) daysSince = 0; // safety — target before anchor shouldn't happen
  const turnIndex = Math.floor(daysSince / days) % 2;
  return turnIndex === 0
    ? { id: t.rotationDoer1Id, name: t.rotationDoer1Name, turnLabel: '1st', dayInTurn: (daysSince % days) + 1 }
    : { id: t.rotationDoer2Id, name: t.rotationDoer2Name, turnLabel: '2nd', dayInTurn: (daysSince % days) + 1 };
}
function _isRotationTemplate(t) {
  const v = t.rotationEnabled;
  const on = (v === true) || (typeof v === 'string' && v.toUpperCase() === 'TRUE');
  return on && t.rotationDoer1Id && t.rotationDoer2Id;
}

function generateDailyTasks(dateOverride) {
  // Prevent concurrent executions from racing each other and generating duplicate batches —
  // e.g. if a duplicate trigger exists, or a manual run overlaps with the real nightly trigger.
  const _lock = LockService.getScriptLock();
  try {
    _lock.waitLock(30000); // wait up to 30s for any in-progress run to finish
  } catch (e) {
    Logger.log('generateDailyTasks: could not acquire lock, another run is in progress — skipping this call.');
    return 0;
  }
  try {
    return _generateDailyTasksLocked(dateOverride);
  } finally {
    _lock.releaseLock();
  }
}

function _generateDailyTasksLocked(dateOverride) {
  // Time-driven triggers auto-pass an event object as the first arg — only trust dateOverride
  // when it's a real date (string/Date), never blindly new Date(dateOverride).
  let todayD;
  if (dateOverride && (typeof dateOverride === 'string' || dateOverride instanceof Date)) {
    const parsed = new Date(dateOverride);
    todayD = isNaN(parsed.getTime()) ? new Date() : parsed;
  } else {
    todayD = new Date();
  }
  todayD.setHours(0, 0, 0, 0);
  const ds = Utilities.formatDate(todayD, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const dow = todayD.getDay(), dom = todayD.getDate(), mon = todayD.getMonth();
  const tmpls = _readAll('Templates').filter(t => t.isActive !== false && t.isActive !== 'FALSE');
  const existing = _readAll('Tasks').filter(t => _dstr(t.plannedDate) === ds);

  // ── HOLIDAY SHIFT SUPPORT ──────────────────────────────────────────────
  // A holiday itself only ever resolves to 'prev' or 'next' (no holiday-level skip anymore).
  // Whether a given TEMPLATE actually shifts is decided per-template via t.skipHolidays:
  //   skipHolidays = true  -> this template's occurrence on a holiday is simply dropped, no shift
  //   skipHolidays = false -> this template's occurrence shifts per that holiday's prev/next
  const holidayRows = _readAll('Holidays');
  const holidayDateSet = {}; holidayRows.forEach(function (h) { holidayDateSet[_dstr(h.date)] = h; });
  const workDays = _readWorkDays();
  function _isWorkingDate(dateStr) {
    const dd = new Date(dateStr + 'T00:00:00');
    return workDays.indexOf(dd.getDay()) > -1 && !holidayDateSet[dateStr];
  }
  function _nearestWorkingDate(fromDateStr, dir) {
    let d = new Date(fromDateStr + 'T00:00:00');
    for (let step = 0; step < 60; step++) {
      d.setDate(d.getDate() + dir);
      const cand = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (_isWorkingDate(cand)) return cand;
    }
    return fromDateStr; // safety fallback
  }
  function _isWorkingDateInclusive(dateStr, dir) {
    return _isWorkingDate(dateStr) ? dateStr : _nearestWorkingDate(dateStr, dir);
  }
  // Resolves each holiday's shift direction ('prev'/'next', default 'prev' for legacy/blank rows)
  // to its target working day, independent of any template.
  const holidayShiftTarget = {}; // holidayDateStr -> resolved target working day
  holidayRows.forEach(function (h) {
    const dir = (h.shiftDir === 'next') ? 'next' : 'prev';
    const hds = _dstr(h.date);
    holidayShiftTarget[hds] = _nearestWorkingDate(hds, dir === 'prev' ? -1 : 1);
  });
  // Map: target working day -> list of {holidayDate} that resolve to it — built per-template below,
  // since whether a template even lands here depends on its own skipHolidays flag.
  function _holidaysShiftingInto(dateStr) {
    return Object.keys(holidayShiftTarget).filter(function (hds) { return holidayShiftTarget[hds] === dateStr; });
  }

  // ── NON-WORKING WEEKDAY SHIFT (e.g. Sunday) — applies the SAME shift concept as holidays to any
  // calendar day whose weekday isn't a Working Day, for any frequency EXCEPT Daily (Daily already
  // has its own unconditional Sunday-skip, unrelated to this — kept as-is on purpose).
  // Default direction is always 'prev' — Saturday is structurally guaranteed to exist as the
  // previous working day for a Sunday under the default Mon–Sat setting, so no lead-time/"too
  // close" exception is needed here the way it is for user-added holidays.
  const weekendShiftTarget = {};
  for (let i = 0; i < 21; i++) {
    const d = new Date(todayD); d.setDate(todayD.getDate() + i);
    const wds = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (holidayDateSet[wds]) continue; // real holidays already have their own explicit shift config
    if (workDays.indexOf(d.getDay()) > -1) continue; // already a working day, nothing to shift
    weekendShiftTarget[wds] = _nearestWorkingDate(wds, -1);
  }
  function _weekendShiftsInto(dateStr) {
    return Object.keys(weekendShiftTarget).filter(function (wds) { return weekendShiftTarget[wds] === dateStr; });
  }

  // ── LAST-DAY-OF-MONTH SUPPORT — e.g. a Monthly template set to run on the 31st, in a month
  // that only has 28/29/30 days. Always resolves backward to a real working day; not configurable.
  const _lastDom = new Date(todayD.getFullYear(), todayD.getMonth() + 1, 0).getDate();
  let _monthEndTarget = null;
  if (_lastDom < 31) {
    const _lastDateStr = Utilities.formatDate(new Date(todayD.getFullYear(), todayD.getMonth(), _lastDom), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    _monthEndTarget = _isWorkingDateInclusive(_lastDateStr, -1);
  }

  if (holidayDateSet[ds]) { Logger.log(ds + ': holiday, skipped — any non-skipping template shifts in via the block below'); return 0; }

    // ── END-OF-DAY RESOLUTION for YESTERDAY'S still-pending tasks ──────────────
  // High priority: must be done same-day — lock as permanently missed (full penalty), no further completion possible.
  // Medium/Low priority: stay pending, now "overdue" — still completable, lighter penalty than an outright High miss.
  const yd = new Date(todayD); yd.setDate(yd.getDate() - 1);
  const ydStr = Utilities.formatDate(yd, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const todayStr = Utilities.formatDate(todayD, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  // Hard rule: only a date strictly BEFORE today can ever be swept as missed.
  // This holds true no matter how frequently this function runs (every minute or once a day).
  // Runs on every invocation (no "already swept" flag) — re-checking is cheap and guarantees
  // any pending High-priority task from a past day gets caught, even if it was created after
  // an earlier sweep already ran for that date.
  if (ydStr < todayStr) {
    const pastPending = _readAll('Tasks').filter(t => _dstr(t.plannedDate) < todayStr && t.status === 'pending' && t.priority === 'H');
    if (pastPending.length) {
      const sh = _sheet('Tasks'), hdr = SCHEMA.Tasks;
      pastPending.forEach(t => {
        const merged = Object.assign({}, t, { status: 'missed', lockedNotDone: true });
        sh.getRange(t._row, 1, 1, hdr.length).setValues([hdr.map(h => merged[h] !== undefined ? merged[h] : '')]);
      });
    }
  }
  // Medium/Low pending tasks from yesterday are intentionally left untouched — they remain
  // status:'pending' with their original plannedDate, showing as "overdue" in the UI, and
  // can still be marked done afterward.

  const activeSwaps = _readAll('Swaps').filter(s => s.status === 'active' && ds >= _dstr(s.startDate) && ds <= _dstr(s.endDate));
  // Bidirectional lookup: returns {toId, toName} whichever side this doer is on in an active swap
  const swapFor = (doerId) => {
    const s = activeSwaps.find(x => String(x.fromDoerId) === String(doerId) || String(x.toDoerId) === String(doerId));
    if (!s) return null;
    const isFrom = String(s.fromDoerId) === String(doerId);
    return { id: s.id, toDoerId: isFrom ? s.toDoerId : s.fromDoerId, toDoerName: isFrom ? s.toDoerName : s.fromDoerName };
  };

    // Parse a template's scheduleConfig JSON safely; returns null if absent/invalid (falls back to legacy single-slot behavior)
  function _parseSchedule(t) {
    if (!t.scheduleConfig) return null;
    try { return JSON.parse(t.scheduleConfig); } catch (e) { return null; }
  }
  // Returns array of time strings this template is due at for the given day, or [] if not due today
  function _dueTimesForDay(t, cfg) {
    const freq = t.freq;
    if (freq === 'D') {
      if (dow === 0) return [];
      return (cfg && cfg.times && cfg.times.length) ? cfg.times : [t.plannedTime || '00:00'];
    }
    if (freq === 'W') {
      const days = (cfg && cfg.days && cfg.days.length) ? cfg.days : [1];
      if (days.indexOf(dow) === -1) return [];
      return (cfg && cfg.times && cfg.times.length) ? cfg.times : [t.plannedTime || '00:00'];
    }
    if (freq === 'M') {
      const dates = (cfg && cfg.dates && cfg.dates.length) ? cfg.dates : [1];
      if (dates.indexOf(dom) === -1) return [];
      return (cfg && cfg.times && cfg.times.length) ? cfg.times : [t.plannedTime || '00:00'];
    }
    if (freq === 'Q') {
      if (mon % 3 !== 0) return []; // only Jan/Apr/Jul/Oct
      const dates = (cfg && cfg.dates && cfg.dates.length) ? cfg.dates : [1];
      if (dates.indexOf(dom) === -1) return [];
      return (cfg && cfg.times && cfg.times.length) ? cfg.times : [t.plannedTime || '00:00'];
    }
    if (freq === 'Y') {
      const dates = (cfg && cfg.dates && cfg.dates.length) ? cfg.dates : [{ month: 1, day: 1 }];
      const match = dates.some(function (d) { return (d.month - 1) === mon && d.day === dom; });
      if (!match) return [];
      return (cfg && cfg.times && cfg.times.length) ? cfg.times : [t.plannedTime || '00:00'];
    }
    return [];
  }
  // Same rules as _dueTimesForDay but for an arbitrary calendar date (used to re-evaluate a
  // shifted holiday's original date's due-templates, landing them on a different plannedDate).
  function _dueTimesForCalendarDate(t, cfg, cDow, cDom, cMon) {
    const freq = t.freq;
    const dfltTimes = (cfg && cfg.times && cfg.times.length) ? cfg.times : [t.plannedTime || '00:00'];
    if (freq === 'D') return cDow === 0 ? [] : dfltTimes;
    if (freq === 'W') { const days = (cfg && cfg.days && cfg.days.length) ? cfg.days : [1]; return days.indexOf(cDow) === -1 ? [] : dfltTimes; }
    if (freq === 'M') { const dates = (cfg && cfg.dates && cfg.dates.length) ? cfg.dates : [1]; return dates.indexOf(cDom) === -1 ? [] : dfltTimes; }
    if (freq === 'Q') { if (cMon % 3 !== 0) return []; const dates = (cfg && cfg.dates && cfg.dates.length) ? cfg.dates : [1]; return dates.indexOf(cDom) === -1 ? [] : dfltTimes; }
    if (freq === 'Y') { const dates = (cfg && cfg.dates && cfg.dates.length) ? cfg.dates : [{ month: 1, day: 1 }]; return dates.some(function (d) { return (d.month - 1) === cMon && d.day === cDom; }) ? dfltTimes : []; }
    return [];
  }

  const due = [];
  tmpls.forEach(t => {
    const cfg = _parseSchedule(t);
    const times = _dueTimesForDay(t, cfg);
    if (!times.length) return;
    const isRot = _isRotationTemplate(t);
    const rot = isRot ? _resolveRotationDoer(t, ds) : null;
    const sw = isRot ? null : swapFor(t.doerId); // rotation templates are never touched by Swaps
    times.forEach(function (tm) {
      // Dedup key = (templateId + time + originDate), NOT plannedDate — so a normal occurrence
      // and any holiday-shifted-in occurrence landing on the same day stay as separate rows.
      const already = existing.concat(due).some(x => String(x.templateId) === String(t.id) && _tstr(x.plannedTime) === _tstr(tm) && _dstr(x.originDate || x.plannedDate) === ds);
      if (already) return;
      due.push({
        clientId: t.clientId, templateId: t.id, name: t.name, dept: t.dept,
        doerId: isRot ? rot.id : (sw ? sw.toDoerId : t.doerId),
        doerName: isRot ? rot.name : (sw ? sw.toDoerName : t.doerName),
        swapId: sw ? sw.id : '', originalDoerId: sw ? t.doerId : '', originalDoerName: sw ? t.doerName : '',
        freq: t.freq, priority: t.priority || 'M',
        plannedDate: ds, plannedTime: tm || '00:00', originDate: ds, status: 'pending',
        photoRequired: t.photoRequired, isActive: true,
        notes: isRot ? ('Rotation: ' + rot.name + '\'s turn (' + rot.turnLabel + ', day ' + rot.dayInTurn + '/' + (parseInt(t.rotationDays, 10) || 7) + ')') : ''
      });
    });
  });

  // ── INCOMING HOLIDAY SHIFTS — for every holiday whose resolved target working day is today (ds),
  // check EACH template individually: only templates with skipHolidays !== true shift in. ──
  _holidaysShiftingInto(ds).forEach(function (hds) {
    const h = holidayDateSet[hds]; if (!h) return;
    const hd = new Date(hds + 'T00:00:00');
    const hDow = hd.getDay(), hDom = hd.getDate(), hMon = hd.getMonth();
    tmpls.forEach(function (t) {
      if (t.skipHolidays === true || t.skipHolidays === 'true') return; // this template opts out of shifting
      const cfg = _parseSchedule(t);
      const times = _dueTimesForCalendarDate(t, cfg, hDow, hDom, hMon);
      if (!times.length) return;
      const isRot = _isRotationTemplate(t);
      const rot = isRot ? _resolveRotationDoer(t, ds) : null;
      const sw = isRot ? null : swapFor(t.doerId);
      const dir = (h.shiftDir === 'next') ? 'next' : 'prev';
      times.forEach(function (tm) {
        const already = existing.concat(due).some(x => String(x.templateId) === String(t.id) && _tstr(x.plannedTime) === _tstr(tm) && _dstr(x.originDate || x.plannedDate) === hds);
        if (already) return;
        due.push({
          clientId: t.clientId, templateId: t.id, name: t.name, dept: t.dept,
          doerId: isRot ? rot.id : (sw ? sw.toDoerId : t.doerId),
          doerName: isRot ? rot.name : (sw ? sw.toDoerName : t.doerName),
          swapId: sw ? sw.id : '', originalDoerId: sw ? t.doerId : '', originalDoerName: sw ? t.doerName : '',
          freq: t.freq, priority: t.priority || 'M',
          plannedDate: ds, plannedTime: tm || '00:00', originDate: hds, status: 'pending',
          photoRequired: t.photoRequired, isActive: true,
          notes: 'Shifted from holiday "' + (h.name || hds) + '" (' + dir + ')' + (isRot ? ' — Rotation: ' + rot.name + '\'s turn' : '')
        });
      });
    });
  });

  // ── INCOMING NON-WORKING WEEKDAY SHIFTS (e.g. a Monthly task whose due date lands on a
  // Sunday) — generated on the resolved earlier working day instead of silently going out. ──
  _weekendShiftsInto(ds).forEach(function (wds) {
    const wd = new Date(wds + 'T00:00:00');
    const wDow = wd.getDay(), wDom = wd.getDate(), wMon = wd.getMonth();
    tmpls.forEach(function (t) {
      if (t.freq === 'D') return; // Daily keeps its own separate, unconditional Sunday-skip
      if (t.skipHolidays === true || t.skipHolidays === 'true') return; // same opt-out as holidays
      const cfg = _parseSchedule(t);
      const times = _dueTimesForCalendarDate(t, cfg, wDow, wDom, wMon);
      if (!times.length) return;
      const isRot = _isRotationTemplate(t);
      const rot = isRot ? _resolveRotationDoer(t, ds) : null;
      const sw = isRot ? null : swapFor(t.doerId);
      times.forEach(function (tm) {
        const already = existing.concat(due).some(x => String(x.templateId) === String(t.id) && _tstr(x.plannedTime) === _tstr(tm) && _dstr(x.originDate || x.plannedDate) === wds);
        if (already) return;
        due.push({
          clientId: t.clientId, templateId: t.id, name: t.name, dept: t.dept,
          doerId: isRot ? rot.id : (sw ? sw.toDoerId : t.doerId),
          doerName: isRot ? rot.name : (sw ? sw.toDoerName : t.doerName),
          swapId: sw ? sw.id : '', originalDoerId: sw ? t.doerId : '', originalDoerName: sw ? t.doerName : '',
          freq: t.freq, priority: t.priority || 'M',
          plannedDate: ds, plannedTime: tm || '00:00', originDate: wds, status: 'pending',
          photoRequired: t.photoRequired, isActive: true,
          notes: 'Shifted from non-working day (' + wds + ') to previous working day' + (isRot ? ' — Rotation: ' + rot.name + '\'s turn' : '')
        });
      });
    });
  });

  // ── MONTH-END 31st HANDLING — only fires on the resolved target day computed above ──
  if (_monthEndTarget && ds === _monthEndTarget) {
    tmpls.forEach(function (t) {
      if (t.freq !== 'M') return;
      let cfg = {}; try { cfg = t.scheduleConfig ? JSON.parse(t.scheduleConfig) : {}; } catch (e) { }
      const dates = (cfg.dates && cfg.dates.length) ? cfg.dates : [1];
      if (dates.indexOf(31) === -1) return; // this template isn't meant to run at month-end
      const times = (cfg.times && cfg.times.length) ? cfg.times : [t.plannedTime || '00:00'];
      const isRot = _isRotationTemplate(t);
      const rot = isRot ? _resolveRotationDoer(t, ds) : null;
      const sw = isRot ? null : swapFor(t.doerId);
      times.forEach(function (tm) {
        const already = existing.concat(due).some(x => String(x.templateId) === String(t.id) && _tstr(x.plannedTime) === _tstr(tm) && _dstr(x.originDate || x.plannedDate) === ds);
        if (already) return;
        due.push({
          clientId: t.clientId, templateId: t.id, name: t.name, dept: t.dept,
          doerId: isRot ? rot.id : (sw ? sw.toDoerId : t.doerId),
          doerName: isRot ? rot.name : (sw ? sw.toDoerName : t.doerName),
          swapId: sw ? sw.id : '', originalDoerId: sw ? t.doerId : '', originalDoerName: sw ? t.doerName : '',
          freq: t.freq, priority: t.priority || 'M',
          plannedDate: ds, plannedTime: tm || '00:00', originDate: ds, status: 'pending',
          photoRequired: t.photoRequired, isActive: true,
          notes: 'Month-end (31st) task — this month has ' + _lastDom + ' days' + (isRot ? ' — Rotation: ' + rot.name + '\'s turn' : '')
        });
      });
    });
  }

  const sh = _sheet('Tasks'), hdr = SCHEMA.Tasks;
  let kSeq = null;
  const rows = due.map(d => {
    if (kSeq === null) {
      const first = _nextId('Tasks');
      const m = /^([A-Za-z]+)(\d+)$/.exec(first);
      kSeq = m ? { pfx: m[1], n: parseInt(m[2], 10), w: m[2].length } : null;
      d.id = first;
    } else { kSeq.n++; d.id = kSeq.pfx + String(kSeq.n).padStart(kSeq.w, '0'); }
    d.createdAt = _now();
    return hdr.map(h => d[h] !== undefined ? d[h] : '');
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, hdr.length).setValues(rows);
  Logger.log(ds + ': generated ' + rows.length + ' task(s)');
  return rows.length;
} // end of _generateDailyTasksLocked

/*** RUN ONCE — installs the nightly trigger ***/
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'generateDailyTasks') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('generateDailyTasks').timeBased().atHour(0).nearMinute(15).everyDays(1).create();
  Logger.log('Trigger installed. Total triggers: ' + ScriptApp.getProjectTriggers().length);
}

/*** RUN ONCE — creates the missed days (default: last 7) ***/
function backfillTasks(days) {
  const n = days || 7, log = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() - i);
    log.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' -> ' + generateDailyTasks(d));
  }
  Logger.log(log.join('\n'));
  return log.join('\n');
}

/*** ONE-TIME BOOTSTRAP — run manually from the editor, then delete the row you don't need ***/
function bootstrap() {
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty('SA_EMAIL', 'Pankajkannavedia@dtableanalytics.com');
  sp.setProperty('SA_PASS', 'Dtable123@');

  // Seed the first super admin into SystemConfig (idempotent)
  if (!_readSAs().some(x => String(x.email || '').toLowerCase() === 'pankajkannavedia@dtableanalytics.com')) {
    _sheet('SystemConfig').appendRow([_uid(), 'Pankajkannavedia@dtableanalytics.com', 'Dtable123@']);
  }

  // Create a client admin so you can log in as an Admin too
  const existing = _readAll('Clients').find(c =>
    String(c.adminEmail).toLowerCase() === 'pankajkannavedia@dtableanalytics.com');
  if (!existing) {
    OPS.add({
      path: 'clients', data: {
        companyName: 'Vivek Industries',
        adminName: 'Pankaj',
        adminEmail: 'Pankajkannavedia@dtableanalytics.com',
        adminPassword: 'Dtable123@',
        plan: 'enterprise',
        accessExpiry: '2027-12-31',
        isLocked: false
      }
    });
  }
  Logger.log('Bootstrap done. SA_EMAIL=' + sp.getProperty('SA_EMAIL'));
}

/*** ONE-TIME CLEANUP — fix any task rows corrupted by the trigger-event-object bug (plannedDate stuck at epoch) ***/
function fixEpochTasks() {
  const sh = _sheet('Tasks'), hdr = SCHEMA.Tasks;
  const rows = _readAll('Tasks').filter(t => _dstr(t.plannedDate) === '1970-01-01');
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  rows.forEach(t => {
    const merged = Object.assign({}, t, { plannedDate: todayStr });
    sh.getRange(t._row, 1, 1, hdr.length).setValues([hdr.map(h => merged[h] !== undefined ? merged[h] : '')]);
  });
  Logger.log('Fixed ' + rows.length + ' epoch-dated task(s) -> ' + todayStr);
  return rows.length;
}


function fixWronglyLockedToday() {
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sh = _sheet('Tasks'), hdr = SCHEMA.Tasks;
  const rows = _readAll('Tasks').filter(t => _dstr(t.plannedDate) === todayStr && t.status === 'missed' && t.lockedNotDone);
  rows.forEach(t => {
    const merged = Object.assign({}, t, { status: 'pending', lockedNotDone: false });
    sh.getRange(t._row, 1, 1, hdr.length).setValues([hdr.map(h => merged[h] !== undefined ? merged[h] : '')]);
  });
  Logger.log('Reverted ' + rows.length + ' wrongly-locked task(s) back to pending.');
}

/*** ONE-TIME REPAIR — clears any Monday-promise entries corrupted to the literal "[object Object]" ***/
function fixCorruptedPromises() {
  const sh = _sheet('Settings');
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0] || [];
  const iDoc = hdr.indexOf('docId'), iData = hdr.indexOf('data');
  let fixed = 0;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][iDoc]) !== 'weeklyPromises') continue;
    let obj;
    try { obj = JSON.parse(vals[i][iData] || '{}'); } catch (e) { continue; }
    Object.keys(obj).forEach(function (k) {
      const entry = obj[k];
      if (entry && typeof entry === 'object' && entry.text === '[object Object]') {
        entry.text = '';
        fixed++;
      } else if (entry === '[object Object]') {
        obj[k] = '';
        fixed++;
      }
    });
    sh.getRange(i + 1, iData + 1).setValue(JSON.stringify(obj));
  }
  Logger.log('Fixed ' + fixed + ' corrupted promise entr(ies).');
  return fixed;
}


function fixRotationEnabledFlags() {
  const sh = _sheet('Templates');
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0];
  const iEn = hdr.indexOf('rotationEnabled'), iD1 = hdr.indexOf('rotationDoer1Id'), iD2 = hdr.indexOf('rotationDoer2Id');
  let fixed = 0;
  for (let i = 1; i < vals.length; i++) {
    const hasBoth = vals[i][iD1] && vals[i][iD2];
    const en = vals[i][iEn];
    const looksOn = en !== '' && en !== false && String(en).toUpperCase() !== 'FALSE';
    if (hasBoth && looksOn && en !== true) { sh.getRange(i + 1, iEn + 1).setValue(true); fixed++; }
  }
  Logger.log('Fixed ' + fixed + ' rotationEnabled flag(s).');
  return fixed;
}

function _testGenFor17Sept() {
  return generateDailyTasks('2026-09-17');
}


function _fixOneOverdueRow() {
  const taskId = 'K170'; // e.g. 'K123' — check the row in your sheet
  const sh = _sheet('Tasks');
  const rows = _readAll('Tasks').filter(t => t.id === taskId);
  if (!rows.length) { Logger.log('Task not found: ' + taskId); return; }
  const t = rows[0];
  const merged = Object.assign({}, t, {
    plannedDate: '2026-09-19',
    originDate: '2026-09-20',
    notes: 'Shifted from non-working day (2026-09-20) to previous working day'
  });
  sh.getRange(t._row, 1, 1, SCHEMA.Tasks.length).setValues([SCHEMA.Tasks.map(h => merged[h] !== undefined ? merged[h] : '')]);
  Logger.log('Fixed row ' + taskId);
}