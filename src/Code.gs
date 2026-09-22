// =============================================================
// G8 PE Profile — Apps Script server (FIS, Grade 8 PE)
// =============================================================
// Lives INSIDE the FIS-owned Google Sheet that holds the roster and the
// profile data. Deployed as a web app: Execute as Me (the Sheet owner),
// access "Anyone within fis.edu". The student's identity comes from
// Session.getActiveUser() on the server. The browser never sends an email
// and never receives anyone else's row.
//
// Tabs (menu "PE Profile → Set up tabs" creates them):
//   Config    Key · Value · What it does
//   Students  Email · Name · Class             the roster, teacher-maintained
//   Teachers  Email · Name                     who gets the teacher view
//   Profile   Email · Style · Goal · Updated   written by the app, one row per student
//
// Privacy rules baked in (PRIVACY.md is the reviewable version):
//   • identity_() is the only place an email is read. It comes from Google, never the client.
//   • bootstrap() returns the caller's own row only. Nothing returns the roster.
//   • Errors thrown to the client never contain an email (they reach Stackdriver logs).
//   • The photo is a URL to the student's own Google directory photo, fetched at
//     runtime only when photo_lookup is TRUE. Nothing is copied or stored.
// =============================================================

var TABS = {
  Config:   ['Key', 'Value', 'What it does'],
  Students: ['Email', 'Name', 'Class'],
  Teachers: ['Email', 'Name'],
  Profile:  ['Email', 'Style', 'Goal', 'Updated']
};
var PROFILE_KEY = 'Email';

var CONFIG_DEFAULTS = {
  domain:       ['fis.edu', 'Only accounts in this Google Workspace domain are served. Checked on the server as well as by the deployment setting.'],
  year_label:   ['26/27', 'Shown on the card header'],
  app_title:    ['My PE Profile', 'Browser tab title'],
  photo_lookup: ['FALSE', 'TRUE to show the student\'s Google directory photo (runtime only, never stored). Needs the People API advanced service enabled for the script.'],
  web_fonts:    ['FALSE', 'TRUE to load Archivo + Inter from Google Fonts (sends the viewer\'s IP address to fonts.googleapis.com). FALSE = system fonts, no third-party requests.'],
  goal_max:     ['140', 'Maximum characters for the student\'s own goal']
};

// The six PE Styles. Keys are what the Profile tab stores; the client owns the colours.
var STYLES = {
  competitor: 'The Competitor',
  team:       'The Team Player',
  improver:   'The Improver',
  explorer:   'The Explorer',
  energizer:  'The Energizer',
  thinker:    'The Active Thinker'
};

var CACHE_KEY_CONFIG = 'g8pe_config_v1';
var CACHE_SECONDS = 120;

// ---------- Sheet helpers ----------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tab_(name) { return ss_().getSheetByName(name); }

function ensureTab_(name, headers) {
  var book = ss_();
  var s = book.getSheetByName(name);
  if (!s) {
    s = book.insertSheet(name);
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    s.setFrozenRows(1);
    return s;
  }
  var lastCol = Math.max(1, s.getLastColumn());
  var existing = s.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  while (existing.length && existing[existing.length - 1] === '') existing.pop();
  var missing = headers.filter(function (h) { return existing.indexOf(h) === -1; });
  if (existing.length === 0) s.getRange(1, 1, 1, headers.length).setValues([headers]);
  else if (missing.length) s.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  s.setFrozenRows(1);
  return s;
}

function readTab_(name) {
  var s = tab_(name);
  if (!s || s.getLastRow() < 2) return [];
  var data = s.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r], obj = {}, empty = true;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      var v = row[c];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      obj[headers[c]] = v;
      if (v !== '' && v !== null && v !== undefined) empty = false;
    }
    if (!empty) out.push(obj);
  }
  return out;
}

function str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function lower_(v) { return str_(v).toLowerCase(); }
function bool_(v) { return /^(true|yes|1)$/i.test(str_(v)); }

// ---------- Config ----------
function config_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(CACHE_KEY_CONFIG);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* rebuild */ } }
  var kv = {};
  Object.keys(CONFIG_DEFAULTS).forEach(function (k) { kv[k] = CONFIG_DEFAULTS[k][0]; });
  readTab_('Config').forEach(function (r) { var k = str_(r.Key); if (k && str_(r.Value) !== '') kv[k] = str_(r.Value); });
  cache.put(CACHE_KEY_CONFIG, JSON.stringify(kv), CACHE_SECONDS);
  return kv;
}
function clearConfigCache() { CacheService.getScriptCache().remove(CACHE_KEY_CONFIG); }

// ---------- Identity (the only place an email is read) ----------
// Returns { role, ... }. role is one of:
//   anonymous     no Google account (deployment misconfigured, or preview)
//   wrong_domain  signed in, but not an account in the configured domain
//   student       on the Students tab  → name, klass (and teacher:true if also a teacher)
//   teacher       on the Teachers tab, or the Sheet owner, but not on the roster
//   unknown       a domain account that is on neither tab
function identity_(cfg) {
  var email = lower_(Session.getActiveUser().getEmail());
  var owner = lower_(Session.getEffectiveUser().getEmail());
  if (!email) return { role: 'anonymous' };
  var dom = lower_(cfg.domain);
  if (dom && email.slice(-(dom.length + 1)) !== '@' + dom) return { role: 'wrong_domain' };
  var teacher = email === owner || readTab_('Teachers').some(function (t) { return lower_(t.Email) === email; });
  var me = readTab_('Students').filter(function (r) { return lower_(r.Email) === email; })[0];
  if (me) return { role: 'student', email: email, name: str_(me.Name), klass: str_(me.Class), teacher: teacher };
  if (teacher) return { role: 'teacher', email: email };
  return { role: 'unknown', email: email };
}
function requireStudent_(cfg) {
  var id = identity_(cfg);
  if (id.role !== 'student') throw new Error('Not signed in with a school account that is on the class list');
  return id;
}
// Roster names arrive as "Last, First" from the school system (sometimes with
// only an initial before the comma) or as "First Last". The page greets by
// first name and shows the first name on the card; the part before a comma is
// ignored. Without a comma the full name is kept for display.
function nameParts_(raw) {
  var name = str_(raw).replace(/\s+/g, ' ');
  var first, display;
  if (name.indexOf(',') !== -1) {
    first = str_(name.split(',').slice(1).join(',')).split(' ')[0] || str_(name.split(',')[0]);
    display = first;
  } else {
    first = name.split(' ')[0] || '';
    display = name;
  }
  return { first: first, display: display, initials: initials_(display) };
}
function initials_(name) {
  var p = str_(name).split(/\s+/);
  return ((p[0] || '').charAt(0) + (p.length > 1 ? p[p.length - 1].charAt(0) : '')).toUpperCase();
}

// Google directory photo, runtime only. Off unless Config photo_lookup = TRUE.
// Tries, in order, whichever advanced services are enabled in the editor
// (Services → +): People API, then Admin SDK Directory (domain_public view,
// which any domain user may read when directory sharing is on). The URL is
// handed to the browser, which loads the image from Google. Nothing is stored.
// Any failure (service off, no directory access, no photo) falls back to initials.
function photoUrl_(cfg, email) {
  if (!bool_(cfg.photo_lookup)) return '';
  try {
    var res = People.People.searchDirectoryPeople({
      query: email, readMask: 'photos', pageSize: 3,
      sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE']
    });
    var people = (res && res.people) || [];
    for (var i = 0; i < people.length; i++) {
      var photos = people[i].photos || [];
      for (var j = 0; j < photos.length; j++) if (photos[j].url && !photos[j].default) return photos[j].url;
    }
  } catch (e) { /* try the next service */ }
  try {
    var u = AdminDirectory.Users.get(email, { viewType: 'domain_public', projection: 'basic' });
    if (u && u.thumbnailPhotoUrl && !u.isDefaultPhoto) return u.thumbnailPhotoUrl;
  } catch (e) { /* fall back to initials */ }
  return '';
}

// ---------- Web app entry ----------
// When built as a single file (dist/Code.gs) the HTML files are embedded here,
// so there is only one thing to paste into Apps Script.
var EMBEDDED_HTML = {};
function doGet() {
  var cfg = config_();
  var t = EMBEDDED_HTML.Index ? HtmlService.createTemplate(EMBEDDED_HTML.Index) : HtmlService.createTemplateFromFile('Index');
  t.appTitle = cfg.app_title;
  t.fontsLink = bool_(cfg.web_fonts)
    ? '<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@600;700;800&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">'
    : '';
  return t.evaluate()
    .setTitle(cfg.app_title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}
function include(name) { return EMBEDDED_HTML[name] !== undefined ? EMBEDDED_HTML[name] : HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ---------- Reads ----------
function profileRow_(email) {
  return readTab_('Profile').filter(function (r) { return lower_(r[PROFILE_KEY]) === email; })[0] || null;
}

// Everything the page needs, for the caller only.
function bootstrap() {
  var cfg = config_();
  var id = identity_(cfg);
  var out = {
    role: id.role,
    config: { yearLabel: cfg.year_label, appTitle: cfg.app_title, goalMax: parseInt(cfg.goal_max, 10) || 140 },
    styles: STYLES
  };
  if (id.role === 'student') {
    var p = profileRow_(id.email) || {};
    var nm = nameParts_(id.name);
    out.student = {
      name: nm.display,
      first: nm.first,
      klass: id.klass,
      initials: nm.initials,
      style: STYLES[str_(p.Style)] ? str_(p.Style) : '',
      goal: str_(p.Goal),
      photoUrl: photoUrl_(cfg, id.email)
    };
    if (id.teacher) out.alsoTeacher = true;
  } else if (id.role === 'teacher' || id.role === 'unknown') {
    out.email = id.email; // the caller's own address, so they can see which account they used
  }
  return out;
}

// ---------- Writes (always for the caller's own row) ----------
function upsertProfile_(email, fields) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = ensureTab_('Profile', TABS.Profile);
    var lastRow = s.getLastRow(), lastCol = s.getLastColumn();
    var headers = s.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    var keyCol = headers.indexOf(PROFILE_KEY);
    var rowIdx = -1;
    if (lastRow > 1) {
      var keys = s.getRange(2, keyCol + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) if (lower_(keys[i][0]) === email) { rowIdx = i + 2; break; }
    }
    var row;
    if (rowIdx === -1) { rowIdx = lastRow + 1; row = headers.map(function () { return ''; }); row[keyCol] = email; }
    else row = s.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
    Object.keys(fields).forEach(function (k) { var c = headers.indexOf(k); if (c !== -1) row[c] = fields[k]; });
    var u = headers.indexOf('Updated'); if (u !== -1) row[u] = new Date();
    s.getRange(rowIdx, 1, 1, lastCol).setValues([row]);
  } finally { lock.releaseLock(); }
}

function saveStyle(style) {
  var cfg = config_();
  var id = requireStudent_(cfg);
  var key = lower_(style);
  if (!STYLES[key]) throw new Error('Unknown style');
  upsertProfile_(id.email, { Style: key });
  return { ok: true, style: key };
}

function saveGoal(goal) {
  var cfg = config_();
  var id = requireStudent_(cfg);
  var max = parseInt(cfg.goal_max, 10) || 140;
  var clean = str_(goal).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').slice(0, max);
  upsertProfile_(id.email, { Goal: clean });
  return { ok: true, goal: clean };
}

// ---------- Sheet menu (teacher, inside the Sheet) ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PE Profile')
    .addItem('1. Set up tabs', 'setupTabs')
    .addItem('Clear config cache', 'clearConfigCache')
    .addSeparator()
    .addItem('End of year: clear Profile data…', 'clearProfileData')
    .addToUi();
}

function setupTabs() {
  Object.keys(TABS).forEach(function (n) { ensureTab_(n, TABS[n]); });
  var c = tab_('Config');
  var have = readTab_('Config').map(function (r) { return str_(r.Key); });
  var add = Object.keys(CONFIG_DEFAULTS).filter(function (k) { return have.indexOf(k) === -1; })
    .map(function (k) { return [k, CONFIG_DEFAULTS[k][0], CONFIG_DEFAULTS[k][1]]; });
  if (add.length) c.getRange(c.getLastRow() + 1, 1, add.length, 3).setValues(add);
  var t = tab_('Teachers');
  if (t.getLastRow() < 2) t.getRange(2, 1, 1, 2).setValues([[Session.getEffectiveUser().getEmail(), 'Sheet owner (automatic)']]);
  clearConfigCache();
  try { SpreadsheetApp.getUi().alert('Tabs are ready. Fill Students (Email · Name · Class), then Deploy → New deployment → Web app: Execute as Me, access Anyone within the school.'); } catch (e) { /* no UI when run headless */ }
}

// Retention: wipes every student-written row (styles, goals). Roster is untouched.
function clearProfileData() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert('Clear all Profile rows?', 'This deletes every student\'s saved style and goal. The Students tab is not touched. Continue?', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  var s = ensureTab_('Profile', TABS.Profile);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  ui.alert('Profile data cleared.');
}
