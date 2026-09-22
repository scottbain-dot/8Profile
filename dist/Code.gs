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


// =============================================================
// Embedded web-app files (generated by dev/build-single.js — do not edit here;
// edit src/*.html and rebuild).
// =============================================================
EMBEDDED_HTML = {
  Index: "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"UTF-8\">\n<base target=\"_top\">\n<title><?= appTitle ?></title>\n<?!= fontsLink ?>\n<?!= include('Styles') ?>\n</head>\n<body>\n<div id=\"app\"><div class=\"loading\"><div class=\"spinner\"></div>Loading your profile…</div></div>\n<div class=\"ov\" id=\"ov\"><div class=\"ovcard\" id=\"ovcard\"></div></div>\n<div class=\"toast\" id=\"toast\"></div>\n<?!= include('Rubrics') ?>\n<?!= include('App') ?>\n</body>\n</html>\n",
  Styles: "<style>\n/* Visual system from design/My_PE_Profile_mockup.html (profile, light) and\n   design/G8_Build_Your_Card.html (card, dark). FIS maroon + gold, Archivo + Inter\n   with a system fallback unless Config web_fonts = TRUE. */\n:root{\n  --maroon:#6f1d2c;--maroon-t:#f3e6e8;--gold:#e7b64b;\n  --fit:#e8730c;--fit-t:#fdecdc;--fund:#0e8f86;--fund-t:#d9f2f0;--skill:#2f6df0;--skill-t:#e8f0fe;\n  --ap:#1f9d57;--ap-t:#e6f6ec;--rep:#8a3bb0;--rep-t:#f3e8fa;\n  --ink:#161213;--ink-soft:#463c3e;--muted:#7a6f71;--paper:#f7f4f2;--card:#fff;--line:#e7e0df;\n  --disp:\"Archivo\",\"Barlow Condensed\",\"Arial Narrow\",system-ui,-apple-system,\"Segoe UI\",Roboto,sans-serif;\n  --body:\"Inter\",system-ui,-apple-system,\"Segoe UI\",Roboto,sans-serif;\n  /* card skin — set by the chosen PE Style */\n  --ink1:#7d2233;--ink2:#6f1d2c;--ink3:#4a1019;--acc:#e7b64b;--accs:#f3d689;--rib:#3a1008;\n  --frame:linear-gradient(145deg,#e0a566,#a15c2b 50%,#7a441f 62%,#d68a4e); /* bronze */\n}\n*{box-sizing:border-box}\nhtml,body{margin:0}\nbody{background:var(--paper);color:var(--ink);font-family:var(--body);line-height:1.5;-webkit-font-smoothing:antialiased;min-height:100vh}\nbody.dark{background:radial-gradient(circle at 50% 0%,#2a2024,#120d0e);color:#fff}\nbutton{font:inherit}\n.wrap{max-width:940px;margin:0 auto;padding:24px 16px 60px}\n\n/* loading / toast */\n.loading{display:flex;align-items:center;gap:12px;justify-content:center;padding:80px 0;color:var(--muted);font-size:14px}\n.spinner{width:22px;height:22px;border:3px solid var(--line);border-top-color:var(--maroon);border-radius:50%;animation:spin .9s linear infinite}\n@keyframes spin{to{transform:rotate(360deg)}}\n.toast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%) translateY(20px);background:var(--ink);color:#fff;padding:10px 16px;border-radius:10px;font-size:13.5px;font-weight:500;opacity:0;pointer-events:none;transition:.2s;box-shadow:0 8px 24px rgba(0,0,0,.3);z-index:60}\n.toast.show{opacity:1;transform:translateX(-50%) translateY(0)}\n.toast.err{background:#ffd9d9;color:#7a1414}\n\n/* ── profile view (light) ───────────────────────────────────────── */\n.topline{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:14px;flex-wrap:wrap}\n.demo-tag{display:inline-block;font-size:11px;color:var(--muted);background:#fff;border:1px solid var(--line);border-radius:999px;padding:3px 11px}\n.yearlab{font-family:var(--disp);font-weight:700;font-size:12px;letter-spacing:.14em;color:var(--maroon)}\n.cardteaser{display:flex;align-items:center;gap:16px;background:linear-gradient(135deg,var(--ink1),var(--ink2) 55%,var(--ink3));border:1px solid #c9982f;border-radius:16px;padding:13px 18px;margin:6px 0 10px;position:relative;overflow:hidden;box-shadow:0 10px 26px rgba(0,0,0,.28),0 0 18px rgba(231,182,75,.4);flex-wrap:wrap;transition:background .3s}\n.cardteaser::after{content:\"\";position:absolute;top:-60%;left:-70%;width:45%;height:220%;background:linear-gradient(115deg,transparent,rgba(255,255,255,.28),transparent);transform:rotate(9deg);animation:sweep 4.2s ease-in-out infinite;pointer-events:none}\n@keyframes sweep{0%{left:-70%}62%,100%{left:150%}}\n@media(prefers-reduced-motion:reduce){.cardteaser::after,.card,.cinner::after{animation:none}}\n.ct-av{width:52px;height:52px;border-radius:50%;overflow:hidden;border:2px solid var(--acc);flex:none;position:relative;z-index:2;background:rgba(0,0,0,.25);display:grid;place-items:center;font-family:var(--disp);font-weight:700;font-size:20px;color:var(--accs)}\n.ct-av img{width:100%;height:100%;object-fit:cover;display:block}\n.ct-id{position:relative;z-index:2;min-width:150px;flex:1}\n.ct-name{font-family:var(--disp);font-weight:700;font-size:22px;color:#fff;line-height:1;display:flex;align-items:center;gap:8px;flex-wrap:wrap}\n.ct-tier{font-family:var(--disp);font-size:11px;font-weight:700;color:var(--rib);background:var(--frame);padding:2px 8px;border-radius:20px}\n.ct-sub{font-size:11px;color:rgba(255,255,255,.7);margin-top:1px}\n.ct-arch{font-family:var(--disp);font-weight:600;font-size:13px;color:var(--accs);margin-top:3px}\n.ct-stats{display:flex;gap:16px;position:relative;z-index:2;margin-left:6px}\n.ct-stats div{text-align:center}.ct-stats b{font-family:var(--disp);font-size:18px;color:#fff;display:block;line-height:1}\n.ct-stats span{font-size:9px;color:rgba(255,255,255,.65);text-transform:uppercase;letter-spacing:.04em}\n.ct-btn{margin-left:auto;position:relative;z-index:2;background:linear-gradient(145deg,#f3d689,#c9982f);color:#3a1008;font-family:var(--disp);font-weight:700;font-size:14px;text-decoration:none;padding:9px 16px;border-radius:10px;white-space:nowrap;box-shadow:0 3px 8px rgba(0,0,0,.25);border:none;cursor:pointer}\n.ct-btn:hover{filter:brightness(1.06)}\n.ct-btn.pulse{animation:pulse 1.6s ease-in-out infinite}\n@keyframes pulse{0%,100%{box-shadow:0 3px 8px rgba(0,0,0,.25)}50%{box-shadow:0 3px 8px rgba(0,0,0,.25),0 0 0 6px rgba(231,182,75,.25)}}\n.goals{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:16px 18px;margin:0 0 8px}\n.goals .hd{font-family:var(--disp);font-weight:600;font-size:17px;margin:0 0 6px}\n.goals .empty{font-size:12.5px;color:var(--muted)}\nh3.sec{font-family:var(--disp);font-weight:600;font-size:22px;margin:30px 0 3px;display:flex;align-items:center;gap:8px;flex-wrap:wrap}\nh3.sec .em{font-size:16px;width:31px;height:31px;border-radius:9px;display:inline-grid;place-items:center;background:var(--maroon-t)}\n.seclede{font-size:12.5px;color:var(--muted);margin:0 0 13px}\n.tap{font-size:10.5px;font-weight:600;font-family:var(--disp);color:var(--muted)}\n.grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}\n@media(max-width:760px){.grid3,.grid4{grid-template-columns:1fr 1fr}}\n@media(max-width:420px){.grid3,.grid4{grid-template-columns:1fr}}\n.tile{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:13px 15px 14px;cursor:pointer;transition:.15s;position:relative;text-align:left;width:100%;color:var(--ink)}\n.tile:hover{box-shadow:0 5px 16px rgba(0,0,0,.08);transform:translateY(-1px)}\n.tile:focus-visible{outline:2px solid var(--maroon);outline-offset:2px}\n.tile .tnrow{display:flex;align-items:center;gap:6px}.tile .em{font-size:15px}.tile .tn{font-weight:600;font-size:13px}\n.tile .tsub{font-size:11px;color:var(--muted);margin:1px 0 10px;min-height:13px}\n.tile .plus{position:absolute;top:11px;right:12px;color:#c9c1c0;font-size:14px;line-height:1}\n.track{display:flex;gap:4px}.track i{height:9px;border-radius:3px;flex:1;background:#ece7e6}\n.tracklab{font-family:var(--disp);font-size:12px;font-weight:600;margin-top:6px;color:var(--muted)}\n.rung{font-family:var(--disp);font-size:11px;color:var(--muted);margin-top:5px}\n.measure{font-family:var(--disp);font-weight:700;font-size:22px;line-height:1;color:#c9c1c0}.measure small{font-size:11px;color:var(--muted);font-weight:400}\n.a-fit .track i.on{background:var(--fit)}.a-fund .track i.on{background:var(--fund)}.a-skill .track i.on{background:var(--skill)}.a-ap .track i.on{background:var(--ap)}\n.fhero{background:var(--fit-t);border:1.5px solid var(--fit);border-radius:14px;padding:14px 16px;cursor:pointer;display:flex;align-items:center;gap:13px;margin-bottom:12px;transition:.15s;width:100%;text-align:left;color:var(--ink)}\n.fhero:hover{box-shadow:0 5px 16px rgba(232,115,12,.15)}\n.fhero .fx{font-size:22px;flex:none}.fhero .fm{flex:1}.fhero .fm .n{font-weight:600;font-size:15px}.fhero .fm .s{font-size:12px;color:var(--ink-soft)}\n.fhero .fv{font-family:var(--disp);font-weight:700;font-size:24px;color:var(--fit);text-align:right;line-height:1}.fhero .fv span{font-size:11px;color:var(--muted);font-weight:400;display:block}\n.band{background:var(--ink);border-radius:14px;padding:15px 16px;cursor:pointer;color:#fff;transition:.15s;position:relative;text-align:left;width:100%;border:none}\n.band:hover{box-shadow:0 6px 18px rgba(0,0,0,.18);transform:translateY(-1px)}\n.band .bn{font-family:var(--disp);font-weight:700;font-size:17px;margin:0 0 8px;display:flex;align-items:center;gap:6px}\n.band .btrack{display:flex;gap:3px;margin-bottom:6px}.band .btrack i{height:7px;flex:1;border-radius:2px;background:#4a3b3e}.band .btrack i.on{background:#d7a6ec}\n.band .blab{font-family:var(--disp);font-size:12px;color:#bfb0c4;font-weight:600}.band .plus{position:absolute;top:13px;right:14px;color:#9a8aa0;font-size:15px}\n.foot{margin-top:22px;border:1px solid var(--line);background:#fff;border-radius:14px;padding:14px 18px;font-size:13px;color:var(--ink-soft);display:flex;gap:11px;align-items:center}.foot b{color:var(--maroon)}\n.foot .arrow{font-family:var(--disp);font-size:24px;color:var(--maroon)}\n.sec-glance,.sec-refl{color:var(--maroon)}.sec-fit{color:var(--fit)}.sec-fit .em{background:var(--fit-t)}.sec-fund{color:var(--fund)}.sec-fund .em{background:var(--fund-t)}.sec-skill{color:var(--skill)}.sec-skill .em{background:var(--skill-t)}.sec-ap{color:var(--ap)}.sec-ap .em{background:var(--ap-t)}.sec-rep{color:var(--rep)}.sec-rep .em{background:var(--rep-t)}\n\n/* overlay */\n.ov{position:fixed;inset:0;background:rgba(22,18,19,.55);display:none;align-items:center;justify-content:center;padding:20px;z-index:50}.ov.show{display:flex}\n.ovcard{background:var(--card);color:var(--ink);border-radius:18px;max-width:470px;width:100%;padding:22px 24px;box-shadow:0 20px 60px rgba(0,0,0,.3);animation:pop .16s ease;max-height:90vh;overflow:auto}\n@keyframes pop{from{transform:scale(.96);opacity:.4}to{transform:scale(1);opacity:1}}\n.ovcard .x{float:right;border:none;background:#f0eae9;width:30px;height:30px;border-radius:9px;font-size:16px;cursor:pointer;color:var(--ink-soft)}\n.ovcard h4{font-family:var(--disp);font-weight:700;font-size:22px;margin:2px 0 2px;display:flex;align-items:center;gap:7px}.ovcard .om{font-size:12.5px;color:var(--muted);margin:0 0 15px}\n.ovnow{background:var(--maroon-t);border-left:3px solid var(--maroon);border-radius:0 9px 9px 0;padding:11px 14px;margin-bottom:9px}\n.ovnow .l{font-family:var(--disp);font-size:12px;font-weight:600;color:var(--maroon);letter-spacing:.04em;margin-bottom:3px}.ovnow .t{font-size:13.5px;color:var(--ink);line-height:1.5}\n.ovnext{border:1px dashed var(--line);border-radius:9px;padding:11px 14px}.ovnext .l{font-family:var(--disp);font-size:12px;font-weight:600;color:var(--ink-soft);letter-spacing:.04em;margin-bottom:3px}.ovnext .t{font-size:13px;color:var(--ink-soft);line-height:1.5}\n.lad{border:1px solid var(--line);border-radius:9px;padding:9px 12px;margin-bottom:6px;font-size:12.5px;color:var(--ink-soft);display:flex;gap:9px;align-items:flex-start}\n.lad .lv{font-family:var(--disp);font-weight:700;font-size:13px;color:var(--muted);flex:none;width:20px}\n.ovfeeds{font-size:12px;color:var(--muted);margin-top:14px;border-top:1px solid var(--line);padding-top:10px}.ovfeeds b{color:var(--ink-soft)}\n\n/* ── card view (dark) ───────────────────────────────────────────── */\n.back{display:inline-flex;align-items:center;gap:6px;color:#c8bcbe;text-decoration:none;font-size:13px;background:none;border:1px solid #3a2e31;border-radius:999px;padding:6px 14px;cursor:pointer}\n.back:hover{border-color:#5a4a4e;color:#fff}\n.hero{text-align:center;margin:10px 0 22px}\n.hero .k{font-family:var(--disp);font-weight:700;letter-spacing:.22em;font-size:13px;color:var(--gold)}\n.hero h1{font-family:var(--disp);font-weight:800;font-size:clamp(30px,5vw,52px);margin:6px 0 4px;letter-spacing:-.01em}\n.hero p{color:#c8bcbe;font-size:14px;margin:0}\n.grid{display:grid;grid-template-columns:1.05fr .95fr;gap:34px;align-items:start}\n@media(max-width:820px){.grid{grid-template-columns:1fr}}\n.step{font-family:var(--disp);font-weight:700;font-size:13px;letter-spacing:.14em;color:var(--gold);text-transform:uppercase;margin-bottom:12px;display:flex;align-items:center;gap:10px}\n.step .saved{font-size:11px;letter-spacing:.04em;color:#7fe0a8;text-transform:none;font-weight:600}\n.styles{display:grid;grid-template-columns:1fr 1fr;gap:11px}\n@media(max-width:420px){.styles{grid-template-columns:1fr}}\n.sbtn{background:#241d20;border:1.5px solid #3a2e31;border-radius:14px;padding:15px 14px;cursor:pointer;transition:.15s;text-align:left;color:#fff;width:100%}\n.sbtn:hover{border-color:#5a4a4e;transform:translateY(-1px)}\n.sbtn:focus-visible{outline:2px solid var(--gold);outline-offset:2px}\n.sbtn.sel{border-color:var(--acc);background:#2e2528;box-shadow:0 6px 18px rgba(231,182,75,.2)}\n.sbtn[disabled]{opacity:.6;cursor:progress}\n.sbtn .se{font-size:24px}.sbtn .sn{font-family:var(--disp);font-weight:700;font-size:16px;margin:5px 0 2px}\n.sbtn .sd{font-size:11.5px;color:#c8bcbe;line-height:1.35}\n.cardstage{display:flex;flex-direction:column;align-items:center;gap:14px;position:sticky;top:18px}\n.card{width:min(360px,100%);border-radius:24px;padding:4px;background:var(--frame);box-shadow:0 26px 64px rgba(0,0,0,.55),0 0 30px rgba(231,182,75,.35);animation:float 5s ease-in-out infinite}\n@keyframes float{0%,100%{transform:translateY(0) rotate(-.4deg)}50%{transform:translateY(-8px) rotate(.4deg)}}\n.cinner{background:linear-gradient(160deg,var(--ink1),var(--ink2) 42%,var(--ink3));border-radius:20px;padding:17px 19px;position:relative;overflow:hidden;transition:.3s}\n.cinner::after{content:\"\";position:absolute;top:-60%;left:-75%;width:55%;height:220%;background:linear-gradient(115deg,transparent,rgba(255,255,255,.42),transparent);transform:rotate(9deg);animation:sweep 4.2s ease-in-out infinite;pointer-events:none}\n.chd{display:flex;justify-content:space-between;align-items:center;font-family:var(--disp);font-weight:700;font-size:11px;color:var(--acc);position:relative;z-index:2}\n.ctier{font-family:var(--disp);font-weight:700;font-size:10px;color:var(--rib);background:var(--frame);padding:2px 9px;border-radius:20px}\n.cwho{display:flex;align-items:center;gap:10px;margin-top:10px;position:relative;z-index:2}\n.cav{width:44px;height:44px;border-radius:50%;overflow:hidden;border:2px solid var(--acc);background:rgba(0,0,0,.25);display:grid;place-items:center;font-family:var(--disp);font-weight:700;font-size:17px;color:var(--accs);flex:none}\n.cav img{width:100%;height:100%;object-fit:cover;display:block}\n.cname{font-family:var(--disp);font-weight:800;font-size:24px;line-height:1;text-transform:uppercase;word-break:break-word;color:#fff}\n.cclass{font-size:11px;color:rgba(255,255,255,.72);margin-top:2px}\n.crib{background:linear-gradient(90deg,var(--acc),var(--accs));border-radius:9px;padding:6px 13px;font-family:var(--disp);font-weight:700;color:var(--rib);margin:11px 0;display:flex;justify-content:space-between;align-items:center;font-size:18px;position:relative;z-index:2}\n.crad{position:relative;z-index:2;margin:2px 0}\n.rlabel{font-family:var(--disp);font-size:9px;fill:rgba(255,255,255,.82);font-weight:600}\n.rnote{font-family:var(--body);font-size:10px;fill:rgba(255,255,255,.6)}\n.cstat{display:flex;gap:7px;margin-top:6px;position:relative;z-index:2}\n.cstat div{flex:1;background:rgba(0,0,0,.24);border:1px solid rgba(255,255,255,.18);border-radius:9px;padding:8px;text-align:center}\n.cstat b{font-family:var(--disp);font-size:18px;display:block;color:#fff}.cstat span{font-size:8px;color:rgba(255,255,255,.72);text-transform:uppercase;letter-spacing:.05em}\n.cbadge{display:flex;gap:9px;justify-content:center;margin-top:11px;position:relative;z-index:2}\n.cbadge span{width:34px;height:34px;border-radius:50%;background:radial-gradient(circle at 40% 35%,#fff3d6,var(--acc) 62%,#b9871f);display:grid;place-items:center;font-size:15px;opacity:.45;filter:grayscale(.4)}\n.cfoot{margin-top:11px;text-align:center;font-family:var(--disp);font-weight:600;font-size:11px;color:var(--acc);position:relative;z-index:2;border-top:1px solid rgba(255,255,255,.18);padding-top:8px}\n.cardhint{font-size:12px;color:#c8bcbe;text-align:center;max-width:34ch;line-height:1.45}\n.done{margin-top:34px;text-align:center}\n.done .ct-btn{margin:0}\n\n/* ── messages for non-student states ────────────────────────────── */\n.msg{max-width:520px;margin:60px auto;background:#fff;border:1px solid var(--line);border-radius:18px;padding:26px 28px;text-align:center}\n.msg .big{font-size:38px}\n.msg h2{font-family:var(--disp);font-weight:800;font-size:24px;margin:8px 0 6px}\n.msg p{color:var(--muted);font-size:14px;line-height:1.5;margin:0 0 8px}\n.msg code{background:var(--paper);padding:2px 7px;border-radius:6px;font-size:13px;color:var(--ink)}\n.msg .btn{display:inline-block;margin-top:10px;background:linear-gradient(145deg,#f3d689,#c9982f);color:#3a1008;border:none;font-family:var(--disp);font-weight:700;font-size:14px;padding:10px 20px;border-radius:11px;cursor:pointer;text-decoration:none}\n.small{font-size:11.5px;color:var(--muted);text-align:center;margin-top:30px;line-height:1.5}\n</style>\n",
  Rubrics: "<script>\n// GENERATED by dev/build-rubrics.js from docs/G8_PE_Rubric_Bank.md — do not edit.\nwindow.RUBRICS = {\n \"ladders\": {\n  \"throw\": {\n   \"name\": \"Throw\",\n   \"levels\": [\n    \"Short throw to a partner, stationary — step and follow through\",\n    \"Accurate overarm to a target over distance\",\n    \"Throw on the move, or vary power and height\",\n    \"Accurate under pressure or off-balance; disguise the throw\"\n   ]\n  },\n  \"catch\": {\n   \"name\": \"Catch\",\n   \"levels\": [\n    \"Large ball, two hands, gentle throw, still\",\n    \"Smaller ball, two hands, on the move\",\n    \"One-handed, or while changing direction\",\n    \"Under pressure — rebound, at speed, or with a distraction\"\n   ]\n  },\n  \"strike\": {\n   \"name\": \"Strike\",\n   \"levels\": [\n    \"Strike a stationary ball with control\",\n    \"Strike a moving/tossed ball with accuracy\",\n    \"Strike with direction and power to a target\",\n    \"Under pressure — vary placement, beat an opponent\"\n   ]\n  },\n  \"dribble\": {\n   \"name\": \"Dribble\",\n   \"levels\": [\n    \"Control at walking pace, straight line\",\n    \"Jogging pace, change direction round cones\",\n    \"At speed, close control, both sides/hands\",\n    \"Under pressure — protect the ball, beat a defender, head up\"\n   ]\n  },\n  \"kick\": {\n   \"name\": \"Kick\",\n   \"levels\": [\n    \"Stationary ball to a partner, basic control\",\n    \"Accurate to a target over distance (instep)\",\n    \"Kick a moving ball; use both feet\",\n    \"Under pressure — pass/shoot on the move, vary power and placement\"\n   ]\n  },\n  \"balance\": {\n   \"name\": \"Balance\",\n   \"levels\": [\n    \"Hold a still balance, two/one foot, a few seconds\",\n    \"One foot eyes closed, or different shapes\",\n    \"Move into and out of balances with control\",\n    \"Under challenge — narrow base, added movement, on apparatus\"\n   ]\n  },\n  \"agility\": {\n   \"name\": \"Agility\",\n   \"levels\": [\n    \"Change direction at jogging pace with control\",\n    \"Change direction at speed through a course\",\n    \"React and change direction to a signal\",\n    \"Under pressure — dodge an opponent, unpredictable cue\"\n   ]\n  },\n  \"jump\": {\n   \"name\": \"Jump & land\",\n   \"levels\": [\n    \"Jump and land on two feet, balanced (absorb)\",\n    \"Jump for distance/height, land under control\",\n    \"One-foot take-off / landing, in combinations\",\n    \"Jump-land-link into the next move under control\"\n   ]\n  },\n  \"core\": {\n   \"name\": \"Core stability\",\n   \"levels\": [\n    \"Hold a plank with good form for a short time\",\n    \"Hold longer, or a side plank, form intact\",\n    \"Keep a stable core while a limb moves (bird-dog, dead-bug)\",\n    \"Keep core control under load or in dynamic movement\"\n   ]\n  }\n }\n};\n<\/script>\n",
  App: "<script>\n(function () {\n'use strict';\n\n// ═══════════════ utils ═══════════════\nconst esc = s => String(s == null ? '' : s).replace(/[&<>\"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;' }[c]));\nconst $ = (sel, root) => (root || document).querySelector(sel);\nfunction toast(msg, err) { const t = $('#toast'); t.textContent = msg; t.className = 'toast show' + (err ? ' err' : ''); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 2400); }\n\n// ═══════════════ server bridge ═══════════════\nfunction rawCall(fn, ...args) {\n  return new Promise((resolve, reject) => {\n    if (!window.google || !google.script || !google.script.run) return reject(new Error('Not running inside the Apps Script web app'));\n    google.script.run.withSuccessHandler(resolve).withFailureHandler(e => reject(new Error(e && e.message ? e.message : String(e)))) [fn](...args);\n  });\n}\n// Apps Script caps simultaneous runs; a whole class tapping at once produces\n// transient \"Busy\"/timeout errors. Retry those with backoff; permanent ones surface.\nconst PERMANENT = /Unknown style|not signed in|class list|Not running inside/i;\nconst sleep = ms => new Promise(r => setTimeout(r, ms));\nasync function call(fn, ...args) {\n  let delay = (window.__preview && window.__preview.fastRetry) ? 30 : 1200;\n  for (let attempt = 0; ; attempt++) {\n    try { return await rawCall(fn, ...args); }\n    catch (e) { if (PERMANENT.test(e.message) || attempt >= 4) throw e; await sleep(delay + Math.random() * delay); delay = Math.min(delay * 2, 15000); }\n  }\n}\n\n// ═══════════════ PE Styles (visual skins; the server holds the names) ═══════════════\nconst SKINS = {\n  competitor: { e: '🏆', d: 'I love to compete and give everything.',          ink: ['#8f2130', '#7a1b28', '#4d0f18'], acc: '#f0c24a', accs: '#ffe08a', rib: '#3a1008' },\n  team:       { e: '🤝', d: 'I’m at my best lifting the team up.',         ink: ['#12726a', '#0e615a', '#063d38'], acc: '#8fe0c8', accs: '#c9f5e8', rib: '#053d36' },\n  improver:   { e: '📈', d: 'I’m all about getting better than I was.',    ink: ['#b5560e', '#9c4a0c', '#5e2c06'], acc: '#f3c07a', accs: '#ffe0b8', rib: '#4a2200' },\n  explorer:   { e: '🧭', d: 'I love trying new activities and challenges.',    ink: ['#22508c', '#1d4577', '#0f2747'], acc: '#8fbff0', accs: '#cfe4fb', rib: '#0c2747' },\n  energizer:  { e: '⚡', d: 'I bring energy and fun to every session.',        ink: ['#a3236b', '#8c1d5b', '#520f36'], acc: '#f7b0d8', accs: '#ffd6ec', rib: '#4a0f30' },\n  thinker:    { e: '🧠', d: 'I understand the why and think tactically.',      ink: ['#5a3b9c', '#4d3288', '#2c1c52'], acc: '#c3a6f0', accs: '#e2d3fb', rib: '#241452' }\n};\nconst ORDER = ['competitor', 'team', 'improver', 'explorer', 'energizer', 'thinker'];\nconst DEFAULT_SKIN = { e: '✨', ink: ['#7d2233', '#6f1d2c', '#4a1019'], acc: '#e7b64b', accs: '#f3d689', rib: '#3a1008' };\nfunction hexA(h, a) { h = h.replace('#', ''); return 'rgba(' + parseInt(h.slice(0, 2), 16) + ',' + parseInt(h.slice(2, 4), 16) + ',' + parseInt(h.slice(4, 6), 16) + ',' + a + ')'; }\nfunction setSkin(k) {\n  const s = SKINS[k] || DEFAULT_SKIN, r = document.documentElement.style;\n  r.setProperty('--ink1', s.ink[0]); r.setProperty('--ink2', s.ink[1]); r.setProperty('--ink3', s.ink[2]);\n  r.setProperty('--acc', s.acc); r.setProperty('--accs', s.accs); r.setProperty('--rib', s.rib);\n}\n\n// ═══════════════ radar (placeholder shape until Combine data exists — v3) ═══════════════\nconst AX = ['STAMINA', 'STRENGTH', 'SPEED', 'SKILLS', 'TEAM', 'EFFORT'], PLACE = [.62, .55, .6, .55, .7, .72];\nconst cx = 150, cy = 90, R = 60;\nfunction pt(i, v) { const a = (-90 + i * 60) * Math.PI / 180; return [cx + R * v * Math.cos(a), cy + R * v * Math.sin(a)]; }\nfunction radar(skin) {\n  let g = '';\n  for (let r = 1; r <= 4; r++) { const p = []; for (let i = 0; i < 6; i++) p.push(pt(i, r / 4).join(',')); g += '<polygon points=\"' + p.join(' ') + '\" fill=\"none\" stroke=\"rgba(255,255,255,.14)\"/>'; }\n  for (let i = 0; i < 6; i++) { const [x, y] = pt(i, 1.3); const an = Math.abs(x - cx) < 5 ? 'middle' : (x > cx ? 'start' : 'end'); g += '<text x=\"' + x.toFixed(1) + '\" y=\"' + (y + 3).toFixed(1) + '\" text-anchor=\"' + an + '\" class=\"rlabel\">' + AX[i] + '</text>'; }\n  g += '<polygon points=\"' + PLACE.map((v, i) => pt(i, v).map(n => n.toFixed(1)).join(',')).join(' ') + '\" fill=\"' + hexA(skin.acc, .3) + '\" stroke=\"' + skin.accs + '\" stroke-width=\"2.2\"/>';\n  g += '<text x=\"150\" y=\"187\" text-anchor=\"middle\" class=\"rnote\">your shape fills in at the Combine</text>';\n  return '<svg viewBox=\"0 0 300 190\" width=\"100%\" height=\"160\" xmlns=\"http://www.w3.org/2000/svg\" role=\"img\" aria-label=\"Six-sided fitness and skills chart, placeholder until the Combine\">' + g + '</svg>';\n}\n\n// ═══════════════ profile content (v1: everything is \"coming\"; data arrives in v2–v4) ═══════════════\nconst FUND = [['throw', '🎯'], ['catch', '🧤'], ['strike', '🏏'], ['dribble', '⛹️'], ['kick', '⚽'], ['balance', '🤸'], ['agility', '🏃'], ['jump', '⬆️'], ['core', '🧱']];\nconst D = {\n  'gl-fit':   { e: '🏃', t: 'Fitness', m: 'Combine trend', what: 'Five measures of your physical capacities: the mile, push-ups, sit-to-stand, broad jump and the 40 m sprint.', when: 'Your first numbers come from the September Combine. Then you chase your own records in January and May.' },\n  'gl-fund':  { e: '🧩', t: 'Fundamental skills', m: 'Nine building blocks, on the Combine ladder', what: 'Throw, catch, strike, dribble, kick, balance, agility, jump & land, core stability. Each is a 4-level ladder: stationary → on the move → isolated challenge → under pressure.', when: 'Tested identically every Combine, so your level is a trend. Cleared self → peer → teacher.' },\n  'gl-ap':    { e: '🤝', t: 'Participation', m: 'Your biggest all-year story', what: 'How you show up: five keys, tracked through every unit.', when: 'You rate yourself at the start and end of the year; your teacher checks in each unit.' },\n  'gl-know':  { e: '🧠', t: 'Understanding', m: 'Using what you know in new games', what: 'Knowing the why: fitness concepts, safe practice, and taking a tactic from one game into another.', when: 'Grows through Fitness, Parkour and especially the Tactics Cup.' },\n  'fit-cv':   { e: '🫀', t: 'Cardiovascular endurance', m: 'The mile run · Sept → Jan → May', what: 'Heart and lung staying power. The one everybody can move.', when: 'First mile at the September Combine.' },\n  'fit-me':   { e: '💪', t: 'Muscular endurance', m: 'Push-ups (upper) & sit-to-stand (lower)', what: 'How many quality reps in 60 seconds. Sit-ups are retired; core stability is its own skill ladder.', when: 'September Combine.' },\n  'fit-pw':   { e: '💥', t: 'Power', m: 'Standing broad jump', what: 'Explosive strength, measured in centimetres.', when: 'September Combine.' },\n  'fit-sp':   { e: '⚡', t: 'Speed', m: '40 m sprint', what: 'Acceleration and top speed, measured in seconds.', when: 'September Combine.' },\n  'sk-net':   { e: '🏸', t: 'Net & wall', m: 'Badminton · Volleyball', what: 'Serve with purpose, sustain and vary a rally, read play and position early.', when: 'Assessed inside the Badminton unit.' },\n  'sk-inv':   { e: '🏀', t: 'Invasion games', m: 'Basketball · Football · Tchoukball', what: 'Pass, support, find space, and contribute to the team.', when: 'Assessed inside Tchoukball and the Tactics Cup.' },\n  'sk-str':   { e: '⚾', t: 'Striking & target', m: 'Accuracy & contact', what: 'Solid contact, reliable throw and catch, accuracy at a target.', when: 'Assessed inside the striking unit.' },\n  'ap-ready': { e: '✅', t: 'Be ready', m: 'Kit on, on time, set to learn', what: 'Prepared with full kit, on time, ready to start without reminders.', when: 'Every lesson, all year.' },\n  'ap-effort':{ e: '🔥', t: 'Best effort', m: 'Give it my all', what: 'Full effort and engagement the whole lesson, including in activities you don’t naturally enjoy.', when: 'Every lesson, all year.' },\n  'ap-growth':{ e: '🌱', t: 'Growth mindset', m: 'Keep trying; learn from setbacks', what: 'Embrace challenges, try new things, stay positive after setbacks.', when: 'Every lesson, all year.' },\n  'ap-team':  { e: '🙌', t: 'Great teammate', m: 'Include, encourage, communicate', what: 'Communicate positively, encourage others, help teammates take part.', when: 'Every lesson, all year.' },\n  'ap-manage':{ e: '🧘', t: 'Manage myself & my feelings', m: 'Listen · safe environment · take responsibility', what: 'Manage behaviour and emotions positively without reminders.', when: 'Every lesson, all year.' },\n  'rp-know':  { e: '🧠', t: 'Knowledge', m: 'Understanding fitness, health & safety', what: 'Best fit across: fitness concepts (K4), safe practice (K5), health.', when: 'Reported on Canvas, 1–7.' },\n  'rp-skill': { e: '🎯', t: 'Skills', m: 'Performing movements & using them in games', what: 'Best fit across: fundamental skills (S1), game skills (S1/S3), planning (S2), team contribution (S4).', when: 'Reported on Canvas, 1–7.' },\n  'rp-ct':    { e: '🔄', t: 'Concept transfer', m: 'Using what you know in new sports & situations', what: 'Best fit across: Tactics Cup transfer (C2), meaningful participation (C1), health decisions (C3).', when: 'Reported on Canvas, 1–7.' }\n};\n\n// ═══════════════ views ═══════════════\nlet state = null, view = 'profile';\n\nfunction avatar(st, cls) {\n  const ini = esc(st.initials || '?');\n  if (st.photoUrl) return '<span class=\"' + cls + '\"><img src=\"' + esc(st.photoUrl) + '\" alt=\"\" referrerpolicy=\"no-referrer\" onerror=\"this.parentNode.textContent=\\'' + ini + '\\'\"></span>';\n  return '<span class=\"' + cls + '\">' + ini + '</span>';\n}\nfunction track(cls, n) { let s = ''; for (let i = 0; i < n; i++) s += '<i></i>'; return '<div class=\"track\">' + s + '</div>'; }\nfunction tile(cls, key, em, name, sub, lab) {\n  return '<button type=\"button\" class=\"tile ' + cls + '\" data-ov=\"' + key + '\"><span class=\"plus\">⊕</span><div class=\"tnrow\"><span class=\"em\">' + em + '</span><span class=\"tn\">' + esc(name) + '</span></div><div class=\"tsub\">' + esc(sub) + '</div>' + track(cls, 4) + '<div class=\"tracklab\">' + esc(lab) + '</div></button>';\n}\nfunction mtile(key, em, name, sub, unit) {\n  return '<button type=\"button\" class=\"tile a-fit\" data-ov=\"' + key + '\"><span class=\"plus\">⊕</span><div class=\"tnrow\"><span class=\"em\">' + em + '</span><span class=\"tn\">' + esc(name) + '</span></div><div class=\"tsub\">' + esc(sub) + '</div><div class=\"measure\">—<small>' + esc(unit) + '</small></div></button>';\n}\nfunction band(key, em, name) {\n  return '<button type=\"button\" class=\"band\" data-ov=\"' + key + '\"><span class=\"plus\">⊕</span><div class=\"bn\">' + em + ' ' + esc(name) + '</div><div class=\"btrack\"><i></i><i></i><i></i><i></i><i></i><i></i><i></i></div><div class=\"blab\">not yet reported</div></button>';\n}\n\nfunction profileView(b) {\n  const st = b.student, styles = b.styles, k = st.style, skin = SKINS[k];\n  const fund = FUND.map(([key, em]) => '<button type=\"button\" class=\"tile a-fund\" data-ov=\"fund-' + key + '\"><span class=\"plus\">⊕</span><div class=\"tnrow\"><span class=\"em\">' + em + '</span><span class=\"tn\">' + esc(RUBRICS.ladders[key].name) + '</span></div><div class=\"tsub\">building block</div>' + track('a-fund', 4) + '<div class=\"rung\">Level — of 4 · first Combine</div></button>').join('');\n  return '<div class=\"wrap\">' +\n    '<div class=\"topline\"><span class=\"demo-tag\">Hi ' + esc(st.first) + ' · your profile fills in through the year · tap any tile</span><span class=\"yearlab\">GRADE 8 PE · ' + esc(b.config.yearLabel) + '</span></div>' +\n    '<div class=\"cardteaser\">' + avatar(st, 'ct-av') +\n      '<div class=\"ct-id\"><div class=\"ct-name\">' + esc(st.name) + ' <span class=\"ct-tier\">◆ BRONZE</span></div>' + (st.klass ? '<div class=\"ct-sub\">' + esc(st.klass) + '</div>' : '') +\n      '<div class=\"ct-arch\">' + (skin ? skin.e + ' ' + esc(styles[k]).toUpperCase() : '✨ PICK YOUR PE STYLE') + '</div></div>' +\n      '<div class=\"ct-stats\"><div><b>—</b><span>Mile</span></div><div><b>—</b><span>Skills</span></div><div><b>—</b><span>Effort</span></div></div>' +\n      '<button type=\"button\" class=\"ct-btn' + (k ? '' : ' pulse') + '\" id=\"tocard\">' + (k ? 'View my card ↗' : 'Build my card ↗') + '</button>' +\n    '</div>' +\n    '<div class=\"goals\"><div class=\"hd\">⭐ Working towards</div><div class=\"empty\">Your goals for fitness, skills and participation appear here after the September Combine.</div></div>' +\n    '<h3 class=\"sec sec-glance\"><span class=\"em\">👀</span> At a glance <span class=\"tap\">· tap to open</span></h3><p class=\"seclede\">The whole picture in one look — position, not scores.</p>' +\n    '<div class=\"grid4\">' + tile('a-fit', 'gl-fit', '🏃', 'Fitness', 'from September', 'at the Combine') + tile('a-fund', 'gl-fund', '🧩', 'Fundamental skills', '9 building blocks', 'Level 1 to 4') + tile('a-ap', 'gl-ap', '🤝', 'Participation', 'all year', 'your five keys') + tile('a-skill', 'gl-know', '🧠', 'Understanding', 'use it in new games', 'Tactics Cup') + '</div>' +\n    '<h3 class=\"sec sec-fit\"><span class=\"em\">🏃</span> My fitness <span class=\"tap\">· tap a part</span></h3><p class=\"seclede\">Physical capacities — measured three times a year. Records that move.</p>' +\n    '<button type=\"button\" class=\"fhero\" data-ov=\"fit-cv\"><div class=\"fx\">🫀</div><div class=\"fm\"><div class=\"n\">Cardiovascular endurance</div><div class=\"s\">Heart &amp; lung staying power · the mile run</div></div><div class=\"fv\">—<span>September Combine</span></div></button>' +\n    '<div class=\"grid3\">' + mtile('fit-me', '💪', 'Muscular endurance', 'push-ups · sit-to-stand', '') + mtile('fit-pw', '💥', 'Power', 'broad jump', 'cm') + mtile('fit-sp', '⚡', 'Speed', '40m sprint', 's') + '</div>' +\n    '<h3 class=\"sec sec-fund\"><span class=\"em\">🧩</span> Fundamental skills <span class=\"tap\">· self → peer → teacher · tap a skill</span></h3><p class=\"seclede\">Nine building blocks, each a 4-level ladder. Same test every Combine, so your level climbs over the year.</p>' +\n    '<div class=\"grid3\">' + fund + '</div>' +\n    '<h3 class=\"sec sec-skill\"><span class=\"em\">🎯</span> Game skills <span class=\"tap\">· in real games</span></h3><p class=\"seclede\">The fundamentals, applied in a game. Assessed inside the units.</p>' +\n    '<div class=\"grid3\">' + tile('a-skill', 'sk-net', '🏸', 'Net & wall', 'badminton · volleyball', 'coming') + tile('a-skill', 'sk-inv', '🏀', 'Invasion', 'basketball · football', 'coming') + tile('a-skill', 'sk-str', '⚾', 'Striking & target', 'accuracy & contact', 'coming') + '</div>' +\n    '<h3 class=\"sec sec-ap\"><span class=\"em\">🤝</span> Active participation <span class=\"tap\">· all year</span></h3><p class=\"seclede\">Your biggest all-year story — how you show up, tracked through every unit.</p>' +\n    '<div class=\"grid3\">' + tile('a-ap', 'ap-ready', '✅', 'Be ready', 'set to learn', 'all year') + tile('a-ap', 'ap-effort', '🔥', 'Best effort', 'give it my all', 'all year') + tile('a-ap', 'ap-growth', '🌱', 'Growth mindset', 'learn from setbacks', 'all year') + tile('a-ap', 'ap-team', '🙌', 'Great teammate', 'include & encourage', 'all year') + tile('a-ap', 'ap-manage', '🧘', 'Manage myself & feelings', 'listen · safe space', 'all year') + '</div>' +\n    '<h3 class=\"sec sec-rep\"><span class=\"em\">🧠</span> My report picture <span class=\"tap\">· tap for how it’s built</span></h3><p class=\"seclede\">The three grades on your Canvas report — a best fit across the descriptors.</p>' +\n    '<div class=\"grid3\">' + band('rp-know', '🧠', 'Knowledge') + band('rp-skill', '🎯', 'Skills') + band('rp-ct', '🔄', 'Concept transfer') + '</div>' +\n    '<div class=\"foot\"><span class=\"arrow\">↻</span><div><b>How this grows:</b> September sets the baseline; the same measures repeat in January and May. Always <b>you vs. your past self</b> — and it maps straight to your report.</div></div>' +\n    (b.alsoTeacher ? '<p class=\"small\">You are on the Teachers tab too, so this is what a student sees.</p>' : '') +\n  '</div>';\n}\n\nfunction cardView(b) {\n  const st = b.student, styles = b.styles, chosen = st.style;\n  const pick = ORDER.filter(k => styles[k]).map(k => {\n    const s = SKINS[k];\n    return '<button type=\"button\" class=\"sbtn' + (k === chosen ? ' sel' : '') + '\" data-k=\"' + k + '\" aria-pressed=\"' + (k === chosen) + '\"><span class=\"se\">' + s.e + '</span><div class=\"sn\">' + esc(styles[k]) + '</div><div class=\"sd\">' + esc(s.d) + '</div></button>';\n  }).join('');\n  return '<div class=\"wrap\">' +\n    '<button type=\"button\" class=\"back\" id=\"toprofile\">← Back to my profile</button>' +\n    '<div class=\"hero\"><div class=\"k\">GRADE 8 PE · ' + esc(b.config.yearLabel) + '</div><h1>Hi ' + esc(st.first) + ', build your Athlete Card</h1><p>Choose your PE Style and make it yours. Your card grows all year.</p></div>' +\n    '<div class=\"grid\">' +\n      '<div><div class=\"step\">1 · Choose your style <span class=\"saved\" id=\"saved\">' + (chosen ? '✓ saved' : '') + '</span></div><div class=\"styles\" id=\"styles\">' + pick + '</div></div>' +\n      '<div class=\"cardstage\"><div class=\"card\" id=\"card\"><div class=\"cinner\">' +\n        '<div class=\"chd\"><span>FIS · PE ATHLETE · ' + esc(b.config.yearLabel) + '</span><span class=\"ctier\">BRONZE</span></div>' +\n        '<div class=\"cwho\">' + avatar(st, 'cav') + '<div><div class=\"cname\">' + esc(st.name) + '</div>' + (st.klass ? '<div class=\"cclass\">' + esc(st.klass) + '</div>' : '') + '</div></div>' +\n        '<div class=\"crib\"><span id=\"archt\">' + (chosen ? esc(styles[chosen]).toUpperCase() : 'PICK A STYLE') + '</span><span id=\"arche\">' + (chosen ? SKINS[chosen].e : '✨') + '</span></div>' +\n        '<div class=\"crad\" id=\"rad\">' + radar(SKINS[chosen] || DEFAULT_SKIN) + '</div>' +\n        '<div class=\"cstat\"><div><b>—</b><span>Mile</span></div><div><b>—</b><span>Skills</span></div><div><b>—</b><span>Effort</span></div></div>' +\n        '<div class=\"cbadge\" aria-label=\"Badges you can earn\"><span>🏃</span><span>📈</span><span>🔥</span><span>⭐</span></div>' +\n        '<div class=\"cfoot\">◆ Your card levels up as you grow ◆</div>' +\n      '</div></div><div class=\"cardhint\">Your style is your choice. None is \"better\": it sets your card’s colour and identity.</div></div>' +\n    '</div>' +\n    '<div class=\"done\"><button type=\"button\" class=\"ct-btn\" id=\"toprofile2\">Done · back to my profile ↗</button></div>' +\n  '</div>';\n}\n\nfunction message(big, title, body, extra) {\n  return '<div class=\"msg\"><div class=\"big\">' + big + '</div><h2>' + esc(title) + '</h2><p>' + body + '</p>' + (extra || '') + '</div>';\n}\nfunction otherView(b) {\n  switch (b.role) {\n    case 'teacher':\n      return message('🧑‍🏫', 'Teacher view', 'You’re signed in as <code>' + esc(b.email) + '</code>, which is on the Teachers tab but not the Students tab. Students who open this link see their own profile.',\n        '<p>To see the student view yourself, add your address to the Students tab. Student data lives only in the Sheet.</p>');\n    case 'unknown':\n      return message('🔎', 'Not on the Grade 8 list yet', 'You’re signed in as <code>' + esc(b.email) + '</code>. That account isn’t on the Grade 8 PE class list.',\n        '<p>Tell your PE teacher, or check you opened the link with your FIS account.</p>');\n    case 'wrong_domain':\n      return message('🔒', 'Use your FIS account', 'This page only works with an FIS Google account.',\n        '<p>Sign out of other Google accounts, or open the link in a window where your FIS account is the active one.</p>');\n    default:\n      return message('🙈', 'Couldn’t confirm who you are', 'Google didn’t tell the page which account is signed in.',\n        '<p>Open the link from your FIS account (not incognito). If it keeps happening, tell your PE teacher: the deployment must be \"Anyone within FIS\".</p>');\n  }\n}\n\n// ═══════════════ overlay ═══════════════\nfunction openOv(k) {\n  let html = '<button class=\"x\" type=\"button\" data-close=\"1\">✕</button>';\n  if (k.startsWith('fund-')) {\n    const key = k.slice(5), lad = RUBRICS.ladders[key], em = (FUND.find(f => f[0] === key) || ['', '🧩'])[1];\n    html += '<h4>' + em + ' ' + esc(lad.name) + '</h4><div class=\"om\">Fundamental skill · 4-level ladder · cleared self → peer → teacher</div>';\n    lad.levels.forEach((t, i) => { html += '<div class=\"lad\"><span class=\"lv\">L' + (i + 1) + '</span><span>' + esc(t) + '</span></div>'; });\n    html += '<div class=\"ovfeeds\"><b>Not yet tested.</b> Your first level comes from the September Combine, then Jan → May.</div>';\n  } else {\n    const d = D[k]; if (!d) return;\n    html += '<h4>' + d.e + ' ' + esc(d.t) + '</h4><div class=\"om\">' + esc(d.m) + '</div>' +\n      '<div class=\"ovnow\"><div class=\"l\">WHAT THIS IS</div><div class=\"t\">' + esc(d.what) + '</div></div>' +\n      '<div class=\"ovnext\"><div class=\"l\">↗ WHEN IT FILLS IN</div><div class=\"t\">' + esc(d.when) + '</div></div>';\n  }\n  $('#ovcard').innerHTML = html; $('#ov').classList.add('show');\n}\nfunction closeOv() { $('#ov').classList.remove('show'); }\n\n// ═══════════════ interactions ═══════════════\nfunction applySkin(k) {\n  const s = SKINS[k]; if (!s) return;\n  setSkin(k);\n  $('#archt').textContent = state.styles[k].toUpperCase(); $('#arche').textContent = s.e;\n  $('#rad').innerHTML = radar(s);\n  document.querySelectorAll('.sbtn').forEach(b => { const on = b.dataset.k === k; b.classList.toggle('sel', on); b.setAttribute('aria-pressed', on); });\n}\nasync function choose(k) {\n  if (!SKINS[k] || k === state.student.style) { applySkin(k); return; }\n  const prev = state.student.style;\n  applySkin(k);\n  const saved = $('#saved'); saved.textContent = 'saving…';\n  document.querySelectorAll('.sbtn').forEach(b => b.disabled = true);\n  try {\n    const res = await call('saveStyle', k);\n    state.student.style = res.style; saved.textContent = '✓ saved';\n    toast('Saved. You’re ' + state.styles[k] + '.');\n  } catch (e) {\n    saved.textContent = prev ? '✓ saved' : '';\n    if (prev) applySkin(prev); else { setSkin(''); $('#archt').textContent = 'PICK A STYLE'; $('#arche').textContent = '✨'; $('#rad').innerHTML = radar(DEFAULT_SKIN); document.querySelectorAll('.sbtn').forEach(b => { b.classList.remove('sel'); b.setAttribute('aria-pressed', 'false'); }); }\n    toast('Couldn’t save: ' + e.message, true);\n  } finally { document.querySelectorAll('.sbtn').forEach(b => b.disabled = false); }\n}\n\nfunction show(v) {\n  view = v;\n  const app = $('#app');\n  document.body.classList.toggle('dark', v === 'card');\n  setSkin(state.student.style);\n  if (v === 'card') {\n    app.innerHTML = cardView(state);\n    $('#styles').addEventListener('click', e => { const btn = e.target.closest('.sbtn'); if (btn) choose(btn.dataset.k); });\n    $('#toprofile').addEventListener('click', () => show('profile'));\n    $('#toprofile2').addEventListener('click', () => show('profile'));\n  } else {\n    app.innerHTML = profileView(state);\n    $('#tocard').addEventListener('click', () => show('card'));\n    app.addEventListener('click', e => { const t = e.target.closest('[data-ov]'); if (t) openOv(t.dataset.ov); });\n  }\n  window.scrollTo(0, 0);\n}\n\nfunction render(b) {\n  state = b;\n  if (b.role !== 'student') { $('#app').innerHTML = otherView(b); return; }\n  show(b.student.style ? 'profile' : 'card'); // first visit goes straight to \"build your card\"\n}\n\n$('#ov').addEventListener('click', e => { if (e.target === e.currentTarget || e.target.closest('[data-close]')) closeOv(); });\ndocument.addEventListener('keydown', e => { if (e.key === 'Escape') closeOv(); });\n\nasync function main() {\n  try { render(await call('bootstrap')); }\n  catch (e) { $('#app').innerHTML = message('⚠️', 'Couldn’t load your profile', esc(e.message), '<p><a class=\"btn\" href=\"javascript:location.reload()\">Try again</a></p>'); }\n}\nif (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', main); else main();\n})();\n<\/script>\n"
};
