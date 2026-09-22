#!/usr/bin/env node
// Runs the real src/Code.gs under Node against the fake runtime and checks the
// privacy guarantees the brief asks for: a student gets their own row only,
// a non-domain account is refused, writes are validated and keyed by the
// verified identity. Synthetic data only.   Usage: node dev/test-server.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.join(__dirname, '..');
const read = p => fs.readFileSync(path.join(root, p), 'utf8');

vm.runInThisContext(read('dev/fake-sheets.js'), { filename: 'fake-sheets.js' });
vm.runInThisContext(read('src/Code.gs'), { filename: 'Code.gs' });

let failed = 0;
function check(name, cond, detail) { if (cond) console.log('  ✓', name); else { failed++; console.log('  ✗', name, detail === undefined ? '' : JSON.stringify(detail)); } }
function as(email, fn) { const prev = FakeSheets.user; FakeSheets.user = email; try { return fn(); } finally { FakeSheets.user = prev; } }
function throws(fn) { try { fn(); return null; } catch (e) { return e.message; } }

// ---- seed a synthetic school ----
FakeSheets.owner = 'teacher@example.edu';
FakeSheets.user = FakeSheets.owner;
setupTabs();
const cfgTab = FakeSheets.book.getSheetByName('Config');
cfgTab.getDataRange().getValues().forEach((r, i) => { if (r[0] === 'domain') cfgTab.getRange(i + 1, 2).setValues([['example.edu']]); });
FakeSheets.book.getSheetByName('Students').getRange(2, 1, 4, 3).setValues([
  ['Sample.One@example.edu', 'Sample One', '8A'],
  ['sample.two@example.edu', 'Two, Sample', '8A'],
  ['sample.three@example.edu', 'T, Sample Jo', '8B'],
  ['teacher@example.edu', 'Teacher Test', '8B']
]);
FakeSheets.book.getSheetByName('Teachers').getRange(3, 1, 1, 2).setValues([['other.teacher@example.edu', 'Other Teacher']]);
clearConfigCache();

console.log('setup');
check('tabs created', ['Config', 'Students', 'Teachers', 'Profile'].every(n => FakeSheets.book.getSheetByName(n)));
check('owner auto-added to Teachers', readTab_('Teachers')[0].Email === 'teacher@example.edu');
check('config defaults + override', config_().domain === 'example.edu' && config_().web_fonts === 'FALSE' && config_().photo_lookup === 'FALSE');

console.log('identity');
const s1 = as('sample.one@example.edu', () => bootstrap());
check('student (case-insensitive email) → own name & class', s1.role === 'student' && s1.student.name === 'Sample One' && s1.student.klass === '8A', s1);
check('student payload has no email and no roster', !('email' in s1) && !('roster' in s1) && !JSON.stringify(s1).includes('example.edu'), s1);
check('student initials', s1.student.initials === 'SO');
check('"First Last" → first + full display', s1.student.first === 'Sample' && s1.student.name === 'Sample One');
check('"Last, First" → first name only, initial ignored', as('sample.two@example.edu', () => { const st = bootstrap().student; return st.first === 'Sample' && st.name === 'Sample' && st.initials === 'S'; }));
check('"L, First Middle" → first token after the comma', as('sample.three@example.edu', () => { const st = bootstrap().student; return st.first === 'Sample' && st.name === 'Sample'; }));
check('nameParts_ edge cases', nameParts_('  ').first === '' && nameParts_('Solo').display === 'Solo' && nameParts_(',').first === '' );
check('no style yet → empty', s1.student.style === '' && s1.student.photoUrl === '');
check('owner on roster → student + alsoTeacher', as('teacher@example.edu', () => { const b = bootstrap(); return b.role === 'student' && b.alsoTeacher === true; }));
check('teacher not on roster → teacher view with own email only', as('other.teacher@example.edu', () => { const b = bootstrap(); return b.role === 'teacher' && b.email === 'other.teacher@example.edu' && !b.student; }));
check('domain account not on any tab → unknown', as('nobody@example.edu', () => bootstrap().role === 'unknown'));
check('other domain → wrong_domain, nothing else returned', as('sample.one@example.edu.evil.test', () => { const b = bootstrap(); return b.role === 'wrong_domain' && !b.email && !b.student; }));
check('gmail-style address → wrong_domain', as('someone@gmail.example', () => bootstrap().role === 'wrong_domain'));
check('no account → anonymous', as('', () => bootstrap().role === 'anonymous'));

console.log('writes');
check('saveStyle stores for the caller', as('sample.one@example.edu', () => saveStyle('Improver').style === 'improver' && bootstrap().student.style === 'improver'));
check('Profile row keyed by verified email', readTab_('Profile').length === 1 && readTab_('Profile')[0].Email === 'sample.one@example.edu');
check('saveStyle upserts (no duplicate rows)', as('sample.one@example.edu', () => { saveStyle('team'); return readTab_('Profile').length === 1 && readTab_('Profile')[0].Style === 'team'; }));
check('other student unaffected and cannot see it', as('sample.two@example.edu', () => bootstrap().student.style === ''));
check('unknown style rejected', /Unknown style/.test(throws(() => as('sample.one@example.edu', () => saveStyle('ninja')))));
check('teacher (not on roster) cannot write', /class list/.test(throws(() => as('other.teacher@example.edu', () => saveStyle('team')))));
check('wrong domain cannot write', /class list/.test(throws(() => as('x@gmail.example', () => saveStyle('team')))));
check('anonymous cannot write', /class list/.test(throws(() => as('', () => saveStyle('team')))));
check('error messages contain no email', !throws(() => as('x@gmail.example', () => saveStyle('team'))).includes('@'));
check('saveGoal trims, strips control chars, caps length', as('sample.two@example.edu', () => { const g = saveGoal('  Mile\tunder 8:30 ' + 'x'.repeat(300)); return g.goal.length === 140 && g.goal.startsWith('Mile under 8:30'); }));

console.log('predictions');
const full = {}; PREDICT_ITEMS.forEach((k, i) => { full[k] = { rating: RATINGS[i % 3], freq: FREQS[i % 3] }; });
check('Predictions tab has 29 columns', TABS.Predictions.length === 29 && TABS.Predictions[3] === 'cv_rating' && TABS.Predictions[28] === 'core_freq');
check('no prediction yet → empty object, combine null', as('sample.one@example.edu', () => { const st = bootstrap().student; return JSON.stringify(st.predictions) === '{}' && st.combine === null; }));
check('bootstrap exposes the item/rating vocab', as('sample.one@example.edu', () => { const b = bootstrap(); return b.predict.items.length === 13 && b.predict.ratings.length === 3 && b.predict.freqs.length === 3; }));
check('incomplete submission rejected, nothing written', /Incomplete.*12 left/.test(throws(() => as('sample.one@example.edu', () => savePrediction('intro', { cv: { rating: 'strength', freq: 'often' } })))) && readTab_('Predictions').length === 0);
check('bad value rejected', /Incomplete/.test(throws(() => as('sample.one@example.edu', () => { const bad = JSON.parse(JSON.stringify(full)); bad.kick.rating = 'amazing'; return savePrediction('intro', bad); }))));
check('unknown checkpoint rejected', /Unknown checkpoint/.test(throws(() => as('sample.one@example.edu', () => savePrediction('final', full)))));
check('full submission saved and echoed', as('sample.one@example.edu', () => { const r = savePrediction('Intro', full); return r.checkpoint === 'intro' && r.prediction.items.cv.rating === 'strength' && r.prediction.items.core.freq === FREQS[12 % 3] && r.prediction.timestamp.length >= 10; }));
check('one row, keyed by verified email + checkpoint', readTab_('Predictions').length === 1 && readTab_('Predictions')[0].Email === 'sample.one@example.edu' && readTab_('Predictions')[0].Checkpoint === 'intro');
check('re-save overwrites the same row', as('sample.one@example.edu', () => { const again = JSON.parse(JSON.stringify(full)); again.cv.rating = 'work-on'; savePrediction('intro', again); return readTab_('Predictions').length === 1 && bootstrap().student.predictions.intro.items.cv.rating === 'work-on'; }));
check('another student sees no prediction', as('sample.two@example.edu', () => JSON.stringify(bootstrap().student.predictions) === '{}'));
check('teacher / wrong domain / anonymous cannot save', [ 'other.teacher@example.edu', 'x@gmail.example', '' ].every(e => /class list/.test(throws(() => as(e, () => savePrediction('intro', full))))));
check('client-supplied email in answers is ignored', as('sample.two@example.edu', () => { const sneaky = JSON.parse(JSON.stringify(full)); sneaky.Email = 'sample.one@example.edu'; savePrediction('intro', sneaky); const rows = readTab_('Predictions'); return rows.length === 2 && rows.filter(r => r.Email === 'sample.two@example.edu').length === 1; }));
check('bootstrap payload never carries another student\'s rows', as('sample.two@example.edu', () => !JSON.stringify(bootstrap()).includes('sample.one')));

console.log('photo lookup');
check('off by default → no People call', as('sample.one@example.edu', () => bootstrap().student.photoUrl === ''));
cfgTab.getDataRange().getValues().forEach((r, i) => { if (r[0] === 'photo_lookup') cfgTab.getRange(i + 1, 2).setValues([['TRUE']]); });
clearConfigCache();
check('on, but People service missing → initials fallback, no throw', as('sample.one@example.edu', () => bootstrap().student.photoUrl === ''));
globalThis.People = { People: { searchDirectoryPeople: o => ({ people: [{ photos: [{ url: 'https://photo.example/' + encodeURIComponent(o.query), default: false }] }] }) } };
check('on, directory returns a photo → url passed through, nothing stored', as('sample.one@example.edu', () => bootstrap().student.photoUrl === 'https://photo.example/sample.one%40example.edu') && !JSON.stringify(FakeSheets.book.toJSON()).includes('photo.example'));
delete globalThis.People;
globalThis.AdminDirectory = { Users: { get: (email, o) => ({ thumbnailPhotoUrl: 'https://photo.example/dir/' + encodeURIComponent(email) + '/' + o.viewType, isDefaultPhoto: false }) } };
check('People off, Admin Directory on → domain_public thumbnail used', as('sample.one@example.edu', () => bootstrap().student.photoUrl === 'https://photo.example/dir/sample.one%40example.edu/domain_public'));
globalThis.AdminDirectory = { Users: { get: () => ({ thumbnailPhotoUrl: 'https://photo.example/default', isDefaultPhoto: true }) } };
check('Admin Directory default photo → initials', as('sample.one@example.edu', () => bootstrap().student.photoUrl === ''));
delete globalThis.AdminDirectory;

console.log('retention');
FakeSheets.uiAnswer = 'YES';
clearProfileData();
check('clearProfileData empties Profile + Predictions, keeps Students', readTab_('Profile').length === 0 && readTab_('Predictions').length === 0 && readTab_('Students').length === 4);

console.log('web app');
EMBEDDED_HTML.Index = read('src/Index.html');
const page = doGet().getContent();
check('doGet: no Google Fonts by default', !page.includes('fonts.googleapis.com'));
cfgTab.getDataRange().getValues().forEach((r, i) => { if (r[0] === 'web_fonts') cfgTab.getRange(i + 1, 2).setValues([['TRUE']]); });
clearConfigCache();
check('doGet: fonts link when web_fonts = TRUE', doGet().getContent().includes('fonts.googleapis.com'));

console.log(failed ? `\n${failed} check(s) FAILED` : '\nall checks passed');
process.exit(failed ? 1 : 0);
