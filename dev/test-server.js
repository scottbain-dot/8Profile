#!/usr/bin/env node
// Runs the real src/Code.gs under Node against the fake runtime and checks the
// privacy guarantees: identity comes only from a verified Google token, a
// student gets their own rows only, non-domain accounts are refused, writes are
// validated. Synthetic data only.   Usage: node dev/test-server.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.join(__dirname, '..');
const read = p => fs.readFileSync(path.join(root, p), 'utf8');

vm.runInThisContext(read('dev/fake-sheets.js'), { filename: 'fake-sheets.js' });
vm.runInThisContext(read('src/Code.gs'), { filename: 'Code.gs' });

let failed = 0;
function check(name, cond, detail) { if (cond) console.log('  ✓', name); else { failed++; console.log('  ✗', name, detail === undefined ? '' : JSON.stringify(detail)); } }
const CLIENT = 'test-client';
const now = () => Math.floor(Date.now() / 1000);
const tokenFor = (email, extra) => FakeSheets.token(Object.assign({ aud: CLIENT, email, email_verified: 'true', hd: email.split('@')[1], exp: now() + 3600, name: 'Tok Name', picture: 'https://lh3.googleusercontent.com/a/photo' }, extra || {}));
const post = (action, payload, tok) => JSON.parse(doPost({ postData: { contents: JSON.stringify(Object.assign({ action, token: tok }, payload || {})) } }).getContent());
const as = (email, action, payload, extra) => post(action, payload, tokenFor(email, extra));
const S1 = 'sample.one@example.edu', S2 = 'sample.two@example.edu', S3 = 'sample.three@example.edu', T = 'other.teacher@example.edu', OWNER = 'teacher@example.edu';

// ---- seed a synthetic school ----
FakeSheets.owner = OWNER;
setupTabs();
const cfgTab = FakeSheets.book.getSheetByName('Config');
const setConfig = (k, v) => { cfgTab.getDataRange().getValues().forEach((r, i) => { if (r[0] === k) cfgTab.getRange(i + 1, 2).setValues([[v]]); }); clearConfigCache(); };
setConfig('domain', 'example.edu'); setConfig('oauth_client_id', CLIENT);
FakeSheets.book.getSheetByName('Students').getRange(2, 1, 4, 3).setValues([
  ['Sample.One@example.edu', 'Sample One', '8A'],
  [S2, 'Two, Sample', '8A'],
  [S3, 'T, Sample Jo', '8B'],
  [OWNER, 'Teacher Test', '8B']
]);
FakeSheets.book.getSheetByName('Teachers').getRange(3, 1, 1, 2).setValues([[T, 'Other Teacher']]);

console.log('setup');
check('tabs created', ['Config', 'Students', 'Teachers', 'Profile', 'Predictions'].every(n => FakeSheets.book.getSheetByName(n)));
check('owner auto-added to Teachers', readTab_('Teachers')[0].Email === OWNER);
check('Predictions tab has 29 columns', TABS.Predictions.length === 29 && TABS.Predictions[3] === 'cv_rating' && TABS.Predictions[28] === 'core_freq');

console.log('token verification');
const authErr = r => !r.ok && r.authRequired === true;
check('no token → authRequired', authErr(post('bootstrap', {}, '')));
check('garbage token → authRequired, no tokeninfo call', (() => { const n = FakeSheets.fetchCount; const r = post('bootstrap', {}, 'not a token'); return authErr(r) && FakeSheets.fetchCount === n; })());
check('token google rejects → authRequired', authErr(post('bootstrap', {}, 'real.looking.jwt')));
check('token for another client (aud) → authRequired', authErr(as(S1, 'bootstrap', {}, { aud: 'someone-else' })));
check('unverified email → authRequired', authErr(as(S1, 'bootstrap', {}, { email_verified: 'false' })));
check('expired token → authRequired', authErr(as(S1, 'bootstrap', {}, { exp: now() - 5 })));
check('valid token → ok', as(S1, 'bootstrap').ok === true);
check('verified token is cached (second call makes no tokeninfo request)', (() => { const tok = tokenFor(S3); post('bootstrap', {}, tok); const n = FakeSheets.fetchCount; post('bootstrap', {}, tok); return FakeSheets.fetchCount === n; })());
check('unknown action → error, not authRequired', (() => { const r = as(S1, 'nothing'); return !r.ok && /Unknown action/.test(r.error) && !r.authRequired; })());
check('bad JSON body → Bad request', (() => { const r = JSON.parse(doPost({ postData: { contents: '{nope' } }).getContent()); return !r.ok && /Bad request/.test(r.error); })());
check('doGet returns no data', (() => { const r = JSON.parse(doGet().getContent()); return r.ok === false && !JSON.stringify(r).includes('example.edu'); })());

console.log('identity');
const s1 = as(S1, 'bootstrap').data;
check('student (case-insensitive roster email) → own name & class', s1.role === 'student' && s1.student.name === 'Sample One' && s1.student.klass === '8A', s1);
check('student payload has no email and no roster', !('email' in s1) && !JSON.stringify(s1).includes('example.edu'), s1);
check('photo comes from the token picture claim', s1.student.photoUrl === 'https://lh3.googleusercontent.com/a/photo');
check('non-Google picture URL is dropped', as(S1, 'bootstrap', {}, { picture: 'https://evil.example/x.png' }).data.student.photoUrl === '');
check('"Last, First" → first name only, initial ignored', (() => { const st = as(S2, 'bootstrap').data.student; return st.first === 'Sample' && st.name === 'Sample' && st.initials === 'S'; })());
check('"L, First Middle" → first token after the comma', (() => { const st = as(S3, 'bootstrap').data.student; return st.first === 'Sample' && st.name === 'Sample'; })());
check('nameParts_ edge cases', nameParts_('  ').first === '' && nameParts_('Solo').display === 'Solo' && nameParts_(',').first === '');
check('owner on roster → student + alsoTeacher', (() => { const b = as(OWNER, 'bootstrap').data; return b.role === 'student' && b.alsoTeacher === true; })());
check('teacher not on roster → teacher view with own email only', (() => { const b = as(T, 'bootstrap').data; return b.role === 'teacher' && b.email === T && !b.student; })());
check('domain account on no tab → unknown', as('nobody@example.edu', 'bootstrap').data.role === 'unknown');
check('other domain (hd) → wrong_domain, nothing else', (() => { const b = as('someone@gmail.example', 'bootstrap').data; return b.role === 'wrong_domain' && !b.email && !b.student; })());
check('lookalike domain → wrong_domain', as('x@example.edu.evil.test', 'bootstrap').data.role === 'wrong_domain');
check('right email but missing hd claim → wrong_domain', as(S1, 'bootstrap', {}, { hd: '' }).data.role === 'wrong_domain');
check('roster email does not grant access without a token for it', (() => { const r = post('bootstrap', { email: S1 }, tokenFor('nobody@example.edu')); return r.data.role === 'unknown'; })());

console.log('writes');
check('saveStyle stores for the caller', as(S1, 'saveStyle', { style: 'Improver' }).data.style === 'improver' && as(S1, 'bootstrap').data.student.style === 'improver');
check('Profile row keyed by verified email', readTab_('Profile').length === 1 && readTab_('Profile')[0].Email === S1);
check('saveStyle upserts (no duplicate rows)', (() => { as(S1, 'saveStyle', { style: 'team' }); return readTab_('Profile').length === 1 && readTab_('Profile')[0].Style === 'team'; })());
check('other student unaffected and cannot see it', as(S2, 'bootstrap').data.student.style === '');
check('unknown style rejected', /Unknown style/.test(as(S1, 'saveStyle', { style: 'ninja' }).error));
check('teacher (not on roster) cannot write', /class list/.test(as(T, 'saveStyle', { style: 'team' }).error));
check('wrong domain cannot write', /class list/.test(as('x@gmail.example', 'saveStyle', { style: 'team' }).error));
check('no token cannot write', authErr(post('saveStyle', { style: 'team' }, '')));
check('error messages contain no email', !as('x@gmail.example', 'saveStyle', { style: 'team' }).error.includes('@'));
check('saveGoal trims, strips control chars, caps length', (() => { const g = as(S2, 'saveGoal', { goal: '  Mile\tunder 8:30 ' + 'x'.repeat(300) }).data; return g.goal.length === 140 && g.goal.startsWith('Mile under 8:30'); })());

console.log('predictions');
const full = {}; PREDICT_ITEMS.forEach((k, i) => { full[k] = { rating: RATINGS[i % 3], freq: FREQS[i % 3] }; });
check('no prediction yet → empty object, combine null', (() => { const st = as(S1, 'bootstrap').data.student; return JSON.stringify(st.predictions) === '{}' && st.combine === null; })());
check('bootstrap exposes the item/rating vocab', (() => { const b = as(S1, 'bootstrap').data; return b.predict.items.length === 13 && b.predict.ratings.length === 3 && b.predict.freqs.length === 3; })());
check('incomplete submission rejected, nothing written', /Incomplete.*12 left/.test(as(S1, 'savePrediction', { checkpoint: 'intro', answers: { cv: { rating: 'strength', freq: 'often' } } }).error) && readTab_('Predictions').length === 0);
check('bad value rejected', (() => { const bad = JSON.parse(JSON.stringify(full)); bad.kick.rating = 'amazing'; return /Incomplete/.test(as(S1, 'savePrediction', { checkpoint: 'intro', answers: bad }).error); })());
check('unknown checkpoint rejected', /Unknown checkpoint/.test(as(S1, 'savePrediction', { checkpoint: 'final', answers: full }).error));
check('full submission saved and echoed', (() => { const r = as(S1, 'savePrediction', { checkpoint: 'Intro', answers: full }).data; return r.checkpoint === 'intro' && r.prediction.items.cv.rating === 'strength' && r.prediction.timestamp.length >= 10; })());
check('one row, keyed by verified email + checkpoint', readTab_('Predictions').length === 1 && readTab_('Predictions')[0].Email === S1 && readTab_('Predictions')[0].Checkpoint === 'intro');
check('re-save overwrites the same row', (() => { const again = JSON.parse(JSON.stringify(full)); again.cv.rating = 'work-on'; as(S1, 'savePrediction', { checkpoint: 'intro', answers: again }); return readTab_('Predictions').length === 1 && as(S1, 'bootstrap').data.student.predictions.intro.items.cv.rating === 'work-on'; })());
check('another student sees no prediction', JSON.stringify(as(S2, 'bootstrap').data.student.predictions) === '{}');
check('teacher / wrong domain / no token cannot save', /class list/.test(as(T, 'savePrediction', { checkpoint: 'intro', answers: full }).error) && /class list/.test(as('x@gmail.example', 'savePrediction', { checkpoint: 'intro', answers: full }).error) && authErr(post('savePrediction', { checkpoint: 'intro', answers: full }, '')));
check('client-supplied email in the body is ignored', (() => { as(S2, 'savePrediction', { checkpoint: 'intro', answers: full, email: S1, Email: S1 }); const rows = readTab_('Predictions'); return rows.length === 2 && rows.filter(r => r.Email === S2).length === 1; })());
check('bootstrap payload never carries another student\'s rows', !JSON.stringify(as(S2, 'bootstrap')).includes('sample.one'));

console.log('retention');
FakeSheets.uiAnswer = 'YES';
clearProfileData();
check('clearProfileData empties Profile + Predictions, keeps Students', readTab_('Profile').length === 0 && readTab_('Predictions').length === 0 && readTab_('Students').length === 4);

console.log(failed ? `\n${failed} check(s) FAILED` : '\nall checks passed');
process.exit(failed ? 1 : 0);
