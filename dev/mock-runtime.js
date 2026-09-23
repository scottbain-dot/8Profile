// Browser-only: wires the real src/Code.gs (running against fake-sheets.js)
// up to a fake API (window.fetch) and a fake Sign in with Google, seeds a
// SYNTHETIC roster, and picks who "you" are from the URL:
//   ?role=student | student2 | teacher | owner | unknown | wrong | anon | expired   (default student)
//   ?as=someone@example.edu   sign in as an exact address
//   &fail=1                   every write fails (to test error handling)
//   &latency=800              simulated round-trip in ms
//   &reset=1                  wipe the fake spreadsheet stored in localStorage
//   &nophoto=1                token without a picture (initials fallback)
// Nothing here is real: names and addresses are made up, domain example.edu.
(function () {
  'use strict';
  const q = new URLSearchParams(location.search);
  const STORE = 'g8pe_fake_book_v2';
  const latency = parseInt(q.get('latency') || '200', 10);
  const failWrites = q.get('fail') === '1';
  window.__preview = { fastRetry: true };
  window.PE_CONFIG = { apiUrl: 'https://api.example/exec', clientId: 'preview-client', domain: 'example.edu' };

  if (q.get('reset') !== '1') {
    try { const j = localStorage.getItem(STORE); if (j) FakeSheets.book = FakeSheets.Book.fromJSON(JSON.parse(j)); } catch (e) {}
  }
  FakeSheets.book.onChange = b => { try { localStorage.setItem(STORE, JSON.stringify(b.toJSON())); } catch (e) {} };

  const ROSTER = [
    ['sample.one@example.edu',   'Sample One',    '8A'],
    ['sample.two@example.edu',   'Two, Sample',   '8A'], // "Last, First" as the school system exports it
    ['sample.three@example.edu', 'Sample Three',  '8B'],
    ['sample.four@example.edu',  'Sample Four',   '8B'],
    ['teacher@example.edu',      'Teacher Test',  '8B'] // the owner, also on the roster to test the student view
  ];
  const PHOTO = 'https://lh3.googleusercontent.com/preview/' + encodeURIComponent('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" fill="#cfe4f2"/><circle cx="50" cy="40" r="18" fill="#e8b98f"/><path d="M18 100C18 74 34 66 50 66s32 8 32 34z" fill="#3f6d8c"/></svg>');
  const tokenFor = (email, extra) => FakeSheets.token(Object.assign({
    aud: 'preview-client', email, email_verified: 'true', hd: email.split('@')[1], name: 'Token Name',
    picture: q.get('nophoto') === '1' ? '' : PHOTO, exp: Math.floor(Date.now() / 1000) + 3600
  }, extra || {}));
  function setConfig(key, val) { const cfg = FakeSheets.book.getSheetByName('Config'); cfg.getDataRange().getValues().forEach((r, i) => { if (r[0] === key) cfg.getRange(i + 1, 2).setValues([[val]]); }); }
  function callApi(body) { return JSON.parse(doPost({ postData: { contents: JSON.stringify(body) } }).getContent()); }

  function seed() {
    FakeSheets.owner = 'teacher@example.edu';
    setupTabs();
    setConfig('domain', 'example.edu'); setConfig('oauth_client_id', 'preview-client');
    const st = FakeSheets.book.getSheetByName('Students');
    st.getRange(2, 1, ROSTER.length, 3).setValues(ROSTER);
    const t = FakeSheets.book.getSheetByName('Teachers');
    t.getRange(t.getLastRow() + 1, 1, 1, 2).setValues([['other.teacher@example.edu', 'Other Teacher']]);
    clearConfigCache();
    const tok = tokenFor('sample.two@example.edu');
    callApi({ action: 'saveStyle', token: tok, style: 'improver' }); // one student has already chosen
    const ans = {}; PREDICT_ITEMS.forEach((k, i) => { ans[k] = { rating: RATINGS[i % 3], freq: FREQS[(i + 1) % 3] }; });
    callApi({ action: 'savePrediction', token: tok, checkpoint: 'intro', answers: ans }); // ...and made their Lesson 1 prediction
  }
  if (!FakeSheets.book.getSheetByName('Students')) seed();
  FakeSheets.owner = 'teacher@example.edu';
  FakeSheets.cache = {};

  // who am I → a fake token in sessionStorage (what a real sign-in leaves behind)
  const role = q.get('role') || 'student';
  let email = ROSTER[0][0], extra = null;
  if (q.get('as')) email = q.get('as');
  else if (role === 'student2') email = ROSTER[1][0];
  else if (role === 'teacher') email = 'other.teacher@example.edu';
  else if (role === 'owner') email = 'teacher@example.edu';
  else if (role === 'unknown') email = 'nobody@example.edu';
  else if (role === 'wrong') email = 'someone@gmail.example';
  else if (role === 'expired') extra = { exp: Math.floor(Date.now() / 1000) - 10 };
  try { sessionStorage.removeItem('pe_token'); } catch (e) {}
  if (role !== 'anon') {
    const tok = tokenFor(email, extra);
    // an expired token is dropped by the page before use; keep it "valid-looking" for the page but rejected by the server
    if (role === 'expired') { const t2 = tokenFor(email, { exp: Math.floor(Date.now() / 1000) + 3600, aud: 'someone-else' }); sessionStorage.setItem('pe_token', t2); }
    else sessionStorage.setItem('pe_token', tok);
  }

  // fake Sign in with Google: a button that hands over a token for the chosen role
  window.google = { accounts: { id: {
    initialize() {}, prompt() {}, disableAutoSelect() {},
    renderButton(el) { el.innerHTML = '<button type="button" id="fakegsi" class="ct-btn">Sign in with Google (preview)</button>'; el.querySelector('button').onclick = () => window.__gsiCallback && window.__gsiCallback({ credential: tokenFor(role === 'anon' ? ROSTER[0][0] : email) }); }
  } } };
  const realInit = window.google.accounts.id.initialize;
  window.google.accounts.id.initialize = o => { window.__gsiCallback = o.callback; realInit(o); };

  // fake API transport: window.fetch → doPost in-page
  const realFetch = window.fetch;
  window.fetch = (url, opts) => {
    if (url !== window.PE_CONFIG.apiUrl) return realFetch(url, opts);
    return new Promise(resolve => setTimeout(() => {
      let out;
      try {
        const body = JSON.parse(opts.body);
        if (failWrites && /^save/.test(body.action)) out = { ok: false, error: 'Preview: writes are set to fail (&fail=1)' };
        else out = callApi(body);
      } catch (e) { out = { ok: false, error: e.message }; }
      resolve({ ok: true, json: async () => out });
    }, latency));
  };
})();
