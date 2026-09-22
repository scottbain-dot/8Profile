// Browser-only: wires the real src/Code.gs (running against fake-sheets.js)
// up to a fake google.script.run, seeds a SYNTHETIC roster, and picks who
// "you" are from the URL:
//   ?role=student | student2 | teacher | unknown | wrong | anon   (default student)
//   ?as=someone@example.edu   sign in as an exact address
//   &fail=1                   every write fails (to test error handling)
//   &latency=800              simulated round-trip in ms
//   &reset=1                  wipe the fake spreadsheet stored in localStorage
//   &photo=1                  pretend the directory returned a photo
// Nothing here is real: names and addresses are made up, domain example.edu.
(function () {
  'use strict';
  const q = new URLSearchParams(location.search);
  const STORE = 'g8pe_fake_book_v1';
  const latency = parseInt(q.get('latency') || '200', 10);
  const failWrites = q.get('fail') === '1';
  window.__preview = { fastRetry: true };

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
  function seed() {
    FakeSheets.owner = 'teacher@example.edu';
    FakeSheets.user = FakeSheets.owner;
    setupTabs();
    const cfg = FakeSheets.book.getSheetByName('Config');
    const rows = cfg.getDataRange().getValues();
    rows.forEach((r, i) => { if (r[0] === 'domain') cfg.getRange(i + 1, 2).setValues([['example.edu']]); });
    const st = FakeSheets.book.getSheetByName('Students');
    st.getRange(2, 1, ROSTER.length, 3).setValues(ROSTER);
    const t = FakeSheets.book.getSheetByName('Teachers');
    t.getRange(t.getLastRow() + 1, 1, 1, 2).setValues([['other.teacher@example.edu', 'Other Teacher']]);
    clearConfigCache();
    FakeSheets.user = 'sample.two@example.edu';
    saveStyle('improver'); // one student has already chosen
  }
  if (!FakeSheets.book.getSheetByName('Students')) seed();
  FakeSheets.owner = 'teacher@example.edu';
  FakeSheets.cache = {};

  const role = q.get('role') || 'student';
  if (q.get('as')) FakeSheets.user = q.get('as');
  else if (role === 'student') FakeSheets.user = ROSTER[0][0];
  else if (role === 'student2') FakeSheets.user = ROSTER[1][0];
  else if (role === 'teacher') FakeSheets.user = 'other.teacher@example.edu';
  else if (role === 'owner') FakeSheets.user = FakeSheets.owner;
  else if (role === 'unknown') FakeSheets.user = 'nobody@example.edu';
  else if (role === 'wrong') FakeSheets.user = 'someone@gmail.example';
  else if (role === 'anon') FakeSheets.user = '';

  if (q.get('photo') === '1') {
    // A fake People API that "finds" a photo: a generated SVG, no real image anywhere.
    const svg = encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" fill="#cfe4f2"/><circle cx="50" cy="40" r="18" fill="#e8b98f"/><path d="M18 100C18 74 34 66 50 66s32 8 32 34z" fill="#3f6d8c"/></svg>');
    window.People = { People: { searchDirectoryPeople: () => ({ people: [{ photos: [{ url: 'data:image/svg+xml,' + svg }] }] }) } };
    const cfg = FakeSheets.book.getSheetByName('Config');
    cfg.getDataRange().getValues().forEach((r, i) => { if (r[0] === 'photo_lookup') cfg.getRange(i + 1, 2).setValues([['TRUE']]); });
  }

  const SERVER_FNS = ['bootstrap', 'saveStyle', 'saveGoal'];
  const isWrite = fn => /^save/.test(fn);
  function makeRunner() {
    const r = { _ok: null, _err: null };
    r.withSuccessHandler = fn => { r._ok = fn; return r; };
    r.withFailureHandler = fn => { r._err = fn; return r; };
    SERVER_FNS.forEach(name => {
      r[name] = (...args) => {
        const cloned = JSON.parse(JSON.stringify(args));
        setTimeout(() => {
          try {
            if (failWrites && isWrite(name)) throw new Error('Preview: writes are set to fail (&fail=1)');
            const out = window[name](...cloned);
            if (r._ok) r._ok(out === undefined ? null : JSON.parse(JSON.stringify(out)));
          } catch (e) { if (r._err) r._err({ message: e.message }); }
        }, latency);
      };
    });
    return r;
  }
  window.google = { script: { get run() { return makeRunner(); } } };
})();
