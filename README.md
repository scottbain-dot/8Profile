# G8 PE Profile — FIS Grade 8

A student-facing PE profile for Grade 8 at Frankfurt International School. A student
opens one link, is signed in with their FIS Google account, is greeted by name, picks
their **PE Style** (their athlete-card identity, saved), and sees what the year holds.
Later phases fill the profile and card with their real Active Participation,
Combine and fundamental-skill data (see `PLAN.md`).

**This repo is code only.** No rosters, no names, no scores, no photos, no Sheet IDs.
Real data lives only in an FIS-owned Google Sheet and is read at runtime by the
signed-in student's own Google session. Read `PRIVACY.md` before changing anything
that touches identity or data.

## How it works

```
student's browser ──(FIS Google sign-in, by Google)──▶ Apps Script web app (inside the FIS Sheet)
                                                        │  Session.getActiveUser() = verified fis.edu email
                                                        │  identity_()  → the caller's own Students row
                                                        │  bootstrap()  → name, class, style (own row only)
                                                        │  saveStyle()  → Profile row keyed by that email
                                                        ▼
                                              FIS-owned Google Sheet
                                              Config · Students · Teachers · Profile
```

- **Google Apps Script web app, container-bound** to the Sheet. Deployed *Execute as Me*
  (the Sheet owner) with access *Anyone within fis.edu*. Google does the sign-in; the
  script asks Google who the caller is. The browser never sends an email and never
  receives anyone else's row. Students need no access to the Sheet itself.
- **Frontend** served by `HtmlService` from the same script: `src/Index.html` (shell),
  `src/Styles.html` (CSS, the mockups' visual system), `src/App.html` (client JS).
- **No third-party requests** by default: no CDN scripts, no analytics, no cookies beyond
  Google's own session. Web fonts are off unless the Config flag `web_fonts` is TRUE.
- **Photo**: the student's Google directory photo, looked up at runtime only when Config
  `photo_lookup` is TRUE and the People API advanced service is enabled. Never copied or
  stored. Initials otherwise.

This is the same pattern as the `net-games` PE tracker, which FIS students already use.

## Layout

```
src/
  Code.gs          server: tabs, config, identity, bootstrap, saveStyle/saveGoal, Sheet menu
  Index.html       page shell (Apps Script template)
  Styles.html      CSS
  App.html         client: greeting, style picker + live card, teaser profile, "what's coming"
  appsscript.json  manifest: least-privilege scopes, web app = DOMAIN access, executes as deployer
dist/
  Code.gs          GENERATED single file (server + embedded HTML) — what gets pasted. node dev/build-single.js
dev/
  fake-sheets.js   in-memory stand-in for SpreadsheetApp & co, so Code.gs runs in Node / a browser
  mock-runtime.js  fake google.script.run + synthetic roster seed (browser preview)
  build-preview.js builds dev/preview.html from the real app files
  build-single.js  builds dist/Code.gs
  test-server.js   Node checks of the privacy guarantees (own row only, domain refused, writes validated)
  smoke.js         headless Chromium run of the preview (Playwright); screenshots in dev/shots/ (ignored)
data/dummy/        synthetic test data only
design/            the mockups the app is built from (fictional sample data)
docs/              G8_PE_Rubric_Bank.md (data source of truth for v2+), BUILD_BRIEF.md
PRIVACY.md         data flows, scopes, storage, retention — for the FIS data-protection lead
PLAN.md            architecture decisions, milestones v1–v4, open questions
```

## Developing

```
npm test                                      # node dev/test-server.js — server checks against the fake Sheet
npm run preview                               # dev/preview.html — open ?role=student | student2 | teacher | unknown | wrong | anon
                                              #   &photo=1  &fail=1  &latency=800  &reset=1  &as=someone@example.edu
NODE_PATH=$(npm root -g) npm run smoke        # needs playwright + Chromium
npm run build                                 # regenerate dist/Code.gs — commit it
```

The preview runs the real `Code.gs` in the browser against a fake spreadsheet kept in
`localStorage`, so identity and write logic are exercised too. All preview data is
synthetic (`example.edu`).

## Deploying (teacher, once)

1. Create a Google Sheet **in the FIS Workspace**, owned by the PE teacher account.
2. Extensions → Apps Script. Replace the contents of `Code.gs` with `dist/Code.gs`.
   Project Settings → tick *Show "appsscript.json"* and replace it with `src/appsscript.json`.
3. Save. In the editor's function dropdown pick **`setupTabs`** and click **Run**. Authorise
   with your FIS account when Google asks (Advanced → Go to project if it warns the app is
   unverified). This creates the four tabs. Then **reload the Sheet**: the **PE Profile**
   menu appears. (The menu is built on open and cannot ask for authorisation itself, so it
   stays hidden until this one manual run.)
4. Fill **Students** (Email · Name · Class) with the Grade 8 roster. Add colleagues to
   **Teachers**. Check **Config** (`domain` = `fis.edu`, `year_label`).
5. Deploy → New deployment → Web app: *Execute as* **Me**, *Who has access* **Anyone within
   Frankfurt International School**. Copy the `/exec` URL. That is the link students open.
6. Test with a **dummy student account** in the FIS domain (on the roster), a domain
   account not on the roster, and a non-FIS account. Expected: own profile / "not on the
   list" / "use your FIS account". Only then share the link.
7. Optional photo: Apps Script editor → Services → add **People API**, set Config
   `photo_lookup` to TRUE, redeploy. If the directory does not return photos, students
   simply see initials.

After code changes: `npm run check`, commit, paste the new `dist/Code.gs`, then
Deploy → Manage deployments → edit → *New version*. The URL stays the same.

### Troubleshooting

- **No "PE Profile" menu.** Run `setupTabs` once from the editor (step 3) and reload the Sheet.
- **"Cannot call SpreadsheetApp.getActiveSpreadsheet"** or tabs appear nowhere: the script is
  not bound to the Sheet. Open the editor from the Sheet via Extensions → Apps Script and paste there.
- **Student sees "Couldn't confirm who you are".** The deployment is not "Anyone within FIS",
  or the student is signed into a personal Google account in that browser profile.
- **Changed the Sheet's Config and nothing happens.** Config is cached for two minutes; use
  PE Profile → Clear config cache.

Keep the Sheet ID, script ID and `/exec` URL out of this repo (`config.example.json`
shows the shape; `config.json` is git-ignored).
