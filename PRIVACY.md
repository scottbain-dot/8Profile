# Privacy & data protection — G8 PE Profile

For review by the FIS data-protection lead before go-live. This app displays, to each
Grade 8 student, PE information about that student only. It handles personal data of
minors under GDPR, so it is built to keep everything inside the FIS Google Workspace.

## Summary

| | |
|---|---|
| Data subjects | Grade 8 PE students; PE teachers (as users) |
| Controller | Frankfurt International School |
| Processor | Google (Workspace for Education, under the existing FIS agreement). No other processor. |
| Where data lives | One Google Sheet owned by the PE teacher's FIS account, inside the FIS Workspace |
| Where code runs | Google Apps Script, bound to that Sheet, in the FIS Workspace |
| Who can open the app | Only signed-in `@fis.edu` accounts (deployment setting *and* a server-side check) |
| What a student sees | Their own row only. The class list never reaches a browser. |
| Third parties | None by default. No CDNs, analytics, trackers or cookies beyond Google's session. |
| This repository | Code only. No data, no IDs, no secrets. Test data is synthetic (`example.edu`). |

## Data inventory (v1)

| Tab | Fields | Source | Purpose |
|---|---|---|---|
| Students | Email, Name, Class | Existing PE roster, entered by the teacher | Recognise the signed-in student and greet them |
| Teachers | Email, Name | Teacher | Show the teacher screen instead of a student profile |
| Profile | Email, Style (one of six labels), Goal (≤140 chars, optional, own words), Updated | Written by the student through the app | Persist the student's chosen PE Style / goal |
| Config | key/value settings | Teacher | No personal data |

Later phases add data already collected in PE (Active Participation bands, Combine
fitness numbers, fundamental-skill levels, report bands). The app creates no new
categories of data and no free text beyond the student's own goal. There are no special
categories (health data is not collected; fitness test results are PE assessment data
already held by the school).

## Data flows

1. **Sign-in.** Google authenticates the student on the FIS domain. The app never sees a
   password or token. The script reads the verified address with
   `Session.getActiveUser().getEmail()` on the server. `identity_()` in `src/Code.gs` is
   the only place this happens.
2. **Domain check.** The deployment is restricted to the FIS Workspace, and the server
   additionally refuses any address outside the configured domain.
3. **Read.** `bootstrap()` matches the address against the Students tab and returns that
   one row plus that student's Profile row. No function returns the roster or any other
   student's data. The response does not even contain the student's own email.
4. **Write.** `saveStyle()` / `saveGoal()` validate the value and upsert one Profile row
   keyed by the *server-verified* address. The client cannot choose whose row is written.
5. **Photo (optional, off by default).** With Config `photo_lookup` = TRUE and the People
   API service enabled, the server looks up the student's own Google directory photo and
   returns its URL. The browser loads the image straight from Google. Nothing is copied,
   stored or logged. Initials are shown when there is no photo or the lookup is off.
6. **Logs.** Apps Script keeps execution logs in Google Cloud (Stackdriver) for the FIS
   project. The code never logs emails or data, and error messages thrown to users are
   written so they contain no email.

## Google API scopes (least privilege)

Declared in `src/appsscript.json`:

| Scope | Why |
|---|---|
| `spreadsheets.currentonly` | Read/write the one Sheet the script is bound to. No access to other Drive files. |
| `userinfo.email` | Learn the signed-in user's address (the identity check). |
| *(optional)* People API directory read | Only if the photo feature is switched on; added by enabling the service. |

The script runs as the teacher account, so students are never asked to grant scopes and
need no access to the Sheet.

## Third-party requests

- Default: none. Fonts use the system stack.
- Config `web_fonts` = TRUE loads Archivo and Inter from `fonts.googleapis.com`, which
  sends the viewer's IP address to Google outside the Workspace agreement. German courts
  have treated remote Google Fonts as a GDPR issue (LG München, 2022). Recommendation:
  leave it off, or embed the fonts in the page in a later change.
- The PNG "download my card" feature from the mockups uses the `html2canvas` library
  from a CDN. It is **not** in the app. It will be added only with the library vendored
  into the repo, so no third-party request is made.

## Retention & deletion

- Profile rows (style, goal) are the only app-written data. The Sheet menu
  **PE Profile → End of year: clear Profile data…** deletes them all. Proposed retention:
  delete at the end of each school year. **To be confirmed by the DP lead.**
- Students and Teachers tabs are the school's existing roster data; retention follows the
  school's roster policy.
- A single student's data can be removed by deleting their row(s) in the Sheet.
- No copies exist outside the Sheet: no browser storage beyond the current page, no
  exports, no emails.

## Points to confirm with the data-protection lead before go-live

1. **Lawful basis and parental information.** The app displays PE assessment data the
   school already holds. Confirm whether the existing school/parent information covers
   it, or whether a notice (or consent for the optional photo) is needed.
2. **Retention period** for Profile rows and, in later phases, for assessment data shown
   in the app (proposal above).
3. **Google directory photo.** Showing the account photo inside a school app is
   arguably within existing Workspace use, but confirm before enabling `photo_lookup`.
4. **Web fonts.** Keep off unless approved (see above).
5. **Teacher accounts.** Teachers see only a teacher screen in v1. Later phases may add a
   teacher "view a student's profile" function; that is a new access path and should be
   reviewed when designed.
6. **Later phases** (AP, Combine, skill levels, report bands) widen the data displayed.
   Each phase updates the inventory above and is re-reviewed.

## The GitHub Pages address

`index.html` at the repo root is published on GitHub Pages as a short, memorable front
door. It contains no data and no tracking: it forwards the browser to the Apps Script
`/exec` URL and nothing else. That URL is therefore public, which is acceptable because
access to the app is enforced by Google (FIS accounts only), not by keeping the address
secret. The page and the profile app never exchange data.

## Repository hygiene

- `.gitignore` excludes config, data, exports and screenshots. Only `data/dummy/` (synthetic)
  is allowed in.
- Commits, issues and screenshots must never contain real students. Test with dummy
  FIS accounts and synthetic data only.
- No AI tool is ever given real student data; development uses synthetic data only.
