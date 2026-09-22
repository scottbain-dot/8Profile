# Plan — G8 PE Profile

Status: **v1 deployed to the FIS Sheet and working for the owner account; profile look + photo added.** Last updated 2026-09-22.

## Decisions taken (and why)

| Decision | Choice | Reason |
|---|---|---|
| Sign-in pattern | **Reuse `net-games`**: Apps Script web app inside the Sheet, `Session.getActiveUser()` server-side, deployed *Execute as Me* + *Anyone within fis.edu* | Proven with FIS students already; every privacy rule in the brief falls out of it. The `aa-dash` pattern (static GitHub Pages + Google Identity Services) sends a client-supplied email to Apps Script with no server verification or domain check, so anyone who knows an address could read that record. Not acceptable here. |
| Execute as *Me*, not *user accessing* | Me (the Sheet owner) | The brief suggested "execute as user". That would require every student to have read access to the Sheet (exposing the whole roster) and to grant scopes. Executing as owner keeps students out of the Sheet; the server filters to their own row. |
| Static hosting | None | Code and data both stay inside Google. The repo is the source; `dist/Code.gs` is pasted into the Sheet's script. |
| Fonts | System stack by default; Google Fonts behind Config `web_fonts` (off) | Remote Google Fonts is a known GDPR issue in Germany. The DP lead can flip it or we embed the fonts later. |
| Display font | Archivo (brief) with Barlow Condensed → system fallback | The Build-Your-Card mockup uses Archivo; the profile/card mockups use Barlow Condensed. Brief says Archivo. |
| Photo | Google directory photo via People API, runtime only, behind Config `photo_lookup` (off) | "Execute as Me" gives no direct handle on the *caller's* photo; the directory lookup is the least-privilege route. Falls back to initials. Needs a tenant check (below). |
| Card download | Deferred | Needs `html2canvas`; ship only when vendored (no CDN). Not in the v1 definition of done. |
| Profile key | Server-verified email | Simplest correct key; the client never supplies it. If the DP lead prefers pseudonymous IDs, add a `StudentID` column to Students and key Profile on it (one-line change in `PROFILE_KEY`). |
| Teacher view (v1) | A plain teacher screen | Any "view as student" function is a new access path; design it with the DP review in v2. |

## Make-or-break checks on first deploy (cannot be verified from here)

1. **Students can open a domain-restricted web app in the FIS tenant.** `net-games` is
   precedent that they can. Confirm with one dummy student account on this deployment.
2. **`Session.getActiveUser().getEmail()` returns the student's address** when the app
   executes as the owner. It does for same-domain Workspace users (again, `net-games`).
3. **People API directory search returns photos** for a non-admin owner account with
   `DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE`. Depends on FIS directory-sharing settings.
   If not, students see initials and nothing breaks.
4. A **non-FIS Google account is refused** at the Google layer (deployment access) and,
   belt-and-braces, at `identity_()`.

## Milestones

### v1 — sign-in → greet → pick style (saved) → teaser  ✅ scaffolded
- [x] Server: tabs, config, identity, bootstrap (own row only), saveStyle/saveGoal, menu, retention wipe
- [x] Client: profile page in the mockup's look (teaser card strip, goals, at-a-glance, fitness, nine skill ladders from the Rubric Bank, game skills, participation, report bands, tap-to-open overlays) with every section in its "fills in at the Combine" state; "Build my card" screen with the six-style picker and live reskin; first visit lands on the builder, later visits on the profile
- [x] Roster names in "Last, First" form: greet by first name, ignore the initial
- [x] Google photo: People API, then Admin Directory fallback; `src/appsscript.photo.json` turns it on
- [x] Non-student screens: teacher, not-on-list, wrong domain, anonymous, load error
- [x] Dev: fake Sheets runtime, browser preview with synthetic roster, single-file build, 29 server checks, 14 browser smoke checks
- [x] Docs: README (deploy), PRIVACY (for DP lead), this plan, CLAUDE.md
- [x] Deployed to the FIS Sheet; owner sees own profile (check 1 & 2 pass for a teacher account)
- [ ] Confirm with one **student** account, and enable the photo via `appsscript.photo.json` (check 3)
- [ ] Share `PRIVACY.md` with the DP lead; record the answers to its six questions
- [ ] Go-live: fill Students, share the `/exec` link

### v1 + Lesson 1 — "Strengths & challenges" self-prediction  ✅ built
- [x] `Predictions` tab (Email · Checkpoint · Timestamp · 13 × rating/freq), `savePrediction()` validates all 13 and upserts by verified email + checkpoint
- [x] Predict view (traffic light + frequency per item, progress bar, save only when complete), summary view (strengths / work-ons), profile panel with chips + Edit, compare view stub reading prediction next to `combine` (null until v3)
- [x] Re-opening is editable; saving again overwrites the intro row with a new timestamp
- [x] Wording aligned with the Lesson 1 deck (`lesson1/index.html`, hosted on Pages)

### v1.1 — polish once real students have used it
- Vendor `html2canvas` → "Save my card" PNG (runs inside the school login; the file is the student's own)
- Optional: embed Archivo/Inter as base64 WOFF2 in `Styles.html` so fonts need no third-party request
- Student goal UI (server `saveGoal` already exists; Profile has the column)

### v2 — live Active Participation data
- New tabs fed by the existing AP form responses: self (start/end) and teacher checkpoints for the 5 keys
  (Be ready · Best effort · Growth mindset · Great teammate · Manage myself & feelings)
- Generate the "where I am / next step" wording from `docs/G8_PE_Rubric_Bank.md`
  (write a `dev/build-rubrics.js` that emits `src/Rubrics.html` as a JS object, so the Rubric Bank stays the single source)
- Profile: participation section + trend; card: Effort and Team axes become real
- Data-minimisation rule stays: bootstrap returns the caller's rows only

### v3 — Combine + fundamental skills + report bands
- Combine tab: mile, push-ups, sit-to-stand, broad jump, 40 m sprint × Sept/Jan/May
- Skills tab: 9 ladders × level (1–4) × checkpoint, with self → peer → teacher sign-off state
- Report bands: Knowledge · Skills · Concept Transfer (best fit, 1–7) from the Rubric Bank map
- Fill `combineFor_()` so the compare view shows prediction vs. measurement per item
- Card radar from real data; "faint September self" ghost line

### v4 — tiers, badges, progress over time
- Bronze/Silver/Gold from growth · breadth · mastery · consistency (never raw ability)
- Badges (Mile PB, Big Riser, Full Effort, All-Rounder …), per-key sparklines, reflection

## Open questions for Scott
- Confirm the six PE Style names and one-line descriptions as shipped in `src/App.html`.
- Card tier label in v1 shows **BRONZE** for everyone (earned tiers arrive in v4). OK, or hide it until then?
- Class label format on the roster (`8A` vs `8(06B)` as in the mockups): the app shows whatever is in the Class column.
- Who is on the Teachers tab in v1 (only affects who sees the teacher screen).
