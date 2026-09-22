# Build brief — G8 PE Student Profile web app

## Your role
You are building a **student-facing PE profile web app** for Grade 8 PE at Frankfurt International School (FIS). This is a real, privacy-sensitive school system for minors. Read the **Privacy & GDPR constraints** section first and treat it as non-negotiable — it shapes every architecture decision.

## Step 0 — inspect what already exists (do this before proposing anything)
1. **Check my existing FIS projects for the established Google sign-in / auth pattern** (e.g. the Athlete Academy portals / aa-dash, the FIS dashboard, the net-games tracker). Identify how I already do FIS Google sign-in, domain restriction, and Sheets access, and **reuse that exact pattern** unless there's a strong reason not to. Do not invent a new auth approach if a working one already exists in my repos.
2. **Review the design artefacts I'll provide** (built as mockups, to be productionised):
   - `My_PE_Profile_mockup.html` — the full student profile/dashboard (sections, colours, layout system)
   - `PE_Athlete_Card.html` — the collectible athlete card (foil frame, radar, PE-Style skins, Bronze/Silver/Gold tiers, download)
   - `G8_Build_Your_Card.html` — the "build your card" flow (name + style picker + live reskin)
   - `G8_PE_Rubric_Bank.md` — **the data source of truth**: the 5 AP outcomes + band descriptors, the 9 fundamental-skill ladders (4 levels each), the 5 fitness tests, the K/S/CT report bands, and how everything maps.
   Reuse their visual system (fonts Archivo + Inter, FIS maroon `#6f1d2c` + gold `#e7b64b`, the radar, the card, the level bars) rather than redesigning.

## The product
A page a G8 student opens and:
1. **Signs in with their FIS Google account** (automatic, domain-restricted to fis.edu).
2. Is **greeted by name**, with their **Google account photo** in the card avatar (never uploaded, never stored).
3. **Picks their PE Style** (the 6 styles = their athlete-card identity), sees the card reskin live, and their choice is **saved**.
4. Sees a **teaser profile + "what's coming this year"** (Combine, skill ladders, participation, how the card levels up).
5. **Later phases:** their real AP self/teacher data, Combine results, fundamental-skill levels, report bands — populating the profile and card for real.

## Privacy & GDPR constraints (NON-NEGOTIABLE)
This handles data about minors under GDPR. Build to these rules:
- **Everything stays inside the FIS Google Workspace / organisation.** Auth, hosting of data, and data storage must all be within FIS-owned Google (Workspace, Apps Script, Sheets, Drive). No student data on any third-party host, database, or SaaS.
- **No student personal data ever passes through an AI tool** (including you). Build and test with **dummy/synthetic data only**. Real rosters, names, scores, and photos live only in FIS-owned Sheets and are read **at runtime** by the authenticated user's own Google session — never pasted into prompts, commits, logs, or issues.
- **Domain-restricted auth**: only `@fis.edu` accounts (`hd=fis.edu`, verify server/gs-side, not just client). A student can only ever see **their own** record; no client-side access to the full roster.
- **Data minimisation**: store only what the profile needs (AP bands, skill levels, fitness numbers, PE-Style choice, goals). No free-text beyond the student's own goal. The profile only *displays* data already collected in PE; it creates no new sensitive categories.
- **Photos = the Google account avatar at runtime.** Never download, copy, or store a student image. Initials fallback if no photo.
- **No trackers, no third-party analytics, no external fonts/scripts that carry data.** (Google Fonts is acceptable; if strict, self-host the two fonts.) No cookies beyond the Google auth session.
- **Least privilege**: request the narrowest Google API scopes that work. Prefer the signed-in user reading only their own row.
- **Secrets & data never in the repo**: Sheet IDs and any config go in a non-committed config (or Script Properties); `.gitignore` excludes any data files. The repo is code only, safe to be public if it ever is — but **contains no data and no secrets**.
- **Retention & consent**: flag (don't assume) that retention periods and parental-consent handling must be confirmed with the FIS data-protection lead before go-live. Add a short `PRIVACY.md` documenting data flows, scopes, storage location, and retention, so it can be reviewed.

## Recommended architecture (confirm against my existing projects first)
Given "everything inside the Google org," the safest primary option is a **Google Apps Script web app**:
- Hosted by Google under FIS Workspace; deploy "execute as user accessing the app," access restricted to the FIS domain.
- `Session.getActiveUser().getEmail()` gives the verified FIS identity server-side; use it to fetch only that student's row from a FIS-owned **Google Sheet** (the data store, where the AP form responses already land).
- Frontend served via `HtmlService` (can still be a rich SPA-style page reusing the mockup's HTML/CSS/JS).
- Writes (PE-Style choice, goal) go back to the same Sheet, keyed by the verified email.
**Alternative** (if it matches my existing pattern better): a static SPA + **Google Identity Services** for sign-in + **Sheets API** for data — acceptable **only if** all data access is authenticated, domain-restricted, per-user, and no data ever touches a third party. Note: static hosting (e.g. GitHub Pages) puts *code* (no data) outside the org; if the requirement is strictly "all inside Google," prefer the Apps Script web app.
**First check whether student accounts are allowed to run Apps Script web apps / expose `getActiveUser()` in this Workspace-for-Education tenant** — some tenants restrict this for students. This is the make-or-break; verify before committing to the approach.

## Data model (from the Rubric Bank — see `G8_PE_Rubric_Bank.md`)
- **Active Participation**: 5 outcomes (Active Participation/C1, Great teammate/S4, Be ready, Growth mindset, Manage myself & feelings). Values on the FIS band scale: no evidence · 1–2 · 3–4 · 5–6 · 7. Captured self (start + end) and teacher (checkpoints).
- **PE Styles** (card identity, 6): Competitor, Team Player, Improver, Explorer, Energizer, Active Thinker.
- **Fundamental skills** (9, each a 4-level ladder): throw, catch, strike, dribble, kick, balance, agility, jump-land, core stability.
- **Fitness (Combine, 5)**: mile (CV endurance), push-ups + sit-to-stand (muscular endurance), broad jump (power), 40m sprint (speed), measured Sept/Jan/May.
- **Report bands**: Knowledge · Skills · Concept Transfer (best fit across descriptors, 1–7).
- **Card tiers**: Bronze/Silver/Gold, earned by growth/breadth/mastery/consistency (never by raw ability).

## Repo setup
- New private repo, e.g. `fis-g8-pe-profile`. `README.md` (setup, deploy, architecture), `PRIVACY.md` (data flows + scopes + retention, for the DP lead), `.gitignore` (config, data, secrets).
- `/src` app code; `/data/dummy` synthetic test data only; a `config.example` (no real IDs). Never commit real Sheet IDs, rosters, or `.env`.
- Clear commit hygiene: no data, no PII, no screenshots containing real students.

## Milestones
- **v1 (this is the target):** FIS sign-in (domain-restricted) → greet by name + Google photo → PE-Style picker (saved to Sheet) → teaser profile + "what's coming." No live AP/Combine data yet. Deployable, tested with dummy accounts.
- **v2:** live AP data (self start/end + teacher) rendered in the profile + card radar.
- **v3:** Combine fitness + fundamental-skill levels; report-band summary.
- **v4:** tiers, badges, progress-over-time.

## Definition of done for v1
- A single FIS-domain link a student opens; signs in with their FIS account; sees their name + photo; picks a style that persists; sees the teaser + what's coming.
- Verified: a student cannot see another student's data; a non-fis.edu account is refused; no data leaves FIS Google; repo contains no data or secrets; `PRIVACY.md` written for review.

## First reply I want from you
1. What you found in my existing projects for the sign-in pattern, and whether to reuse it.
2. Your architecture recommendation (Apps Script web app vs static+GIS) with the reason, given the privacy constraints and the student-Apps-Script-access question.
3. The repo structure you'll create.
4. Any privacy/GDPR point you'd escalate to the FIS data-protection lead before go-live.
Then scaffold v1.
