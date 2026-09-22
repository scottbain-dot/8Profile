# CLAUDE.md — G8 PE Profile

Student-facing PE profile for Grade 8 at FIS. Google Apps Script web app bound to an
FIS-owned Sheet. Read `PRIVACY.md` and `PLAN.md` first.

## Privacy rules (non-negotiable)

- **No real student data, ever, in this repo or in any prompt.** Names, emails, scores,
  photos, rosters, screenshots of real students: none. Test with `data/dummy/` and the
  preview's synthetic `example.edu` roster only.
- **No Sheet IDs, script IDs, deployment URLs or tokens in the repo.** `config.json` is
  git-ignored; `config.example.json` shows the shape with placeholders.
- **Identity is read in one place**: `identity_()` in `src/Code.gs`, from
  `Session.getActiveUser()`. Never accept an email (or any identifier) from the client.
- **Every read/write is for the caller's own row.** No server function may return the
  roster or another student's data. Keep `dev/test-server.js` green; add a check when you
  add a function.
- **No third-party requests by default.** No CDN scripts, analytics, or fonts unless
  behind a Config flag that defaults off. Vendor libraries into the repo instead.
- **No emails in thrown errors or logs** (they reach Stackdriver).
- Photos: Google account photo at runtime only. Never download, copy or store an image.

## Layout

- `src/Code.gs` server · `src/Index.html` shell · `src/Styles.html` CSS · `src/App.html` client · `src/appsscript.json` manifest
- `dist/Code.gs` generated single file to paste into Apps Script (commit it after `npm run build`)
- `dev/` fake runtime, preview, builders, tests · `design/` mockups · `docs/` rubric bank + brief

## Workflow

```
npm test                                  # server checks (must pass)
npm run preview                           # dev/preview.html?role=student|student2|teacher|unknown|wrong|anon &photo=1 &fail=1
NODE_PATH=$(npm root -g) npm run smoke    # Playwright smoke + screenshots (dev/shots/, ignored)
npm run build                             # regenerate dist/Code.gs before committing
```

Before committing: `npm run check` (tests + build). Commit `dist/Code.gs` with the source
change. No dependencies to install; Node 18+ and, for smoke, a global Playwright.

## Conventions

- Plain ES5-style in `Code.gs` (Apps Script V8 runs modern JS, but keep it simple and
  `var`-based like `net-games`); modern JS is fine in `App.html`.
- Escape everything rendered with `esc()`. Client state lives in `state`; server is the
  source of truth after every save.
- Visual system: FIS maroon `#6f1d2c`, gold `#e7b64b`, Archivo + Inter (system fallback),
  the six PE-Style skins in `SKINS`. Reuse the mockups in `design/`; do not redesign.
- Rubric wording comes from `docs/G8_PE_Rubric_Bank.md` (single source). Generate, do not
  hand-copy, when v2 needs it.
- Same auth pattern as the `net-games` repo; `aa-dash`'s client-supplied-email pattern is
  **not** to be reused here.
