# Design references (mockups, not shipped)

Static mockups the app is productionised from. Open them in a browser. All
people shown are fictional sample data.

| File | What it shows | Status in the app |
|---|---|---|
| `G8_Build_Your_Card.html` | Name + PE-Style picker with live card reskin | **v1** — the "Build my card" screen in `src/App.html` (name comes from sign-in, not typed) |
| `PE_Athlete_Card.html` | The full collectible card: foil tiers (Bronze/Silver/Gold), radar with September ghost line, stats, badges, PNG download | v3 (radar from real data) · v4 (tiers, badges) · download deferred until html2canvas is vendored |
| `My_PE_Profile_mockup.html` | The full profile: at-a-glance tiles, fitness, fundamental-skill ladders, game skills, participation trend, report bands, reflection, detail overlays | **v1** — the main screen, every section in its empty "fills in at the Combine" state; data arrives v2–v4 |

Visual system reused by the app: FIS maroon `#6f1d2c` + gold `#e7b64b`, display font
Archivo (the profile mockups use Barlow Condensed; the brief specifies Archivo, so the
app uses Archivo with a system fallback), body font Inter, the six PE-Style colour
skins, the hexagon radar, the 4-segment level bars.

Note the mockups load Google Fonts and `html2canvas` from CDNs. The app does not do
either by default: fonts are behind the Config flag `web_fonts` (off), and download will
be added only with the library vendored into the repo.
