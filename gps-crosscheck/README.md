# GPS cross-check (Keyman & Cold Weather Patrolling)

Every morning at 08:05 IST this checks the control team's Google Sheet against the two GPS portals
(rilapp and nbtrack) and the start/end photos, and verifies every exception the portals raise.

| Run on day D checks | Window | What |
|---|---|---|
| Cold Weather Patrolling, night D-1 → D | 22:00 → 08:00 | everything |
| Keyman, day D | 04:00 → 08:00 | start only |
| Keyman, day D-1 | 04:00 → 16:00 | whole day (completes yesterday's start check) |

## What it checks
- **Sheet vs GPS**: start, end, km, speed; "report not generated" when the portal has data.
- **Actual duty** (best evidence from GPS + trustworthy photo): late start, early finish, less distance,
  sustained high speed (vehicle), GPS silent periods, GPS switched on late although the patroller was on track.
- **Photos**: missing stamp, wrong date, stamp later than the WhatsApp receive time (altered), very old photo,
  invalid / far-off coordinates, blue-sky daylight on a night stamp, same picture reused (any batch, last 30 days),
  two batches photographed at the same spot within 15 minutes.
- **Portal exceptions**: each one marked CONFIRMED, FALSE (with the evidence) or CANNOT VERIFY.

Tolerances are edited in the dashboard (Settings). Defaults are in `config/howrah.yaml`.

## Parts
- `xcheck/` – the checking program (Python). `tests/` – rule tests (`python -m pytest tests -q`).
- `config/howrah.yaml` – everything specific to the division.
- `apps_script/` – results store + dashboard, pasted into a Google Sheet you own.
- `../.github/workflows/gps-crosscheck.yml` – the 08:05 schedule and the manual "Run workflow" button.

## One-time setup
1. **Results sheet + dashboard**: create a Google Sheet "GPS Cross-check – Results" → Extensions → Apps Script →
   paste `Code.gs` and add an HTML file `Dashboard` with `Dashboard.html` → run `setup` once (allow access) →
   copy the two keys from the execution log → Deploy → New deployment → Web app, *Execute as: Me*,
   *Who has access: Anyone* → copy the web app URL.
2. **GitHub secrets** (repo → Settings → Secrets and variables → Actions): `RILAPP_USER`, `RILAPP_PASSWORD`,
   `NBTRACK_USER`, `NBTRACK_PASSWORD`, `RESULTS_URL` (web app URL), `RESULTS_KEY` (from step 1).
3. Optional, for the dashboard's one-click **Run now**: a fine-grained GitHub token with *Actions: Read and write*
   on this repo, saved as script property `GITHUB_TOKEN`. Without it, Run now opens the GitHub page with the button.
4. Dashboard link: `<web app URL>?key=<DASHBOARD_KEY>`.

## Using it in another division
Copy `config/howrah.yaml` to `config/<division>.yaml` and change the sheet ID, tab names, portal logins (secret
names), valid map area. Set `DIVISION_CONFIG` in the workflow. The sheet reader finds columns by their header
text, so the sheet only needs the same column headings, not the same positions.

## Privacy
The repo is public: logs print counts only. Names, results and photos go only to the results sheet.

## Good to know
- GitHub pauses scheduled jobs in a repo with no commits for 60 days. If the dashboard stops getting new
  dates, open Actions → GPS cross-check → *Enable workflow*.
- Photo stamps are read offline with RapidOCR. A stamp it cannot read is marked *check by eye*, never guessed.
- Same-spot detection needs the same angle; two photos of one place from different angles are not caught.
