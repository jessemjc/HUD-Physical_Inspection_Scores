# FHA Multifamily Inspection Score Lookup

A self-contained tool for looking up HUD physical inspection scores by
FHA number, REMS property ID, or property name. Rebuilds itself weekly
from HUD's published datasets and is served live via GitHub Pages.

## One-time setup

1. **Create the repo.** Push all of these files to a new GitHub repository
   (public repo recommended -- keeps everything on the free tier, and the
   underlying data is already public HUD information).

2. **Turn on GitHub Pages.**
   - Go to the repo's **Settings → Pages**
   - Under "Build and deployment", set **Source** to **GitHub Actions**
   - That's it -- no branch to pick, the workflow handles deployment.

3. **Run it once manually.**
   - Go to the **Actions** tab → "Update FHA Inspection Score Lookup" →
     **Run workflow** (this is the `workflow_dispatch` trigger)
   - This confirms the download/build/deploy pipeline works end-to-end
     before you wait a week for the schedule to fire.

4. **Share the link.** After the first successful run, your tool is live at:
   ```
   https://<your-username>.github.io/<repo-name>/
   ```
   That's the one link to send to your team -- it always reflects the
   most recent successful data pull.

## How it stays current

`.github/workflows/update.yml` runs every **Monday at 13:00 UTC**
(`workflow_dispatch` also lets you trigger it manually any time from the
Actions tab). Each run:

1. Downloads the latest address and inspection score files directly from
   hud.gov
2. Converts the legacy `.xls` scores file with headless LibreOffice
3. Re-extracts and re-embeds the data into `templates/shell.html`
4. Commits the rebuilt `docs/index.html` back to the repo
5. Publishes it via GitHub Pages

If nothing changed on HUD's end, the commit step just no-ops (`git diff
--cached --quiet` catches that) rather than creating empty commits every
week.

## If HUD changes their file URLs or column names

This is the one thing that can break the pipeline. HUD occasionally
reorganizes their download pages. The build script fails loudly instead
of silently publishing bad data:

- **Wrong URL** → the `requests.raise_for_status()` call in
  `scripts/build.py` throws, and the Actions run shows as failed (you'll
  get a notification if you have GitHub Actions email/alerts on).
- **Renamed columns** → the schema check near the top of `main()` in
  `scripts/build.py` explicitly checks for the expected column names and
  exits with a clear error message rather than silently producing an
  empty or malformed lookup tool.

To fix either: go to HUD's multifamily data pages, find the new link or
column name, and update the constants at the top of `scripts/build.py`
(`ADDRESS_URL`, `SCORES_URL`, or the `required_addr_cols` /
`required_score_cols` sets).

Current source pages, if you need to re-find a link:
- Address file: https://www.hud.gov/hud-partners/multifamily-preservation
- Scores file: https://www.hud.gov/stat/mfh/inspection-scores

## Repo structure

```
.github/workflows/update.yml   -- the scheduled GitHub Action
scripts/build.py               -- downloads HUD data, rebuilds the HTML tool
templates/shell.html           -- the tool's HTML/CSS/JS shell (data gets injected into this)
docs/index.html                -- the built, deployable tool (this is what Pages serves)
```
