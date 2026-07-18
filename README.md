# RSMS Activity Points Tracker

A small automation tool for students at Rajagiri (RSET) to pull together
every Activity Point submission scattered across the RSMS portal's
Class Code x Category dropdown grid, and get one Excel file with a real
summary: total approved points, a breakdown by category, and a list of
what's still pending.

## Why this exists

The RSMS "Activity Point Form" only shows submissions for one Class Code and Category combination at a time, with no overview page. Checking your
total points means clicking through every combination by hand. This
script does that for you and builds a formatted report.

## What it does

1. Opens a real Chrome window and takes you to the RSMS login page.
2. Waits for you to log in with Google yourself (this can't and shouldn't
   be automated — no credentials are stored or entered by the script).
3. Once you land on the student home page, it opens the Activity Point
   Form and reads every Class Code and Category option.
4. For each combination, it clicks **"Add Activity"** — which only
   *reveals* the entry form and any existing submissions table, it does
   **not** submit anything — and scrapes the results table if one appears.
5. Writes everything to `activity_points.xlsx`:
   - **Summary** sheet — total approved points, points broken down into
     Professional / Extracurricular / Leadership (auto-bucketed from each
     entry's own category label), a listing of approved submissions with
     their point value, and a listing of pending submissions.
   - **ActivityPoints** sheet — every scraped row, colour-coded by status
     (green = approved, yellow = pending). Fields that mean the same
     thing but are labeled differently per category (e.g. "Name of the
     organizing instituition", "Name of the offering agency", "Organized
     By...") are merged into one shared column instead of each spawning
     its own.
   - **Skipped** sheet — only appears if a combination failed after
     retries, so you can see what to re-check.

Progress is saved to disk after *every* combination, so an interruption
never loses what's already been scraped.

## Important: read before running

- **This only reads data.** The script only ever clicks "Add Activity",
  which just displays your existing submissions — it never touches the
  actual `SUBMIT` button that would create a new activity point entry.
- **Google login is manual, on purpose.** A visible browser window opens
  and waits for you to sign in yourself. Nothing about your credentials
  is captured or stored.
- **This automates your own account for your own use.** Be mindful of
  your institution's terms of use before running bulk automation against
  any college system, and keep the request pace reasonable (the default
  delay between requests is deliberately conservative).
- Column names and page structure are based on the current RSMS layout
  as of writing. If the site changes, see the Troubleshooting section.

## Setup

Requires Python 3.8+.

```bash
pip install playwright openpyxl
playwright install chromium
```

On Windows, if `playwright` isn't recognized as a command after
installing, run it through Python instead:

```bash
python -m pip install playwright openpyxl
python -m playwright install chromium
```

## Usage

```bash
python activity_points_tracker.py
```

- A Chrome window opens to the RSMS login page — log in with Google.
- The script waits (no timeout) until you reach the student home page,
  then takes over from there automatically.
- Progress prints to the console as it works through each combination.
- When it finishes (or if you stop it early), open `activity_points.xlsx`
  — the Summary tab is first.


## Files

| File | Purpose |
|---|---|
| `activity_points_tracker.py` | Main script — login, scrape, format, save |
| `README.md` | This file |

## License

For personal / educational use. Not affiliated with or endorsed by
Rajagiri School of Engineering & Technology.
