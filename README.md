# RSMS Activity Points Tracker

A small automation tool for students at Rajagiri (RSET) to pull together
every Activity Point submission scattered across the RSMS portal's
Class Code x Category dropdown grid, and get one Excel file with a real
summary: total approved points, a breakdown by category, and a list of
what's still pending.

## Why this exists

The RSMS "Activity Point Form" only shows submissions for one Class Code
+ Category combination at a time, with no overview page. Checking your
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

## How the Excel output stays correctly aligned

Different categories use genuinely different table columns on the site —
Sports/Games might have "Level" and "Points" columns where Leadership has
"Documentary evidence" and "Rating By Faculty". The tracker matches
columns **by name**, not position: a shared field like "Start Date"
always lands in the same Excel column no matter which category it came
from, and a category-specific field just gets its own new column instead
of overwriting something unrelated.

Point values are similarly auto-detected: any field whose name contains
"point" or "rating" (other than the status field itself) is treated as
the numeric point value for that row, since the exact label varies by
category.

On top of that, some fields are the *same concept* across categories but
use a different label each time — for example "Name of professional
society...", "Name of the organizing instituition and Place", "Name of
the offering agency", "Name of the Company and Address", and "Organized
By (...)" are all really "who ran this". These known aliases are merged
into one canonical column (see `COLUMN_ALIASES` near the top of the
script) instead of each producing a separate mostly-empty column.

### Finding more aliases: `discover_column_aliases.py`

`COLUMN_ALIASES` was built by hand from a first scrape, which only shows
field names for categories you've actually submitted to. To find field
names for *every* category — including ones with zero submissions — this
script reveals each category's entry form (never touches SUBMIT) and
reads off its field labels directly:

```bash
python discover_column_aliases.py
```

It prints every category's field labels, saves the raw mapping to
`form_fields_by_category.json`, and — using a conservative keyword-based
rule set — proposes merges for fields it's confident mean the same thing,
writing them to `proposed_column_aliases.py` as a ready-to-paste
`COLUMN_ALIASES` dict. Anything it isn't confident about (e.g. "Name of
event" vs "Name of fest" — genuinely different fields, not just a label
difference) is listed separately as unclustered rather than guessed at,
since a wrong merge silently loses data. Review the unclustered list by
hand; if you find a real duplicate, add a keyword rule to
`CANONICAL_RULES` in the discovery script and re-run, or just add the
entry to `COLUMN_ALIASES` in the main script directly.

## Troubleshooting

**A dropdown/button isn't found, or selectors seem wrong for your
version of the site:**

```bash
python inspect_activity_page.py
```

This opens the same login flow, then dumps every `<select>`'s id/name and
options, every button's id/name/value, and the full page HTML to
`activity_page.html` — all without needing DevTools (which the site
blocks). Use this to see what actually changed.

**A specific Class Code + Category combination keeps failing:**
It gets retried automatically, then logged to the `Skipped` sheet with
the error message so you know what to check by hand.

## Files

| File | Purpose |
|---|---|
| `activity_points_tracker.py` | Main script — login, scrape, format, save |
| `discover_column_aliases.py` | Optional — scans every category's entry form to find fields worth merging in `COLUMN_ALIASES` |
| `inspect_activity_page.py` | Optional debugging tool — dumps raw page structure |
| `README.md` | This file |

## License

For personal / educational use. Not affiliated with or endorsed by
Rajagiri School of Engineering & Technology.
