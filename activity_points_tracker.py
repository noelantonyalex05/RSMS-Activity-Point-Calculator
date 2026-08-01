"""
Rajagiri Activity Points Tracker
=================================

Logs into the RSMS student portal, sweeps every Class Code x Category
combination on the Activity Point Form, scrapes whatever submissions
exist for each, and writes everything to a nicely formatted Excel file -
including a Summary sheet with total points, a per-category breakdown,
and pending / rejected submission lists.

Confirmed site behavior:
- Login page: studentlogin/login.php. Google sign-in requires a human,
  so you log in manually in the browser window this script opens.
- Activity.asp has exactly two <select> elements before you interact
  with anything: [0] = Class Code, [1] = Category.
- Clicking "Add Activity" is READ-ONLY - it reveals an entry form AND,
  below it, a results table (if any submissions exist for that
  class+category). It does NOT create a new record. The real
  record-creating action is a separate "SUBMIT" button inside the
  revealed form, which this script never touches.
- Different categories have genuinely different table columns (e.g.
  Sports/Games has "Level"/"Points" where Leadership has "Documentary
  evidence"/"Rating By Faculty"). Columns are matched BY NAME so every
  value lands in the right place regardless of which category it came
  from.
- Occasionally a combo's dropdowns stop populating mid-run - a known
  RSMS session glitch. When this happens repeatedly on the same combo,
  the script automatically closes the browser, has you log in again,
  and resumes from that exact combo (instead of skipping it), up to a
  small number of attempts before finally giving up on it.

Output: <STUDENT NAME>.xlsx (the name is read from Home.asp's "Logged In
User : ..." line; falls back to activity_points.xlsx if it can't be
found), saved after every single combination so progress is never lost
if the script errors out or is interrupted. If the target file already
exists it's simply replaced - unless it's currently locked (e.g. open in
Excel), in which case a new file with a numeric suffix (e.g.
"..._1.xlsx") is used instead rather than crashing.
  - Sheet "Summary"       - total points, category breakdown (approved /
                             pending / rejected counts), pending list,
                             rejected list
  - Sheet "ActivityPoints" - every scraped row, one column per unique
                             field name encountered across all categories
  - Sheet "Skipped"       - only created if a combo failed after retries

Install once:
    pip install playwright openpyxl
    playwright install chromium
    (if 'playwright' isn't recognized as a command on Windows:
     python -m playwright install chromium)

Run:
    python activity_points_tracker.py
"""

import os
import re
import time

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

# ----------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------

LOGIN_URL = "https://rajagiritech.ac.in/stud/KTU/Student/studentlogin/login.php"
# Base/fallback output filename. At runtime this is personalized with the
# logged-in student's name (see build_output_filename), e.g.
# "activity_points_JOHN_DOE.xlsx". This constant is only used as-is if the
# student's name can't be found on Home.asp.
OUTPUT_FILE = "activity_points.xlsx"
DELAY_SECONDS = 1.0            # pause between combinations, be polite to the server
OPTIONS_WAIT_TIMEOUT = 15      # seconds to wait for a <select>'s options to populate
MAX_RETRIES_PER_COMBO = 1      # attempts before a combo is logged to Skipped
MAX_RELOGIN_ATTEMPTS_PER_COMBO = 3  # fresh-login retries before a stuck combo is skipped

# Keyword -> bucket name, matched case-insensitively against the table's
# own "Category" column value (e.g. "Professional and Co-curricular").
BUCKET_KEYWORDS = [
    ("Professional", "professional"),
    ("Extracurricular", "extracurricular"),
    ("Leadership", "leadership"),
]

# ----------------------------------------------------------------------
# Styling
# ----------------------------------------------------------------------

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

TITLE_FONT = Font(bold=True, size=16, color="1F4E78")
SUBHEAD_FONT = Font(bold=True, size=12, color="1F4E78")
KPI_LABEL_FONT = Font(bold=True, size=11)
KPI_VALUE_FONT = Font(bold=True, size=14, color="2E7D32")

APPROVED_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
PENDING_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
REJECTED_FILL = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")

_thin = Side(style="thin", color="CCCCCC")
THIN_BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

# Different categories label the same underlying field differently.
# Any header on the right maps to the canonical name on the left, so they
# share one Excel column instead of each spawning its own.
COLUMN_ALIASES = {
    'Evidence / Certificate': [
        'Certificate/Letter from authorities (File Size<500kb)',
        'Certificate (File Size<500kb)',
        'Certificate/ Documentary evidence (File Size<500kb)',
        'MOOC with final exam and assessment certificate in the class 2026S7CU',
        'Certificate/Letter from Authorities/Documentary evidence (File Size<500kb)',
        'Documentary evidence and photo of product if any (File Size<500kb)',
        'Documentary evidence (File Size<500kb)',
    ],
    'Organizing Institution / Society / Company': [
        'Name of the organizing instituition and Place',
        'Name of the offering agency',
        'Name of proffesional society',
        'Organized By (name of institution with place)',
        'Organized by (name of institution with place)',
        'Name of the Company and Address',
        'StartUp Company (Registered Legally)',
        'Nameof company',
        'Awarding agency',
        'Leadership and Management - Society/Association/Club in the class 2026S7CU',
        'Club activities',
        'Name of professional society/ Association name/Name of Club etc.',
    ],
}

_VARIANT_TO_CANONICAL = {
    variant: canonical
    for canonical, variants in COLUMN_ALIASES.items()
    for variant in variants
}


def canonical_field_name(name):
    """Map a scraped column header to its canonical name, if it's a known
    alias. Unknown headers pass through unchanged."""
    return _VARIANT_TO_CANONICAL.get(name, name)


def merge_field(row_dict, name, value):
    """Add (name, value) to row_dict under its canonical name. If that
    canonical field is already set in this row from a different alias
    (shouldn't normally happen - each category table only uses one
    variant - but handled just in case), the two values are concatenated
    instead of one silently overwriting the other."""
    canonical = canonical_field_name(name)
    existing = row_dict.get(canonical)
    if not existing:
        row_dict[canonical] = value
    elif value and value != existing:
        row_dict[canonical] = f"{existing}; {value}"


class ExcelBuilder:
    """Writes rows to a sheet, creating a new column the first time a
    field name is seen and reusing it after that - so different
    categories' differently-shaped tables all line up correctly."""

    def __init__(self, wb, sheet_name):
        self.wb = wb
        self.ws = wb.active
        self.ws.title = sheet_name
        self.header_to_col = {}
        self.next_col = 1
        self.widths = {}
        self.data_row = 1
        self.ws.freeze_panes = "A2"

    def get_col(self, name):
        if name not in self.header_to_col:
            col = self.next_col
            self.header_to_col[name] = col
            self._set_cell(1, col, name, header=True)
            self.next_col += 1
        return self.header_to_col[name]

    def _set_cell(self, row, col, value, header=False, fill=None):
        cell = self.ws.cell(row=row, column=col, value=value)
        if header:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
        else:
            cell.border = THIN_BORDER
            if fill:
                cell.fill = fill
        length = len(str(value)) if value is not None else 0
        self.widths[col] = max(self.widths.get(col, 0), length)
        return cell

    def write_row(self, row_dict, fill=None):
        self.data_row += 1
        row = self.data_row
        for name, value in row_dict.items():
            col = self.get_col(name)
            self._set_cell(row, col, value, fill=fill)
        return row

    def autofit(self):
        for col, w in self.widths.items():
            letter = get_column_letter(col)
            self.ws.column_dimensions[letter].width = min(max(w + 2, 10), 60)


def write_skipped_sheet(wb, skipped_rows):
    """(Re)writes the Skipped sheet from scratch. No-op if nothing failed."""
    if "Skipped" in wb.sheetnames:
        del wb["Skipped"]
    if not skipped_rows:
        return
    ws = wb.create_sheet("Skipped")
    headers = ["Class Code", "Category", "Error"]
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
    for ri, row in enumerate(skipped_rows, 2):
        for ci, v in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=v).border = THIN_BORDER
    for ci in range(1, 4):
        ws.column_dimensions[get_column_letter(ci)].width = 34


# ----------------------------------------------------------------------
# Summary calculation
# ----------------------------------------------------------------------

def bucket_for(category_value):
    if not category_value:
        return "Other / Unclassified"
    lower = str(category_value).lower()
    for bucket_name, keyword in BUCKET_KEYWORDS:
        if keyword in lower:
            return bucket_name
    return "Other / Unclassified"


def to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s or s == "-":
        return None
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return None
    return float(m.group())


def find_points_value(row_dict):
    """Sum any field whose name contains 'point' or 'rating' (excluding
    the 'Point Status' text field) - this is where the numeric point
    value lives, under whatever name that category's table uses for it."""
    total = 0.0
    found_any = False
    for name, value in row_dict.items():
        name_l = name.lower()
        if name_l == "point status":
            continue
        if "point" in name_l or "rating" in name_l:
            n = to_number(value)
            if n is not None:
                total += n
                found_any = True
    return total if found_any else 0.0


def find_rejection_reason(row_dict):
    """Best-effort lookup of a rejection reason / remark field, since
    different category tables may label this differently (or not have
    one at all)."""
    for name, value in row_dict.items():
        name_l = name.lower()
        if "remark" in name_l or "reason" in name_l:
            if value:
                return value
    return ""


class SummaryStats:
    def __init__(self):
        self.bucket_totals = {}
        self.bucket_counts = {}
        self.bucket_pending = {}
        self.bucket_rejected = {}
        self.total_approved_points = 0.0
        self.total_approved_count = 0
        self.total_rejected_count = 0
        self.pending_rows = []
        self.approved_rows = []
        self.rejected_rows = []

    def accumulate(self, row_dict):
        """Update running totals from one scraped row. Returns the row's
        status ('approved' / 'pending' / 'rejected' / other, lowercased)
        so the caller can color the row accordingly."""
        status = str(row_dict.get("Point Status", "") or "").strip().lower()
        category_value = row_dict.get("Category")
        bucket = bucket_for(category_value)
        points = find_points_value(row_dict)

        if status == "approved":
            self.bucket_totals[bucket] = self.bucket_totals.get(bucket, 0.0) + points
            self.bucket_counts[bucket] = self.bucket_counts.get(bucket, 0) + 1
            self.total_approved_points += points
            self.total_approved_count += 1
            self.approved_rows.append({
                "Class Code": row_dict.get("Class Code"),
                "Selected Category": row_dict.get("Selected Category"),
                "Category": category_value,
                "Activity": row_dict.get("Activity"),
                "Name of event": row_dict.get("Name of event"),
                "Points": points,
            })
        elif status == "pending":
            self.bucket_pending[bucket] = self.bucket_pending.get(bucket, 0) + 1
            self.pending_rows.append({
                "Class Code": row_dict.get("Class Code"),
                "Selected Category": row_dict.get("Selected Category"),
                "Category": category_value,
                "Activity": row_dict.get("Activity"),
                "Name of event": row_dict.get("Name of event"),
            })
        elif status == "rejected":
            self.bucket_rejected[bucket] = self.bucket_rejected.get(bucket, 0) + 1
            self.total_rejected_count += 1
            self.rejected_rows.append({
                "Class Code": row_dict.get("Class Code"),
                "Selected Category": row_dict.get("Selected Category"),
                "Category": category_value,
                "Activity": row_dict.get("Activity"),
                "Name of event": row_dict.get("Name of event"),
                "Reason": find_rejection_reason(row_dict),
            })

        return status

    def ordered_buckets(self):
        all_buckets = (
            set(self.bucket_totals)
            | set(self.bucket_counts)
            | set(self.bucket_pending)
            | set(self.bucket_rejected)
        )
        ordered = [b for b, _ in BUCKET_KEYWORDS if b in all_buckets]
        ordered += sorted(b for b in all_buckets if b not in ordered)
        return ordered

    def write_sheet(self, wb):
        """(Re)builds the Summary sheet from current totals, as the first tab."""
        if "Summary" in wb.sheetnames:
            del wb["Summary"]
        ws = wb.create_sheet("Summary", 0)

        ws.merge_cells("A1:D1")
        title = ws.cell(row=1, column=1, value="Activity Points Summary")
        title.font = TITLE_FONT
        ws.row_dimensions[1].height = 26

        r = 3
        ws.cell(row=r, column=1, value="Total Approved Points").font = KPI_LABEL_FONT
        ws.cell(row=r, column=2, value=self.total_approved_points).font = KPI_VALUE_FONT
        r += 1
        ws.cell(row=r, column=1, value="Total Approved Submissions").font = KPI_LABEL_FONT
        ws.cell(row=r, column=2, value=self.total_approved_count).font = Font(bold=True, size=12)
        r += 1
        ws.cell(row=r, column=1, value="Pending Submissions").font = KPI_LABEL_FONT
        pending_cell = ws.cell(row=r, column=2, value=len(self.pending_rows))
        pending_cell.font = Font(bold=True, size=12, color="BF8F00")
        r += 1
        ws.cell(row=r, column=1, value="Rejected Submissions").font = KPI_LABEL_FONT
        rejected_cell = ws.cell(row=r, column=2, value=self.total_rejected_count)
        rejected_cell.font = Font(bold=True, size=12, color="C00000")
        r += 2

        ws.cell(row=r, column=1, value="Points by Category").font = SUBHEAD_FONT
        r += 1
        headers = ["Category", "Approved Points", "Approved Count", "Pending Count", "Rejected Count"]
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
        r += 1
        for b in self.ordered_buckets():
            ws.cell(row=r, column=1, value=b).border = THIN_BORDER
            ws.cell(row=r, column=2, value=self.bucket_totals.get(b, 0.0)).border = THIN_BORDER
            ws.cell(row=r, column=3, value=self.bucket_counts.get(b, 0)).border = THIN_BORDER
            ws.cell(row=r, column=4, value=self.bucket_pending.get(b, 0)).border = THIN_BORDER
            ws.cell(row=r, column=5, value=self.bucket_rejected.get(b, 0)).border = THIN_BORDER
            r += 1
        r += 1

        ws.cell(row=r, column=1,
                 value=f"Approved Submissions ({len(self.approved_rows)})").font = SUBHEAD_FONT
        r += 1
        aheaders = ["Class Code", "Selected Category", "Category", "Activity", "Name of event", "Points"]
        for ci, h in enumerate(aheaders, 1):
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
        r += 1
        for ar in self.approved_rows:
            for ci, key in enumerate(aheaders, 1):
                cell = ws.cell(row=r, column=ci, value=ar.get(key))
                cell.fill = APPROVED_FILL
                cell.border = THIN_BORDER
            r += 1
        r += 1

        ws.cell(row=r, column=1,
                 value=f"Pending Submissions ({len(self.pending_rows)})").font = SUBHEAD_FONT
        r += 1
        pheaders = ["Class Code", "Selected Category", "Category", "Activity", "Name of event"]
        for ci, h in enumerate(pheaders, 1):
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
        r += 1
        for pr in self.pending_rows:
            for ci, key in enumerate(pheaders, 1):
                cell = ws.cell(row=r, column=ci, value=pr.get(key))
                cell.fill = PENDING_FILL
                cell.border = THIN_BORDER
            r += 1
        r += 1

        ws.cell(row=r, column=1,
                 value=f"Rejected Submissions ({len(self.rejected_rows)})").font = SUBHEAD_FONT
        r += 1
        rheaders = ["Class Code", "Selected Category", "Category", "Activity", "Name of event", "Reason"]
        for ci, h in enumerate(rheaders, 1):
            cell = ws.cell(row=r, column=ci, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
        r += 1
        for rr in self.rejected_rows:
            for ci, key in enumerate(rheaders, 1):
                cell = ws.cell(row=r, column=ci, value=rr.get(key))
                cell.fill = REJECTED_FILL
                cell.border = THIN_BORDER
            r += 1

        widths = [26, 34, 34, 26, 30, 30]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


# ----------------------------------------------------------------------
# Output filename handling
# ----------------------------------------------------------------------

def sanitize_filename_part(text):
    """Strip characters that aren't safe in Windows/Mac/Linux filenames
    and collapse whitespace, so a student's name can be used as part of
    a filename."""
    if not text:
        return ""
    text = text.strip().strip('"').strip()
    text = re.sub(r'[\\/:*?"<>|]', "", text)
    text = re.sub(r"\s+", "_", text)
    return text


def get_student_name(page):
    """Best-effort scrape of the logged-in student's name from Home.asp,
    where the page shows something like: Logged In User : JOHN DOE

    This site is a classic-ASP frameset - the nav bar showing the name
    typically lives in its own frame, not the top-level document - so we
    search the text of every frame, not just page.inner_text("body").
    Returns None if the pattern can't be found anywhere, so callers can
    fall back to a generic filename instead of failing the whole run.
    """
    texts = []
    try:
        texts.append(page.inner_text("body"))
    except Exception:
        pass
    for frame in page.frames:
        try:
            texts.append(frame.inner_text("body"))
        except Exception:
            continue
    body_text = "\n".join(texts)

    # "Logged In User" may or may not have a space before the colon, and
    # the name may or may not be wrapped in quotes, so handle both.
    m = re.search(r'LOGGED\s+IN\s+USER\s*:\s*"([^"]+)"', body_text, re.IGNORECASE)
    if not m:
        m = re.search(r"LOGGED\s+IN\s+USER\s*:\s*([^\n\r]+)", body_text, re.IGNORECASE)
    if not m:
        print(f">>> (debug) scanned {len(page.frames)} frame(s) but found no "
              f"'Logged In User' text. Portal layout may differ from expected.")
        return None

    name = m.group(1).strip().strip('"').strip()
    return name or None


def build_output_filename(student_name):
    """Personalize the output filename with the student's name when
    available, e.g. 'JOHN_DOE.xlsx'. Falls back to the generic
    OUTPUT_FILE name if the name couldn't be found."""
    safe_name = sanitize_filename_part(student_name)
    if safe_name:
        _, ext = os.path.splitext(OUTPUT_FILE)
        return f"{safe_name}{ext}"
    return OUTPUT_FILE


def can_write_to(path):
    """True if `path` can be opened for writing right now - i.e. it
    doesn't exist yet, or it exists but isn't locked by another program
    (like having the file open in Excel)."""
    try:
        with open(path, "a"):
            pass
        return True
    except OSError:
        return False


def resolve_output_path(base_path):
    """Return a path that's safe to write to for this run.
    - If base_path doesn't exist yet, or exists but is writable, it's
      used as-is (a normal save() will simply replace/overwrite it).
    - If it exists and is locked (e.g. currently open in Excel), the
      first available "<base>_1.xlsx", "<base>_2.xlsx", ... is used
      instead, so the run never crashes over this."""
    if not os.path.exists(base_path) or can_write_to(base_path):
        return base_path
    stem, ext = os.path.splitext(base_path)
    n = 1
    while True:
        candidate = f"{stem}_{n}{ext}"
        if not os.path.exists(candidate) or can_write_to(candidate):
            return candidate
        n += 1


def save_workbook(wb, path):
    """Save wb to path, replacing it if it already exists. If the file
    can't be written to right now (e.g. it's open in Excel), falls back
    to a new filename with a numeric suffix instead of crashing, and
    returns whatever path was actually used so the caller can keep using
    it for subsequent saves this run."""
    try:
        wb.save(path)
        return path
    except PermissionError:
        stem, ext = os.path.splitext(path)
        n = 1
        new_path = path
        while True:
            candidate = f"{stem}_{n}{ext}"
            if not os.path.exists(candidate):
                new_path = candidate
                break
            n += 1
        print(f">>> Couldn't save to '{path}' (likely open in another "
              f"program). Saving to '{new_path}' instead.")
        wb.save(new_path)
        return new_path


# ----------------------------------------------------------------------
# Playwright automation
# ----------------------------------------------------------------------

class DropdownNotPopulatedError(Exception):
    """Raised when a <select>'s options don't populate in time. On this
    portal that's a specific, recognizable session glitch on the
    website's end - reloading the same page doesn't fix it, but a fresh
    login does - so it's kept distinct from other/generic failures so
    main() can react to it differently."""
    pass


def get_select_options(select_locator):
    """Return list of (value, label) for every real option in a <select>."""
    options = select_locator.locator("option").all()
    result = []
    for opt in options:
        value = opt.get_attribute("value")
        label = opt.inner_text().strip()
        if value is None or value == "":
            continue
        result.append((value, label))
    return result


def wait_for_options_populated(select_locator, min_options=1, timeout_s=OPTIONS_WAIT_TIMEOUT):
    """Poll until the <select> has more than min_options options, or raise."""
    deadline = time.time() + timeout_s
    count = 0
    while time.time() < deadline:
        count = select_locator.locator("option").count()
        if count > min_options:
            return
        time.sleep(0.2)
    raise DropdownNotPopulatedError(
        f"Select did not populate options within {timeout_s}s (had {count})"
    )


def scrape_one_combo(page, form_url, class_val, cat_val):
    """Navigate to the form, select both dropdowns, click Add Activity
    (read-only - never clicks SUBMIT), and return (header_cells, rows)."""
    page.goto(form_url)
    page.wait_for_load_state("load")

    class_select = page.locator("select").nth(0)
    category_select = page.locator("select").nth(1)
    wait_for_options_populated(class_select)
    wait_for_options_populated(category_select)

    class_select.select_option(class_val, timeout=10000)
    category_select.select_option(cat_val, timeout=10000)

    page.click("text=Add Activity")
    page.wait_for_load_state("load")
    time.sleep(0.5)  # small buffer for classic ASP postback rendering

    results_table = page.locator("table", has_text="Sl.No")
    if results_table.count() == 0:
        return None, []

    rows = results_table.first.locator("tr").all()
    header_cells = None
    scraped = []
    for row in rows:
        cells = [c.inner_text().strip() for c in row.locator("td, th").all()]
        if not cells:
            continue
        if cells[0] == "Sl.No":
            if header_cells is None:
                header_cells = cells
            continue  # header row, not data
        scraped.append(cells)

    return header_cells, scraped


def login_and_open_form(p):
    """Launches a fresh browser, waits for the user to log in manually
    with Google, opens the Activity Point Form, and confirms both
    dropdowns are populated. Returns everything needed to (re)start
    scraping: (browser, context, page, form_url, student_name,
    class_options, category_options).

    Used both for the initial login and to recover from the RSMS session
    glitch where a combo's dropdowns stop populating - closing the
    browser and logging in fresh is, per the site's usual behavior, what
    clears it."""
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    page.goto(LOGIN_URL)
    print(">>> Please log in with Google manually.")
    page.wait_for_url("**/Home.asp", timeout=0)
    print(">>> Login detected.")
    time.sleep(1.5)  # let the frameset's child frames finish loading

    student_name = get_student_name(page)
    if student_name:
        print(f">>> Logged in as: {student_name}")
    else:
        print(">>> Couldn't find the student's name on Home.asp - "
              "using a generic output filename instead.")

    print(">>> Navigating to Activity Point Form...")
    page.click("text=Activity Point Form")
    page.wait_for_load_state("load")
    form_url = page.url

    class_select = page.locator("select").nth(0)
    category_select = page.locator("select").nth(1)
    wait_for_options_populated(class_select)
    wait_for_options_populated(category_select)

    class_options = get_select_options(class_select)
    category_options = get_select_options(category_select)

    return browser, context, page, form_url, student_name, class_options, category_options


def login_and_open_form_with_retry(p, attempts=2):
    """Same as login_and_open_form, but if the dropdowns still don't
    populate right after a completely fresh login (rare, but possible if
    the portal itself is down rather than just this session), tries
    logging in again a couple of times before giving up for good."""
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return login_and_open_form(p)
        except DropdownNotPopulatedError as e:
            last_error = e
            print(f">>> Form dropdowns still empty right after a fresh "
                  f"login (attempt {attempt}/{attempts}): {e}")
    raise last_error


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    wb = openpyxl.Workbook()
    builder = ExcelBuilder(wb, "ActivityPoints")
    builder.get_col("Class Code")
    builder.get_col("Selected Category")

    stats = SummaryStats()
    skipped_rows = []

    with sync_playwright() as p:
        (browser, context, page, form_url, student_name,
         class_options, category_options) = login_and_open_form_with_retry(p)

        output_path = resolve_output_path(build_output_filename(student_name))

        print(f">>> Found {len(class_options)} class codes, "
              f"{len(category_options)} categories "
              f"({len(class_options) * len(category_options)} combinations total)")

        combos = [(c, a) for c in class_options for a in category_options]

        i = 0
        relogin_attempts_this_combo = 0
        while i < len(combos):
            (class_val, class_label), (cat_val, cat_label) = combos[i]
            last_error = None
            header_cells = None
            scraped_rows = None
            dropdown_failure = False

            for attempt in range(1, MAX_RETRIES_PER_COMBO + 1):
                try:
                    header_cells, scraped_rows = scrape_one_combo(
                        page, form_url, class_val, cat_val
                    )
                    last_error = None
                    dropdown_failure = False
                    break
                except DropdownNotPopulatedError as e:
                    last_error = e
                    dropdown_failure = True
                    print(f"  [{i + 1}/{len(combos)}] class={class_label} "
                          f"category={cat_label} -> attempt {attempt} failed "
                          f"(dropdown didn't populate): {e}")
                    time.sleep(1.5)
                except Exception as e:
                    last_error = e
                    dropdown_failure = False
                    print(f"  [{i + 1}/{len(combos)}] class={class_label} "
                          f"category={cat_label} -> attempt {attempt} failed: {e}")
                    time.sleep(1.5)

            # Empty dropdowns after every retry is the known RSMS session
            # glitch - closing the browser and logging back in usually
            # clears it. Retry the SAME combo afterwards rather than
            # skipping it, up to a small cap so a combo that's genuinely
            # broken (not a session issue) doesn't loop forever.
            if (dropdown_failure and last_error is not None
                    and relogin_attempts_this_combo < MAX_RELOGIN_ATTEMPTS_PER_COMBO):
                relogin_attempts_this_combo += 1
                print(f">>> Dropdowns came back empty after {MAX_RETRIES_PER_COMBO} "
                      f"attempts on class={class_label}, category={cat_label}. "
                      f"This is the known RSMS session glitch - closing the "
                      f"browser and logging in again "
                      f"(relogin {relogin_attempts_this_combo}/"
                      f"{MAX_RELOGIN_ATTEMPTS_PER_COMBO} for this combination)...")
                try:
                    browser.close()
                except Exception:
                    pass

                (browser, context, page, form_url, _,
                 class_options, category_options) = login_and_open_form_with_retry(p)

                print(f">>> Logged back in. Resuming from class={class_label}, "
                      f"category={cat_label}...")
                continue  # retry the same combo, i stays put

            relogin_attempts_this_combo = 0  # moving on, one way or another

            if last_error is not None:
                reason = str(last_error)
                if dropdown_failure:
                    reason += (f" (persisted through "
                               f"{MAX_RELOGIN_ATTEMPTS_PER_COMBO} relogin attempt(s))")
                print(f"  [{i + 1}/{len(combos)}] SKIPPING class={class_label} "
                      f"category={cat_label} after {MAX_RETRIES_PER_COMBO} attempts")
                skipped_rows.append([class_label, cat_label, reason])
                write_skipped_sheet(wb, skipped_rows)
                output_path = save_workbook(wb, output_path)
                i += 1
                time.sleep(DELAY_SECONDS)
                continue

            if not scraped_rows:
                print(f"  [{i + 1}/{len(combos)}] class={class_label} "
                      f"category={cat_label} -> no submissions")
            else:
                for cells in scraped_rows:
                    row_dict = {"Class Code": class_label, "Selected Category": cat_label}
                    if header_cells:
                        for name, value in zip(header_cells, cells):
                            merge_field(row_dict, name, value)
                    else:
                        for idx, value in enumerate(cells, 1):
                            merge_field(row_dict, f"Col{idx}", value)

                    status = stats.accumulate(row_dict)
                    fill = (
                        APPROVED_FILL if status == "approved" else
                        PENDING_FILL if status == "pending" else
                        REJECTED_FILL if status == "rejected" else
                        None
                    )
                    builder.write_row(row_dict, fill=fill)

                print(f"  [{i + 1}/{len(combos)}] class={class_label} "
                      f"category={cat_label} -> {len(scraped_rows)} row(s)")

            builder.autofit()
            stats.write_sheet(wb)
            output_path = save_workbook(wb, output_path)
            i += 1
            time.sleep(DELAY_SECONDS)

        builder.autofit()
        stats.write_sheet(wb)
        output_path = save_workbook(wb, output_path)

        print(f"\n>>> Done. Total approved points: {stats.total_approved_points}")
        print(f">>> Pending submissions: {len(stats.pending_rows)}")
        print(f">>> Rejected submissions: {stats.total_rejected_count}")
        print(f">>> Saved to {output_path}")

        browser.close()


if __name__ == "__main__":
    main()