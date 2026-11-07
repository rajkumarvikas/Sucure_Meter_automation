"""
ProQ Meter Removal - Excel to website automation

Install:
    python -m pip install playwright openpyxl
    python -m playwright install chromium

Run from the folder containing your Excel file:
    python removal.py

Optional Windows environment variables:
    set FFS_USER=138VIKASK
    set FFS_PASS=your_password
    set FFS_EXCEL=97_washim.xlsx
    set FFS_SUB_COL=Name & Address of Substation

Safety:
- TEST_LIMIT = None processes the complete workbook; set a number for a limited test.
- The script does not automatically retry a submission when the website's
  confirmation popup times out. Check the website first to avoid duplicates.
"""

import getpass
import os
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout


# ------------------------- CONFIGURATION -------------------------

BASE = "https://ffsaws.fieldservices.online:1070"

USERNAME = os.environ.get("FFS_USER", "138VIKASK")
PASSWORD = os.environ.get("FFS_PASS")

SERVICE = "Meter Removal"
JOB_TYPE = "ProQ Meter Removal"
DEFAULT_DEVICE_TYPE = "Power Transformer"

EXCEL_FILE = os.environ.get("FFS_EXCEL")
SUB_COL_OVERRIDE = os.environ.get("FFS_SUB_COL")

RESULT_FILE = "result.xlsx"
JOBIDS_FILE = "job_ids.xlsx"

HEADLESS = False
SLOW_MO = 100                 # Set to 0 for normal speed
LOADING_TIMEOUT = 90_000
SELECT_TIMEOUT = 15_000
MAX_ROWS_TO_TRY = 1  # One exact meter per Excel row

# IMPORTANT: Test with one row first. Set to None only after confirming.
TEST_LIMIT = None  # Process the complete workbook

OUTPUT_DIR = Path(__file__).resolve().parent


# ------------------------- ERRORS -------------------------

class SubmissionUnconfirmed(Exception):
    """The Save button was clicked but no confirmation was received."""


# ------------------------- GENERAL HELPERS -------------------------

def output_path(filename):
    return OUTPUT_DIR / filename


def safe_screenshot(page, filename):
    """Save a screenshot without allowing screenshot failure to hide the real error."""
    try:
        path = output_path(filename)
        page.screenshot(path=str(path), full_page=True, timeout=15_000)
        print(f"Screenshot saved: {path}")
        return str(path)
    except Exception as exc:
        print(f"WARNING: Could not save screenshot {filename}: {exc}")
        return None


def wait_idle(page):
    """Wait until visible loading overlays have disappeared."""
    page.wait_for_timeout(300)
    page.wait_for_function(
        """() => [...document.querySelectorAll('.loading')].every(el => {
            const rect = el.getBoundingClientRect();
            const style = getComputedStyle(el);
            return style.display === 'none'
                || style.visibility === 'hidden'
                || (rect.width === 0 && rect.height === 0);
        })""",
        timeout=LOADING_TIMEOUT,
    )


def select2_pick(page, container_id, search_text=None, option_text=None):
    """Choose an item from a Select2 dropdown and verify the selection."""
    container = page.locator(f"#{container_id}")
    container.wait_for(state="visible", timeout=10_000)
    container.click()

    if search_text:
        search = page.locator(
            ".select2-search--dropdown .select2-search__field:visible"
        ).first
        search.wait_for(state="visible", timeout=5_000)
        search.fill(search_text)

        option = page.locator(
            ".select2-results__option:visible"
            ":not(.loading-results):not(.select2-results__message)",
            has_text=re.compile(re.escape(search_text), re.I),
        ).first

        try:
            option.wait_for(state="visible", timeout=SELECT_TIMEOUT)
        except PWTimeout:
            safe_screenshot(page, "error_substation_nooption.png")
            raise PWTimeout(
                f"No Select2 result found for substation '{search_text}'."
            )

        option.click()

    elif option_text:
        option = page.locator(
            ".select2-results__option:visible", has_text=option_text
        ).first
        option.wait_for(state="visible", timeout=SELECT_TIMEOUT)
        option.click()
    else:
        raise ValueError("Provide either search_text or option_text.")

    try:
        page.wait_for_function(
            """id => {
                const el = document.getElementById(id);
                const text = el?.textContent?.trim();
                return !!text && text.toLowerCase() !== 'select';
            }""",
            arg=container_id,
            timeout=5_000,
        )
    except PWTimeout:
        safe_screenshot(page, "error_select_not_applied.png")
        raise PWTimeout(f"Selection did not apply in #{container_id}.")


# ------------------------- LOGIN / FORM SETUP -------------------------

def login(page):
    page.goto(f"{BASE}/Login", wait_until="domcontentloaded")
    page.locator("#txtUsername").fill(USERNAME)
    page.locator("#txtPassword").fill(PASSWORD)
    page.locator("form button:has-text('Login')").click()
    page.wait_for_url(re.compile(r".*/Index"), timeout=30_000)
    print("Login successful.")


def tick_hierarchy(page):
    """Open the hierarchy and tick the MSED root node."""
    wait_idle(page)

    dropdown = page.locator(
        "#electricaljobs > div:nth-of-type(1) > "
        "div:nth-of-type(4) > div:nth-of-type(1)"
    )
    dropdown.wait_for(state="visible", timeout=15_000)

    node = page.locator(".jqx-tree-item:visible", has_text="MSED").first

    opened = False
    for attempt in range(4):
        if attempt % 2 == 0:
            dropdown.click()
        else:
            bounds = dropdown.bounding_box()
            if bounds:
                page.mouse.click(
                    bounds["x"] + bounds["width"] - 14,
                    bounds["y"] + bounds["height"] / 2,
                )
        try:
            node.wait_for(state="visible", timeout=4_000)
            opened = True
            break
        except PWTimeout:
            continue

    if not opened:
        safe_screenshot(page, "error_hierarchy.png")
        raise PWTimeout("Could not open the MSED hierarchy dropdown.")

    ticked = node.evaluate(
        """el => {
            const li = el.closest('li');
            const tree = el.closest('.jqx-tree') || li?.closest('[id]');
            try {
                const jq = window.jQuery(tree);
                jq.jqxTree('checkItem', li, true);
                return jq.jqxTree('getCheckedItems').length > 0;
            } catch (e) {
                return false;
            }
        }"""
    )

    if not ticked:
        item = page.locator("li:visible", has=node).last
        checkbox = item.locator(
            ".jqx-checkbox:visible, input[type=checkbox]:visible"
        ).first

        if checkbox.count():
            checkbox.click()
        else:
            bounds = node.bounding_box()
            if not bounds:
                raise RuntimeError("Could not locate the MSED checkbox.")
            page.mouse.click(
                bounds["x"] - 14,
                bounds["y"] + bounds["height"] / 2,
            )

    page.wait_for_timeout(500)
    safe_screenshot(page, "after_tick.png")

    page.keyboard.press("Escape")
    page.locator("body").click(position={"x": 5, "y": 400})
    wait_idle(page)


def open_create_job_form(page):
    page.goto(f"{BASE}/JobManagement/CreateJob", wait_until="domcontentloaded")
    page.locator("#tabelectricaljobs").click()

    select2_pick(
        page,
        "select2-ddlElecservice-container",
        option_text=SERVICE,
    )
    page.locator("#ddlElecjob").select_option(label=JOB_TYPE)
    tick_hierarchy(page)


# ------------------------- DEVICE CHECKBOX HANDLING -------------------------


def get_device_rows(page):
    """
    Return only device data rows, excluding the grid header/master checkbox.
    Each data row is identified by its Meter serial number cell.
    """
    return page.locator("tr").filter(
        has=page.locator('input[type="checkbox"][name="group1[]"]')
    )


def normalize_meter(value):
    """Normalize meter serials without changing their meaningful digits."""
    return re.sub(r"[^A-Z0-9]", "", str(value or "").strip().upper())


def select_only_meter_row(page, meter_serial):
    """
    Find the exact meter serial in the device grid and select only that row.
    The header checkbox is never included because selection is row-scoped.
    """
    wanted = normalize_meter(meter_serial)
    if not wanted:
        raise ValueError("Excel meter serial number is blank.")

    # Locate a table row containing the exact meter serial in a cell.
    matching_rows = page.locator("tr").filter(
        has=page.locator("td").filter(
            has_text=re.compile(rf"^\s*{re.escape(str(meter_serial).strip())}\s*$", re.I)
        )
    )

    # Filter further to rows that actually have a device checkbox.
    candidates = []
    for i in range(matching_rows.count()):
        row = matching_rows.nth(i)
        if row.locator('input[type="checkbox"][name="group1[]"]').count():
            candidates.append(row)

    if len(candidates) != 1:
        # Fallback: inspect each checkbox's closest table row and match row text.
        rows = page.locator("tr")
        candidates = []
        for i in range(rows.count()):
            row = rows.nth(i)
            cb = row.locator('input[type="checkbox"][name="group1[]"]')
            if cb.count() and row.locator("td").count() > 0 and normalize_meter(row.inner_text()).find(wanted) >= 0:
                candidates.append(row)

    if len(candidates) != 1:
        visible_rows = []
        rows = page.locator("tr")
        for i in range(rows.count()):
            row = rows.nth(i)
            if row.locator('input[type="checkbox"][name="group1[]"]').count():
                visible_rows.append(row.inner_text().strip().replace("\n", " | "))
        raise RuntimeError(
            f"Expected exactly one grid row for meter {meter_serial!r}; "
            f"found {len(candidates)}. Device rows: {visible_rows}"
        )

    target_row = candidates[0]
    target_cb = target_row.locator(
        'input[type="checkbox"][name="group1[]"]'
    ).first

    # Only operate on device-row checkboxes. The header select-all checkbox
    # is in a different row and is deliberately excluded.
    rows = page.locator("tr")
    device_checkboxes = []
    for i in range(rows.count()):
        row = rows.nth(i)
        cb = row.locator('input[type="checkbox"][name="group1[]"]')
        if cb.count() and row.locator("td").count() > 0:
            device_checkboxes.append(cb.first)

    # Clear all device selections, then select the exact target.
    for cb in device_checkboxes:
        if cb.evaluate("(el) => el.checked"):
            cb.evaluate("(el) => el.click()")

    if not target_cb.evaluate("(el) => el.checked"):
        target_cb.evaluate("(el) => el.click()")

    selected = []
    for i, cb in enumerate(device_checkboxes):
        if cb.evaluate("(el) => el.checked"):
            selected.append(i)

    if len(selected) != 1:
        raise RuntimeError(
            f"Expected only meter {meter_serial} selected, "
            f"but selected device checkbox indexes are {selected}."
        )

    # Confirm the selected row still contains the exact requested serial.
    if wanted not in normalize_meter(target_row.inner_text()):
        raise RuntimeError(f"Selected row does not match meter {meter_serial}.")


def create_one_job(page, substation_search, device_type, meter_serial):
    select2_pick(
        page,
        "select2-ddlElecSubstation-container",
        search_text=substation_search,
    )
    wait_idle(page)

    page.locator("#ddlPT_Feeder").select_option(label=device_type)
    page.locator("#btnviewSystemData").click()
    wait_idle(page)

    # The screenshot shows the grid columns:
    # PowerTransformer_id | Meter serial number | Substation name | ...
    print(f"    Searching grid for meter: {meter_serial}")
    select_only_meter_row(page, meter_serial)
    print(f"    Selected exact meter: {meter_serial}")

    popup = page.locator("#popupMessage")
    page.locator("#btnSaveSystemData").click()

    try:
        popup.wait_for(state="visible", timeout=LOADING_TIMEOUT)
        message = popup.inner_text().strip()
    except PWTimeout:
        safe_screenshot(page, "error_save_confirmation_missing.png")
        raise SubmissionUnconfirmed(
            "Save was clicked but no confirmation appeared. "
            "Check the website manually before running this meter again."
        )

    close_button = page.locator(
        "#popupBox > div:nth-of-type(1) > span:nth-of-type(2)"
    )
    close_button.click()
    popup.wait_for(state="hidden", timeout=10_000)

    # Numeric Job ID means created; test this before other message handling.
    match = re.search(r"Job\s*id\s*[-:#]?\s*(\d+)", message, re.I)
    if match:
        return match.group(1), message, "created"

    if "already created" in message.lower():
        return None, message, "duplicate"

    return None, message, "unrecognized"


# ------------------------- EXCEL INPUT -------------------------

def find_excel():
    import glob

    if EXCEL_FILE:
        candidate = Path(EXCEL_FILE)
        if not candidate.is_absolute():
            candidate = OUTPUT_DIR / candidate
        return candidate

    patterns = ("*97*washim*.xlsx", "*washim*.xlsx")
    skip_terms = ("job_ids", "result", "~$")

    for pattern in patterns:
        for filename in sorted(glob.glob(str(OUTPUT_DIR / pattern))):
            if not any(term in filename.lower() for term in skip_terms):
                return Path(filename)

    return OUTPUT_DIR / "substations.xlsx"



def load_sheet(path):
    """Read substation, meter serial and existing Job ID from the input workbook."""
    path = Path(path)
    if not path.exists():
        sys.exit(f"Excel file not found: {path}")

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
    finally:
        workbook.close()

    rows = [
        row for row in rows
        if any(value is not None and str(value).strip() for value in row)
    ]
    if not rows:
        sys.exit("Excel file is empty.")

    headers = [str(v).strip() if v is not None else "" for v in rows[0]]
    normalized = [h.lower() for h in headers]

    def find_col(predicate, description):
        for idx, header in enumerate(normalized):
            if predicate(header):
                return idx
        sys.exit(f"Could not identify the {description} column. Headers: {headers}")

    if SUB_COL_OVERRIDE:
        override = SUB_COL_OVERRIDE.strip().lower()
        if override in normalized:
            sub_col = normalized.index(override)
        elif len(override) == 1 and override.isalpha():
            sub_col = ord(override.upper()) - ord("A")
        else:
            sys.exit(f"FFS_SUB_COL '{SUB_COL_OVERRIDE}' does not match a header.")
    else:
        sub_col = find_col(
            lambda h: "substation" in h or ("name" in h and "address" in h),
            "substation",
        )

    meter_col = find_col(
        lambda h: "meter" in h and ("serial" in h or "number" in h),
        "meter serial number",
    )

    job_col = next(
        (i for i, h in enumerate(normalized) if h.strip() in ("job id", "jobid")),
        None,
    )

    print(f"Substation column: {headers[sub_col]!r}")
    print(f"Meter column: {headers[meter_col]!r}")
    if job_col is not None:
        print(f"Existing Job ID column: {headers[job_col]!r}")

    items = []
    for excel_row_num, row in enumerate(rows[1:], start=2):
        row = row + [None] * max(0, len(headers) - len(row))
        sub = row[sub_col]
        meter = row[meter_col]
        existing_job = row[job_col] if job_col is not None else None

        if sub is None or not str(sub).strip():
            continue
        if meter is None or not str(meter).strip():
            print(f"WARNING: Excel row {excel_row_num} has no meter serial; skipping.")
            continue

        items.append({
            "sub": str(sub).strip(),
            "dev": DEFAULT_DEVICE_TYPE,
            "meter": str(meter).strip(),
            "existing_job": existing_job,
            "row": row,
            "excel_row": excel_row_num,
        })

    if not items:
        sys.exit("No valid substation + meter rows found in the Excel file.")
    return headers, items, job_col


# ------------------------- EXCEL OUTPUT -------------------------

def save_workbook_safely(workbook, filename):
    target = output_path(filename)
    try:
        workbook.save(target)
        return target
    except PermissionError:
        alternate = target.with_name(
            f"{target.stem}_new{target.suffix}"
        )
        workbook.save(alternate)
        print(
            f"WARNING: {target.name} is probably open in Excel. "
            f"Saved instead to {alternate.name}."
        )
        return alternate


def save_job_ids(headers, items, outcomes, job_col):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Job IDs"
    status_headers = headers if job_col is not None else headers + ["Job id"]
    worksheet.append(status_headers + ["Automation status", "Automation remark"])

    for item, outcome in zip(items, outcomes):
        job_id, status = outcome if outcome else ("", "Not processed")
        out_row = list(item["row"])
        if job_col is not None:
            if job_id:
                out_row[job_col] = job_id
        else:
            out_row.append(job_id or "")
        worksheet.append(out_row + [status.split(":", 1)[0], status])

    for column in worksheet.columns:
        width = max(
            (len(str(cell.value)) if cell.value is not None else 0)
            for cell in column
        )
        worksheet.column_dimensions[column[0].column_letter].width = min(
            max(width + 2, 10), 45
        )

    return save_workbook_safely(workbook, JOBIDS_FILE)


def save_results(results):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Result"
    worksheet.append([
        "Excel row", "Substation", "Meter serial number",
        "Status", "Job id", "Remark"
    ])

    for result in results:
        worksheet.append(list(result))

    return save_workbook_safely(workbook, RESULT_FILE)


def save_progress(headers, items, outcomes, results, job_col):
    """Persist progress after each input row."""
    save_results(results)
    save_job_ids(headers, items, outcomes, job_col)


# ------------------------- MAIN -------------------------


def main():
    global PASSWORD
    if not PASSWORD:
        PASSWORD = getpass.getpass(f"Password for {USERNAME}: ")
    if not PASSWORD:
        sys.exit("No password entered.")

    excel_path = find_excel()
    headers, items, job_col = load_sheet(excel_path)
    total = len(items)
    print(f"Loaded {total} valid meter row(s) from {excel_path.name}")

    run_items = items[:TEST_LIMIT] if TEST_LIMIT else items
    if TEST_LIMIT:
        print(f"TEST MODE: processing only {len(run_items)} row(s).")
    else:
        print("FULL MODE: processing every valid meter row in the workbook.")

    results = []
    outcomes = [None] * total
    created_ids = []
    browser = None

    # Keep existing Job IDs and never submit those rows again.
    for idx, item in enumerate(items):
        if item["existing_job"] is not None and str(item["existing_job"]).strip():
            existing = str(item["existing_job"]).strip()
            if existing.endswith(".0"):
                existing = existing[:-2]
            outcomes[idx] = (existing, "SKIPPED: Existing Job ID")

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
            context = browser.new_context(
                ignore_https_errors=True,
                viewport={"width": 1530, "height": 950},
            )
            page = context.new_page()
            page.set_default_timeout(20_000)

            login(page)

            for idx, item in enumerate(run_items):
                excel_row = item["excel_row"]
                substation = item["sub"]
                meter = item["meter"]
                device = item["dev"]

                if outcomes[idx] is not None:
                    existing_id, old_remark = outcomes[idx]
                    print(f"\n[{idx+1}/{len(run_items)}] SKIP {meter}: existing Job ID {existing_id}")
                    results.append((excel_row, substation, meter, "Skipped", existing_id, old_remark))
                    save_progress(headers, items, outcomes, results, job_col)
                    continue

                print(f"\n[{idx+1}/{len(run_items)}] Excel row {excel_row}: {substation} | Meter {meter}")

                try:
                    # Re-open/reset the form for every Excel row. This also recovers
                    # the page after a meter mismatch or another row-level error.
                    open_create_job_form(page)
                    job_id, message, status = create_one_job(
                        page, substation, device, meter
                    )

                    if status == "created":
                        print(f"    CREATED - Job ID: {job_id}")
                        outcomes[idx] = (job_id, "CREATED: Job ID received")
                        created_ids.append(job_id)
                        result_status = "Created"
                    elif status == "duplicate":
                        print(f"    DUPLICATE - meter {meter} already has a job")
                        outcomes[idx] = ("", f"DUPLICATE: {message}")
                        result_status = "Duplicate"
                    else:
                        print(f"    ERROR - Unrecognized website response: {message}")
                        outcomes[idx] = ("", f"ERROR: Unrecognized website response: {message}")
                        result_status = "Error"

                    results.append((
                        excel_row, substation, meter, result_status,
                        job_id or "", message
                    ))

                except SubmissionUnconfirmed as exc:
                    # Do not retry this same row: the server may have accepted it.
                    # Record the uncertainty, then continue with the next Excel row.
                    message = str(exc)
                    print(f"    VERIFY MANUALLY; continuing to next row: {message}")
                    safe_screenshot(page, f"error_submission_{excel_row}.png")
                    outcomes[idx] = ("", f"VERIFY MANUALLY: {message}")
                    results.append((
                        excel_row, substation, meter, "Verify manually", "", message
                    ))

                except Exception as exc:
                    # Row-level failures (including exact-meter mismatch) are logged;
                    # they must not terminate the full workbook run.
                    message = f"{type(exc).__name__}: {exc}"
                    print(f"    ERROR - {message}")
                    safe_screenshot(page, f"error_{excel_row}.png")
                    outcomes[idx] = ("", f"ERROR: {message}")
                    results.append((
                        excel_row, substation, meter, "Error", "", message
                    ))

                # Persist after EVERY row so completed work is not lost.
                save_progress(headers, items, outcomes, results, job_col)

            context.close()
            browser.close()
            browser = None

    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass

    save_progress(headers, items, outcomes, results, job_col)
    created_count = sum(1 for r in results if r[3] == "Created")
    skipped_existing = sum(
        1 for o in outcomes if o and o[1].startswith("SKIPPED:")
    )
    error_count = sum(1 for r in results if r[3] == "Error")
    duplicate_count = sum(1 for r in results if r[3] == "Duplicate")
    verify_count = sum(1 for r in results if r[3] == "Verify manually")

    print("\n---------------- SUMMARY ----------------")
    print(f"Input meter rows: {total}")
    print(f"Rows handled this run: {len(results)}")
    print(f"Jobs created this run: {created_count}")
    print(f"Rows skipped with existing Job IDs: {skipped_existing}")
    print(f"Duplicate rows: {duplicate_count}")
    print(f"Error rows (see remarks): {error_count}")
    print(f"Rows requiring manual verification: {verify_count}")
    print(f"New Job IDs: {', '.join(created_ids) if created_ids else 'None'}")
    print(f"Full log: {output_path(RESULT_FILE)}")
    print(f"Updated Job ID workbook: {output_path(JOBIDS_FILE)}")
    print("------------------------------------------")


if __name__ == "__main__":
    main()
