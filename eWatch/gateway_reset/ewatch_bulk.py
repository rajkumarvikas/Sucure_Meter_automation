import os
import re
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright

BASE_URL = "https://msedclpq.ewatch.online:12008/"
INPUT_FILE = "gateway.xlsx"
RESULT_FILE = "result.xlsx"
SHOT_DIR = "screenshots"

# Keep TEST_MODE=True for a trial run: only the first 3 substations in the Excel will run
# (Dongarkinni, Shirwal, Surawadi). Set it to False to run everything.
TEST_MODE = False
TEST_SUBSTATIONS = 3

# Gateways that were already reset earlier and must not be reset again: put their numbers here
ALREADY_DONE = []  # e.g. ["Q1074237", "Q1076500"]

# If the popup shows an error (e.g. "No meter attached"), that gateway's status is written as ERROR.
# False: ERROR gateways are skipped on re-run. True: ERROR gateways are retried too.
RETRY_POPUP_ERRORS = False

# if the popup message contains any of these words, it is treated as an error
ERR_RE = re.compile(r"not sent|error|fail|unable|cannot|can't|invalid|no meter|contact admin", re.I)

os.makedirs(SHOT_DIR, exist_ok=True)

# ---------- Read Excel ----------
raw = pd.read_excel(INPUT_FILE, sheet_name="Sheet1", dtype=str)
raw = raw[["Sr. No.", "Gateway S. No.", "Substation name"]].dropna(how="all")
df = raw.dropna(subset=["Gateway S. No.", "Substation name"]).copy()
for c in df.columns:
    df[c] = df[c].str.strip()

print(f"Excel has {len(raw)} rows, {df['Gateway S. No.'].nunique()} gateways, "
      f"{df['Substation name'].nunique()} substations.")
if len(df) < len(raw):
    skipped = raw[raw[["Gateway S. No.", "Substation name"]].isna().any(axis=1)]
    print(f"WARNING: {len(skipped)} rows have a blank gateway number or substation name and will be skipped. Sr. No.: "
          + ", ".join(skipped["Sr. No."].astype(str)))
dups = df[df["Gateway S. No."].duplicated(keep=False)]
if len(dups):
    print("WARNING: these gateway numbers appear more than once in the Excel:",
          ", ".join(sorted(set(dups["Gateway S. No."]))))

COLS = ["Sr. No.", "Gateway S. No.", "Substation name", "Status", "Remark", "Time", "Screenshot"]

# results: gateway -> record (if an old result.xlsx exists, the run resumes from it)
results = {}
if os.path.exists(RESULT_FILE):
    try:
        xl = pd.ExcelFile(RESULT_FILE)
        if "Result" in xl.sheet_names:
            old = xl.parse("Result", dtype=str).fillna("")
            for _, r in old.iterrows():
                results[r["Gateway S. No."]] = r.to_dict()
            n_ok = sum(str(v.get("Status", "")).startswith("OK") for v in results.values())
            print(f"Found existing result.xlsx: {n_ok} gateways already OK (will be skipped).")
        else:
            print("Existing result.xlsx is not from this script, ignoring it (a new one will be created).")
    except Exception as e:
        print("Could not read existing result.xlsx:", e)


def is_done(gw):
    st = str(results.get(gw, {}).get("Status", ""))
    if st.startswith("OK"):
        return True
    return st.startswith("ERROR") and not RETRY_POPUP_ERRORS


def set_result(row, status, shot="", remark=""):
    results[row["Gateway S. No."]] = {
        "Sr. No.": row["Sr. No."],
        "Gateway S. No.": row["Gateway S. No."],
        "Substation name": row["Substation name"],
        "Status": status,
        "Remark": remark,
        "Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Screenshot": shot,
    }


for _gw in ALREADY_DONE:
    _r = df[df["Gateway S. No."] == _gw]
    if len(_r) and not is_done(_gw):
        set_result(_r.iloc[0], "OK", "", "Already reset in an earlier run (no popup screenshot)")


def write_excel(path, all_df):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        all_df.to_excel(w, sheet_name="Result", index=False)
        ws = w.book["Result"]
        for col, width in zip("ABCDEFG", (8, 16, 45, 12, 60, 20, 50)):
            ws.column_dimensions[col].width = width
        # summary (to the right of the table, columns I:J)
        st = all_df["Status"].astype(str)
        for i, (k, v) in enumerate([
            ("Total gateway", len(all_df)),
            ("Total substation", all_df["Substation name"].nunique()),
            ("OK", int(st.str.startswith("OK").sum())),
            ("ERROR (popup)", int(st.str.startswith("ERROR").sum())),
            ("FAIL", int(st.str.startswith("FAIL").sum())),
            ("PENDING", int((st == "PENDING").sum())),
        ], start=1):
            ws.cell(row=i, column=9, value=k)
            ws.cell(row=i, column=10, value=v)
        ws.column_dimensions["I"].width = 18


def save_excel():
    rows = []
    for _, r in df.iterrows():
        rec = results.get(r["Gateway S. No."])
        if rec is None:
            rec = {"Sr. No.": r["Sr. No."], "Gateway S. No.": r["Gateway S. No."],
                   "Substation name": r["Substation name"], "Status": "PENDING",
                   "Remark": "", "Time": "", "Screenshot": ""}
        rows.append({c: rec.get(c, "") for c in COLS})
    all_df = pd.DataFrame(rows, columns=COLS)
    try:
        write_excel(RESULT_FILE, all_df)
    except PermissionError:
        write_excel("result_latest.xlsx", all_df)
        print(f"   WARNING: {RESULT_FILE} is open in Excel, so results were saved to result_latest.xlsx instead. Please close Excel.")


def safe_name(s):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", s).strip("_")[:80]


def open_substation_page(page):
    page.get_by_role("link", name="Data acquisition").click()
    page.get_by_role("link", name="Master definition").click()
    page.get_by_role("button", name="Substation").click()
    page.get_by_role("textbox", name="Enter keywords").wait_for(state="visible")


def popup_text(btn):
    """Message shown in the popup (read before pressing its button)."""
    label = ""
    try:
        label = btn.inner_text().strip()
    except Exception:
        pass
    xpaths = ["xpath=ancestor::*[@role='dialog' or contains(@class,'modal') or contains(@class,'popup') or contains(@class,'swal')][1]"]
    xpaths += [f"xpath=ancestor::div[{n}]" for n in range(1, 6)]
    for xp in xpaths:
        try:
            box = btn.locator(xp)
            if box.count():
                t = " ".join(box.first.inner_text().split())
                if label:
                    t = re.sub(rf"\s*{re.escape(label)}\s*$", "", t, flags=re.I).strip()
                if len(t) > 8:
                    return t[:300]
        except Exception:
            continue
    return ""


CONFIRM_TXT = "proceed with the gateway reset"

# element (and its parents) fully visible: opacity ~1, visible, not display:none
OPAQUE_JS = """el => {
  for (let n = el; n && n !== document.documentElement; n = n.parentElement) {
    const cs = getComputedStyle(n);
    if (parseFloat(cs.opacity) < 0.9 || cs.visibility === 'hidden' || cs.display === 'none') return false;
  }
  return true;
}"""


def is_shown(loc):
    """Not just 'visible': must be fully displayed (ignores hidden or fading popups)."""
    try:
        return loc.is_visible() and bool(loc.evaluate(OPAQUE_JS))
    except Exception:
        return False


def wait_until(page, cond, timeout=30000, step=200):
    waited = 0
    while waited < timeout:
        if cond():
            return True
        page.wait_for_timeout(step)
        waited += step
    return False


def safe_click(loc):
    """Normal click; if a hidden element blocks it, fall back to a JavaScript click."""
    try:
        loc.click(timeout=5000)
    except Exception:
        loc.evaluate("el => el.click()")


def click_shown_button(page, pattern):
    btns = page.get_by_role("button", name=re.compile(pattern, re.I))
    for i in range(btns.count()):
        b = btns.nth(i)
        if is_shown(b):
            safe_click(b)
            return
    raise Exception(f"button {pattern} not found")


def confirm_reset(page):
    """Reset confirm popup (Yes/No): clicks Yes and waits until the popup closes."""
    confirm = page.get_by_text(CONFIRM_TXT, exact=False).first
    if wait_until(page, lambda: is_shown(confirm), 15000):
        page.wait_for_timeout(300)
        click_shown_button(page, r"^\s*yes\s*$")
        if not wait_until(page, lambda: not is_shown(confirm), 15000):
            raise Exception("Confirm popup did not close after clicking Yes")


def wait_result_popup(page, timeout=30000):
    """Result popup after reset (Done = success, Ok = error); wait until it is fully displayed."""
    btns = page.get_by_role("button", name=re.compile(r"^\s*(ok|done)\s*$", re.I))
    found = {}

    def cond():
        for i in range(btns.count()):
            b = btns.nth(i)
            if is_shown(b):
                t = popup_text(b)
                if t and CONFIRM_TXT not in t.lower():
                    found["btn"], found["text"] = b, t
                    return True
        return False

    if not wait_until(page, cond, timeout):
        raise Exception("Result popup (Done/Ok) did not appear after reset")
    return found["btn"], found["text"]


def reach_sync_data(page, name):
    """After searching, select the substation and reach 'Sync data'."""
    sync = page.get_by_role("link", name="Sync data")

    def visible(ms):
        try:
            sync.first.wait_for(state="visible", timeout=ms)
            return True
        except Exception:
            return False

    # wait until the searched substation's row is visible (so old results do not confuse us)
    row_text = page.get_by_text(name, exact=False).first
    try:
        row_text.wait_for(state="visible", timeout=10000)
    except Exception:
        raise Exception("Substation row not visible after search")

    if not visible(3000):
        print("   'Sync data' not visible, clicking the substation row")
        row = page.locator("tr, [role='row']").filter(has_text=name).first
        (row if row.count() else row_text).click()
        if not visible(4000):
            print("   Row click did not help, trying the three-dot (⋮) menu")
            row = page.locator("tr, [role='row']").filter(has_text=name).first
            row.locator("button, a, [role='button'], .dropdown-toggle, i, span").last.click()
            sync.first.wait_for(state="visible", timeout=8000)

    page.wait_for_timeout(800)
    n = sync.count()
    if n != 1:
        raise Exception(f"search returned {n} substations (expected 1), skipping for safety")
    sync.click()
    print("   Opened Sync data")


def search_substation(page, name):
    """Types the keyword and clicks Search; retries up to 3 times if the list is not filtered."""
    box = page.get_by_role("textbox", name="Enter keywords")
    btn = page.get_by_role("button", name="Search")

    # let the default list finish loading first, otherwise it overwrites the search result later
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:
        pass
    page.wait_for_timeout(1500)

    target = page.get_by_text(name, exact=False).first
    for attempt in (1, 2, 3):
        box.click()
        box.fill("")
        box.press_sequentially(name, delay=25)  # real typing so the site receives input events
        btn.click()
        try:
            target.wait_for(state="visible", timeout=3000)
        except Exception:
            box.press("Enter")  # some sites only search on Enter
        try:
            target.wait_for(state="visible", timeout=5000)
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass
            return
        except Exception:
            print(f"   search attempt {attempt}: list not filtered, retrying")
    raise Exception("Substation result did not appear even after clicking Search")


def process_substation(page, idx, name, group):
    """Resets all gateways of one substation. Raises an exception on any problem."""
    group_gws = list(group["Gateway S. No."])
    gws = [g for g in group_gws if not is_done(g)]

    search_substation(page, name)
    print("   searched")

    reach_sync_data(page, name)

    rows = page.locator("[id^='row'][id$='Gatewaygrid']")
    rows.first.wait_for(state="visible")
    page.wait_for_timeout(500)
    total = rows.count()

    # identify each gateway by its serial number; if not found and the counts match, use row order
    found = {gw: rows.filter(has_text=gw).count() for gw in gws}
    if all(c == 1 for c in found.values()):
        mode = "serial"
    elif total == len(group_gws):
        mode = "order"
    else:
        raise Exception(f"grid has {total} gateways but Excel has {len(group_gws)}; could not match")
    print(f"   grid has {total} gateways, match mode: {mode}")

    for gw in gws:
        row = df[df["Gateway S. No."] == gw].iloc[0]
        target = rows.filter(has_text=gw) if mode == "serial" else rows.nth(group_gws.index(gw))
        target.get_by_role("link", name="reset").click()

        # 1) confirm popup (Yes/No): click Yes and wait until it closes
        confirm_reset(page)
        # 2) result popup (Done or Ok): take the screenshot only once it is fully displayed
        btn, text = wait_result_popup(page)
        page.wait_for_timeout(400)

        label = btn.inner_text().strip().lower()
        is_error = bool(ERR_RE.search(text)) or (label == "ok" and not re.search(r"success|sent", text, re.I))

        # screenshot of the popup BEFORE pressing its button
        suffix = "_ERROR" if is_error else ""
        shot = os.path.join(SHOT_DIR, f"{idx:02d}_{safe_name(name)}_{gw}{suffix}.png")
        page.screenshot(path=shot)
        safe_click(btn)
        wait_until(page, lambda: not is_shown(btn), 8000)
        page.wait_for_timeout(400)

        if is_error:
            set_result(row, "ERROR", shot, text or "Error popup appeared (message not found)")
            print(f"   {gw}: ERROR | {text}")
        else:
            remark = text or "Reset OK (popup text not found)"
            set_result(row, "OK" if mode == "serial" else "OK (by row order)", shot, remark)
            print(f"   {gw}: reset OK | {remark}")

    page.get_by_role("link", name="Back to substation").click()
    page.get_by_role("textbox", name="Enter keywords").wait_for(state="visible")


groups = list(df.groupby("Substation name", sort=False))
if TEST_MODE:
    groups = groups[:TEST_SUBSTATIONS]
    print(f"TEST_MODE: only the first {len(groups)} substations will run.")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(ignore_https_errors=True)
    page = context.new_page()
    page.set_default_timeout(60000)

    # ---- Login: do it yourself in the browser (there is a captcha) ----
    page.goto(BASE_URL)
    print("Log in in the browser (enter the captcha). The script continues once the dashboard opens.")
    print("Do not close the browser window.")
    page.wait_for_url(
        lambda u: ":12008" in u and "signin-oidc" not in u and "Account/Login" not in u,
        timeout=600000,
    )
    page.wait_for_load_state("domcontentloaded")
    print("Logged in.")

    open_substation_page(page)

    for idx, (name, group) in enumerate(groups, start=1):
        if all(is_done(g) for g in group["Gateway S. No."]):
            print(f"[{idx}/{len(groups)}] {name}: already done, skipping")
            continue
        print(f"[{idx}/{len(groups)}] {name} ({len(group)} gateways)")
        try:
            process_substation(page, idx, name, group)
        except Exception as e:
            msg = str(e).splitlines()[0][:150]
            print(f"   FAIL: {msg}")
            err_shot = os.path.join(SHOT_DIR, f"ERROR_{idx:02d}_{safe_name(name)}.png")
            try:
                page.screenshot(path=err_shot, full_page=True)
            except Exception:
                err_shot = ""
            for _, r in group.iterrows():
                if not is_done(r["Gateway S. No."]):
                    set_result(r, "FAIL", err_shot, msg)
            # bring the page back to the right place for the next substation
            try:
                page.goto(BASE_URL)
                page.wait_for_load_state("domcontentloaded")
                open_substation_page(page)
            except Exception as e2:
                print("   Could not recover the page, stopping the script:", str(e2).splitlines()[0])
                save_excel()
                break
        save_excel()

    save_excel()
    browser.close()

ok = sum(str(v.get("Status", "")).startswith("OK") for v in results.values())
err = sum(str(v.get("Status", "")).startswith("ERROR") for v in results.values())
fail = sum(str(v.get("Status", "")).startswith("FAIL") for v in results.values())
total = len(df)
print("-" * 45)
print(f"TOTAL gateway : {total}  ({df['Substation name'].nunique()} substation)")
print(f"OK            : {ok}")
print(f"ERROR (popup) : {err}")
print(f"FAIL          : {fail}")
print(f"PENDING       : {total - ok - err - fail}")
print(f"Excel: {RESULT_FILE} | screenshots folder: {SHOT_DIR}")
