"""
MDAS Meter-Modem Connection auto-filler - OWN-SIM version

Uses ONLY the sims written next to each meter in the Excel (no sims from other rows).
For every meter:
  1. Take "Sim No", then "Optional Sim No" (optional column, may be blank or missing).
  2. A trailing letter such as "U" is removed; the remaining digits are used.
  3. Get Ipaddress with the full digits. If no IP is found, remove the LAST DIGIT and try Get Ipaddress again.
  4. First sim that gives an IP and passes Validate is used. If none does, the meter is marked
     "NO OWN SIM MATCHED" and the script moves on to the next meter.

How to run:
    python mdas_auto_own_sim.py
    A browser opens -> log in yourself (captcha) -> stay on the home page and press ENTER in the terminal.

Input : MDAS_pendding.xlsx
Output: MDAS_result_own_sim.xlsx (updated after every meter)
"""
import re
import json
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright

URL = "https://mda.mahadiscom.in/mda/mda"
INPUT = "MDAS_pendding.xlsx"
OUTPUT = "MDAS_result_own_sim.xlsx"

DRY_RUN = False       # True = fill the form but do NOT click Save
LIMIT = None          # e.g. 3 for a test run; None = all rows
ALLOW_NON_AMR = False # True = if no sim works, try a NON-AMR dummy connection (the site asks for Telephone/IP, so it usually fails)
BAD_FILE = "bad_sims_own_sim.json"   # list of unusable sims is saved here (delete the file to reset)
BAD_SIMS, CLAIMED, ALL_SIMS = set(), set(), []

# values taken from the recorded workflow
METER_TYPE_DLMS = "20"     # 029 - DLMS
DLMS_DEVICE_TYPE = "055"   # SECUR
DLMS_CLASS = "3"           # C
MODEM_MAKE_NEW = "18"      # Modem Make/Model (for a new modem)
PROJECT_WITH_SIM = "MSEDCL"        # when a working sim is found
PROJECT_NON_AMR = "NON-AMR"        # when no sim works (start of the option label)

MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


# ----------------------------------------------------------------- helpers
def load_input():
    raw = pd.read_excel(INPUT, dtype=str)
    raw.columns = raw.columns.str.strip()
    raw["_date"] = pd.to_datetime(pd.read_excel(INPUT)["New Meter Installation Date"], errors="coerce")
    for c in ("Sim No", "Optional Sim No"):      # sim columns may be missing -> treat as blank
        if c not in raw.columns:
            raw[c] = ""
    df = raw.dropna(subset=["Meter Sr. No."]).copy()
    df["Meter Sr. No."] = df["Meter Sr. No."].str.strip()
    df["Modem IMEI"] = "M" + df["Meter Sr. No."]
    def clean_sim(x):
        d = re.sub(r"\D", "", str(x) if pd.notna(x) else "")   # letters such as a trailing 'U' are removed
        return d if len(d) >= 15 else ""                         # 0 / short numbers = no sim
    for c in ("Sim No", "Optional Sim No"):
        df[c + " (original)"] = df[c].fillna("").astype(str).str.strip()
        df[c] = df[c].map(clean_sim)
    return df.reset_index(drop=True)


def field_after(page, label):
    """First input/select after the label cell (td). Exact text first, then 'contains'."""
    tail = "/following::*[self::input or self::select][1]"
    exact = page.locator(f"xpath=//td[normalize-space()='{label}']{tail}")
    if exact.count():
        return exact.first
    return page.locator(f"xpath=//td[contains(normalize-space(), '{label}')]{tail}").first


def select_by_text(sel, text):
    for o in sel.locator("option").all_inner_texts():
        if o.strip().upper().startswith(text.upper()):
            sel.select_option(label=o)
            return True
    return False


def site_msg(page):
    """Red message shown above the form (e.g. 'Meter Sr. No. ... already configured')."""
    try:
        body = page.locator("body").inner_text()
        m = re.search(r"screen shot\.\s*(.*?)\s*Meter Form", body, re.S | re.I)
        return " ".join(m.group(1).split())[:250] if m else ""
    except Exception:
        return ""


def pick_town(sel, circle):
    """Excel Circle Name = Town. Try 'starts with' first, then a match ignoring spaces/brackets."""
    if select_by_text(sel, circle):
        return True
    norm = lambda t: re.sub(r"[^A-Z0-9]", "", t.upper())
    for o in sel.locator("option").all_inner_texts():
        if norm(circle) and norm(circle) in norm(o):
            sel.select_option(label=o)
            return True
    return False


def reason(text):
    t = " ".join((text or "").split())
    m = re.search(r"already configured with Meter Sr\. No\. (\w+)", t, re.I)
    if m:
        return f"sim already linked to meter {m.group(1)}"
    if "does not exist" in t.lower():
        return "sim does not exist on the site"
    return t[:150]


def dismiss_alert(page):
    """Close any popup (NON AMR alert, post-save message). Returns True if something was closed."""
    closed = False
    for _ in range(4):
        found = None
        for loc in (page.get_by_role("link", name="Close", exact=True),
                    page.get_by_role("button", name=re.compile(r"^(Close|OK)$", re.I))):
            try:
                if loc.count() and loc.first.is_visible():
                    found = loc.first
                    break
            except Exception:
                pass
        if not found:
            break
        try:
            found.click(timeout=3000)
            page.wait_for_timeout(700)
            closed = True
        except Exception:
            break
    return closed




def on_login_page(page):
    c = page.get_by_text("Enter Captcha")
    return c.count() > 0 and c.first.is_visible()


def ask_login(page):
    input("Login page detected. Log in in the browser (up to the home page), then press ENTER here... ")


def click_create_new(page):
    """Click 'Create New'. If it is not visible (hover menu), use a JS click (the recording used nth(1))."""
    cn = page.get_by_role("link", name="Create New")
    for i in range(cn.count()):
        if cn.nth(i).is_visible():
            cn.nth(i).click()
            return True
    if cn.count():
        cn.nth(min(1, cn.count() - 1)).evaluate("e => e.click()")
        return True
    return False


def open_form(page):
    """Open a fresh form for each meter: close popups -> Clear if the form is open -> menu > Create New
    -> (if not found) open the base page and use the menu again. Also works right after Save."""
    mt = page.get_by_role("textbox", name="Select OR Type Meter Sr. No.")

    def form_visible():
        return mt.count() > 0 and mt.first.is_visible()

    for attempt in range(3):
        dismiss_alert(page)
        if on_login_page(page):
            ask_login(page)
        if form_visible():                       # form is already open -> clear it
            for name in ("Back", "Clear"):
                b = page.get_by_role("button", name=name)
                if b.count() and b.first.is_visible():
                    b.first.click()
                    page.wait_for_timeout(800)
            dismiss_alert(page)
            if form_visible():
                return
        if attempt >= 1:                         # menu not found -> open the base page (as in the recording)
            try:
                page.goto(URL)
            except Exception:
                pass
            page.wait_for_timeout(1500)
            if on_login_page(page):
                ask_login(page)
        mm = page.get_by_role("link", name="Meter - Modem Connection")
        if mm.count() and mm.first.is_visible():
            mm.first.click()
            page.wait_for_timeout(600)
        click_create_new(page)
        page.wait_for_timeout(1200)
        dismiss_alert(page)
        try:
            mt.first.wait_for(state="visible", timeout=8000)
            return
        except Exception:
            pass
    raise Exception("Meter-Modem form did not open. URL: " + page.url)


def mark_done(page, row, site_ip, dialogs, out, identify):
    """Meter is already entered on the site -> record DONE with its IP and sim."""
    sims = [x for x in (row["Sim No"], row["Optional Sim No"]) if x]
    sim, how = ("", "")
    if identify:
        sim, how = identify_sim(page, row, site_ip, dialogs)
    if not sim and sims:
        sim, how = f"{sims[0]} (unverified)", "could not verify which sim is on the site; see the Excel sim columns"
    out.update({"IP": site_ip, "Sim Used": sim, "Status": "DONE (already entered)",
                "Message": f"Already configured on the site. {how}"})
    return out


def identify_sim(page, row, site_ip, dialogs):
    """For an already-entered meter, use its IP to find which Excel sim is the one on the site (best effort)."""
    sims = [x for x in (row["Sim No"], row["Optional Sim No"]) if x]
    try:
        page.get_by_role("button", name="Validate & Next").first.click()
        page.wait_for_timeout(1500)
        dismiss_alert(page)
        mod = page.get_by_role("textbox", name="Select OR Type Modem Sr. No.")
        for _ in range(16):
            if mod.is_enabled():
                break
            page.wait_for_timeout(500)
        else:
            return "", ""
        mod.click()
        mod.fill(row["Modem IMEI"])
        mod.press("Enter")
        page.wait_for_timeout(1200)
        for sim in sims:
            got, ip, tel, msg = try_sim(page, sim, dialogs)
            if got and ip == site_ip:
                return sim, "IP matched (verified)"
    except Exception:
        pass
    return "", ""


def pick_date(page, dt):
    """Open the calendar, move to the target month with « (year) and < (month), then click the day."""
    page.locator(".labelsLeft > a").first.click()
    page.wait_for_timeout(500)
    hdr = page.get_by_text(re.compile(r"[A-Z][a-z]{2}\s+20\d\d")).first
    for _ in range(60):
        mm_ = re.search(r"([A-Z][a-z]{2})\s+(20\d\d)", hdr.inner_text())
        mon, yr = mm_.group(1), int(mm_.group(2))
        cur, tgt = yr * 12 + MONTHS[mon], dt.year * 12 + dt.month
        if cur == tgt:
            break
        if dt.year < yr:
            sym = "«"
        elif dt.year > yr:
            sym = "»"
        elif tgt < cur:
            sym = "<"
        else:
            sym = ">"
        page.get_by_role("link", name=sym, exact=True).first.click()
        page.wait_for_timeout(250)
    page.get_by_role("link", name=str(dt.day), exact=True).first.click()
    page.wait_for_timeout(500)


def read_ip_tel(page):
    """IP and Telephone only from the form fields (not from the site's red message)."""
    ip = tel = ""
    try:
        ip = field_after(page, "IP Address").input_value().strip()
        tel = field_after(page, "Telephone No.").input_value().strip()
    except Exception:
        pass
    if not (ip and tel):
        return "", ""
    return ip, tel


def try_sim(page, sim, dialogs):
    page.locator("#searchsimsimno").check()
    box = page.locator("#searchtextvalue")
    box.click()
    box.fill(sim)
    dialogs.clear()
    page.get_by_role("button", name="Get Ipaddress").click()
    page.wait_for_timeout(2500)
    dismiss_alert(page)
    ip, tel = read_ip_tel(page)
    return bool(ip), ip, tel, " | ".join(dialogs)


# -------------------------------------------------------------------- main
def process_row(page, row, dialogs):
    out = {"Sim Used": "", "IP": "", "Telephone": "", "Project": "", "Assigned To": "",
           "Status": "", "Message": ""}
    meter = row["Meter Sr. No."]

    open_form(page)

    # ---- 1. METER FORM
    mbox = page.get_by_role("textbox", name="Select OR Type Meter Sr. No.")
    mbox.click()
    mbox.fill(meter)
    mbox.press("Tab")
    page.wait_for_timeout(1500)
    dismiss_alert(page)

    m = re.search(r"already configured with IP Address\s*([\d\.]+)", page.locator("body").inner_text(), re.I)
    if m:
        return mark_done(page, row, m.group(1), dialogs, out, identify=True)

    mtype = field_after(page, "Meter Type")
    if mtype.is_enabled():                       # new meter: meter type must be chosen manually
        mtype.select_option(METER_TYPE_DLMS)
        page.wait_for_timeout(1000)
        if page.locator("#DLMSDeviceTypeId").is_enabled():
            page.locator("#DLMSDeviceTypeId").select_option(DLMS_DEVICE_TYPE)
        if page.locator("#DLMSMeterClassId").is_enabled():
            page.locator("#DLMSMeterClassId").select_option(DLMS_CLASS)

    dialogs.clear()
    page.get_by_role("button", name="Validate & Next").first.click()
    page.wait_for_timeout(1500)
    dismiss_alert(page)

    # ---- 2. MODEM FORM
    mod = page.get_by_role("textbox", name="Select OR Type Modem Sr. No.")
    for _ in range(40):
        if mod.is_enabled():
            break
        page.wait_for_timeout(500)
    else:
        m2 = re.search(r"already configured with IP Address\s*([\d\.]+)", page.locator("body").inner_text(), re.I)
        if m2:                                   # message appeared after Validate
            return mark_done(page, row, m2.group(1), dialogs, out, identify=False)
        out.update({"Status": "METER VALIDATE FAILED", "Message": (" | ".join(dialogs) or site_msg(page) or "Modem field did not become enabled")})
        return out
    mod.click()
    mod.fill(row["Modem IMEI"])
    mod.press("Enter")
    page.wait_for_timeout(1200)

    proj = field_after(page, "Project")
    sims = [x for x in (row["Sim No"], row["Optional Sim No"]) if x]
    tried, had_ip = [], set()
    ok = connected = False

    used = [""]                                   # the sim actually typed (may be trimmed)

    def label(sim):
        return used[0] if used[0] == sim else f"{used[0]} (last digit removed; Excel: {sim})"

    def attempt(sim, need_ip):
        """Enter the sim and Validate. True = Connection Form opened."""
        got, ip, tel, msg = try_sim(page, sim, dialogs)
        used[0] = sim
        if not got and len(sim) > 15:            # no IP -> retry Get IP with the last digit removed
            got2, ip2, tel2, msg2 = try_sim(page, sim[:-1], dialogs)
            if got2:
                got, ip, tel, msg = got2, ip2, tel2, msg2
                used[0] = sim[:-1]
        if got:
            had_ip.add(sim)
        if need_ip and not got:
            BAD_SIMS.add(sim)
            tried.append(f"{sim}: {reason(msg) or 'no IP found on the site'} (also tried with the last digit removed)")
            return False, got, ip, tel
        make = field_after(page, "Modem Make/Model")
        if make.is_enabled():                    # new modem: Make/Model must be chosen manually
            make.select_option(MODEM_MAKE_NEW)
            page.wait_for_timeout(500)
        dialogs.clear()
        page.get_by_role("button", name="Validate & Next").last.click()
        page.wait_for_timeout(1500)
        dismiss_alert(page)
        for _ in range(12):
            sv = page.get_by_role("button", name="Save")
            if sv.count() > 0 and sv.first.is_visible():     # Connection Form opened
                return True, got, ip, tel
            page.wait_for_timeout(500)
        why = " | ".join(dialogs) or site_msg(page)
        if "already configured" in why.lower():
            BAD_SIMS.add(sim)
            BAD_SIMS.add(used[0])
        tried.append(f"{sim}: {reason(why) or 'Validate fail'}")
        return False, got, ip, tel

    # Pass 1: use a sim whose IP is found on the site (on conflict, move to the next sim)
    for sim in sims:
        if sim in BAD_SIMS or sim in CLAIMED:
            tried.append(f"{sim}: already known as bad/used, skipped")
            continue
        res, got, ip, tel = attempt(sim, True)
        if res:
            connected, ok = True, True
            CLAIMED.add(sim)
            out.update({"Sim Used": label(sim), "IP": ip, "Telephone": tel})
            break

    # Pass 2: if no sim works, try a NON-AMR (dummy) connection
    if not connected and ALLOW_NON_AMR:
        for sim in sims:
            if sim in had_ip:                    # this sim has already failed
                continue
            res, got, ip, tel = attempt(sim, False)
            if res:
                connected, ok = True, False
                out.update({"Sim Used": f"{label(sim)} (NON-AMR)", "IP": ip, "Telephone": tel})
                break

    out["Message"] = (f"{len(tried)} sims tried: " + " ; ".join(tried))[:400]
    if not connected:
        out["Status"] = "NO OWN SIM MATCHED"
        return out

    pick_date(page, row["_date"])

    project_label = PROJECT_WITH_SIM if ok else PROJECT_NON_AMR
    if not select_by_text(proj, project_label):
        out.update({"Status": "PROJECT NOT FOUND", "Message": project_label})
        return out
    out["Project"] = project_label
    page.wait_for_timeout(1000)
    dismiss_alert(page)

    if not pick_town(page.locator("#town"), row["Circle Name"]):
        out.update({"Status": "TOWN NOT FOUND", "Message": "Town not found in the list: " + row["Circle Name"]})
        return out
    page.wait_for_timeout(1500)

    assigned = page.locator("#connectedTo")
    town_sel = page.locator("#town")
    sub = field_after(page, "Subdivision")

    def sub_filled():
        try:
            t = sub.evaluate("e => e.tagName == 'SELECT' ? (e.options[e.selectedIndex] || {}).text || '' : e.value")
        except Exception:
            return False
        t = (t or "").strip()
        return bool(t) and "select" not in t.lower()

    def read_opts():
        o = []
        for _ in range(10):                      # wait for options to load
            o = [x for x in assigned.locator("option").all_inner_texts()
                 if x.strip() and not x.strip().startswith("-")]
            if o:
                break
            page.wait_for_timeout(500)
        return o

    def choose(o):
        non_others = [x for x in o if "other" not in x.lower()]
        others = [x for x in o if "other" in x.lower()]
        return non_others[0] if non_others else (others[0] if others else None)

    opts = read_opts()
    pick = choose(opts)
    if not pick:
        out.update({"Status": "MANUAL: Meter Assigned To", "Message": "no option found in the dropdown"})
        return out

    filled = False
    for rnd in range(4):
        if rnd > 0:
            # REFRESH: re-select Town -> the Meter Assigned To list reloads
            town_sel.select_option(index=0)
            page.wait_for_timeout(800)
            pick_town(town_sel, row["Circle Name"])
            page.wait_for_timeout(2000)
            dismiss_alert(page)
            opts = read_opts()
            pick = choose(opts) or pick
        assigned.select_option(label=pick)
        for _ in range(10):                      # wait for Subdivision/Sub-Station to auto-fill
            if sub_filled():
                filled = True
                break
            page.wait_for_timeout(500)
        if filled:
            break
    out["Assigned To"] = pick
    if filled and rnd > 0:
        out["Message"] = f"Meter Assigned To needed {rnd} refresh(es)"
    if not filled:
        diag = []
        for lab in ("Install Point", "Subdivision", "Sub-Stations", "Feeder", "Transformer"):
            try:
                el = field_after(page, lab)
                val = el.evaluate("e => e.tagName == 'SELECT' ? ((e.options[e.selectedIndex] || {}).text || '') : e.value")
                diag.append(f"{lab}={(val or '').strip()[:30] or 'EMPTY'}{'' if el.is_enabled() else '(disabled)'}")
            except Exception:
                diag.append(f"{lab}=?")
        n_opts = len(opts)
        out.update({"Status": "SUBDIVISION NOT FILLED",
                    "Message": f"Assigned To options={n_opts}, picked={pick[:40]}; " + ", ".join(diag) + " | Not saved"})
        return out

    if DRY_RUN:
        out["Status"] = "FILLED (dry run, not saved)"
        return out

    dialogs.clear()
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(2500)
    dismiss_alert(page)
    body = page.locator("body").inner_text()
    success = re.search(r"success", body, re.I) is not None
    errs = [d for d in dialogs if "can not proceed" in d.lower() or "please" in d.lower()]
    if errs:
        out["Status"] = "SAVE ERROR"
    else:
        out["Status"] = "SAVED" if success else "CHECK (verify the Save result)"
    out["Message"] = (" | ".join(dialogs) + " " + out["Message"]).strip()
    return out


def main():
    df = load_input()
    if Path(BAD_FILE).exists():
        try:
            BAD_SIMS.update(json.loads(Path(BAD_FILE).read_text()))
            print(f"{len(BAD_SIMS)} known bad sims loaded from {BAD_FILE}")
        except Exception:
            pass
    results, done = [], set()
    if Path(OUTPUT).exists():
        old = pd.read_excel(OUTPUT, dtype=str).fillna("")
        keep = old[old["Status"].isin(["SAVED", "DONE (already entered)", "ALREADY CONFIGURED"])]
        results = keep.to_dict("records")
        done = set(keep["Meter Sr. No."])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        dialogs = []
        def on_dialog(d):
            dialogs.append(d.message)
            try:
                d.accept()
            except Exception:
                pass
        page.on("dialog", on_dialog)
        page.goto(URL)
        input("Log in in the browser (up to the home page), then press ENTER here... ")

        todo = df[~df["Meter Sr. No."].isin(done)]
        if LIMIT:
            todo = todo.head(LIMIT)

        fails = 0
        for i, row in todo.iterrows():
            try:
                out = process_row(page, row, dialogs)
            except Exception as e:
                out = {"Sim Used": "", "IP": "", "Telephone": "", "Project": "", "Assigned To": "",
                       "Status": "SCRIPT ERROR", "Message": " ".join(str(e).split())[:450]}
                print("  ERROR:", out["Message"])
                try:
                    page.screenshot(path=f"error_{row['Meter Sr. No.']}.png")
                except Exception:
                    pass
            if not (out["Status"].startswith(("SAVED", "FILLED", "DONE")) or out["Status"] == "SCRIPT ERROR"):
                try:
                    page.screenshot(path=f"fail_{row['Meter Sr. No.']}.png")
                except Exception:
                    pass
            rec = {"Meter Sr. No.": row["Meter Sr. No."], "Circle": row["Circle Name"],
                   "Sim No (original)": row["Sim No (original)"],
                   "Optional Sim (original)": row["Optional Sim No (original)"],
                   "Install Date": row["_date"].strftime("%d-%b-%y") if pd.notna(row["_date"]) else "", **out}
            results.append(rec)
            fails = fails + 1 if rec["Status"] == "SCRIPT ERROR" else 0
            print(i + 1, rec["Meter Sr. No."], rec["Status"], rec["Sim Used"])
            pd.DataFrame(results).to_excel(OUTPUT, index=False)
            Path(BAD_FILE).write_text(json.dumps(sorted(BAD_SIMS)))
            if fails >= 5:
                print("Stopped after 5 consecutive SCRIPT ERRORs. Check the Message column and the error_*.png files.")
                break
        browser.close()
    print("Done ->", OUTPUT)


if __name__ == "__main__":
    main()
