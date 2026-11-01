#!/usr/bin/env python3
"""
mdas_report_download.py
Automates the task recorded on 2026-11-01 on https://mda.mahadiscom.in/mda/ :

  1. Open site, enter User ID + password   (CAPTCHA is typed by YOU - it cannot be automated)
  2. Meter Data -> Meter Data Availability
  3. Filters: Attached to = Transformer, Status = ALL, Meter Type = Main Meter,
              Search MM Connection = Meter Serial Number, value = <serial>
  4. Click Search, wait for results
  5. Report type = Excel, click Generate Report, save the downloaded file

Install:
  pip install playwright
  playwright install chromium        (only if Google Chrome is not installed)

Run (examples):
  python mdas_report_download.py                       # serial "x2" as in the recording
  python mdas_report_download.py --serial x2 x3 x4     # one Excel report per value
  python mdas_report_download.py --user 99991001 --out downloads

Password: set env var MDAS_PASSWORD, or you will be asked (typing is hidden).
"""
import argparse
import csv
import getpass
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PWTimeout

BASE = "https://mda.mahadiscom.in/mda/"
SEARCH_URL = "https://mda.mahadiscom.in/mda/mda?uiActionName=getSearchDataAvailablity"


def launch(p):
    last = None
    for channel in ("chrome", None):
        try:
            kw = dict(headless=False, args=["--start-maximized"])
            if channel:
                kw["channel"] = channel
            browser = p.chromium.launch(**kw)
            return browser, browser.new_context(no_viewport=True, accept_downloads=True)
        except Exception as e:
            last = e
    raise last


def login(page, user, password):
    page.goto(BASE, wait_until="domcontentloaded")
    page.fill("#username", user)
    page.fill("#password", password)
    page.click("#textBox")
    print("\n>>> In Chrome: type the CAPTCHA and click 'Sign in'.")
    print(">>> Waiting up to 5 minutes for login...\n")
    page.wait_for_url(re.compile(r".*/mda/mda(\?.*)?$"), timeout=300_000)
    page.wait_for_load_state("domcontentloaded")
    print("Logged in.")


def open_search_page(page):
    page.goto(SEARCH_URL, wait_until="domcontentloaded")
    try:
        page.wait_for_selector("#installpointtype", timeout=8_000)
        return
    except PWTimeout:
        pass
    # Fallback: go through the menu like in the recording
    page.goto(BASE + "mda", wait_until="domcontentloaded")
    page.get_by_text("Meter Data", exact=True).first.hover()
    page.locator("a", has_text="Meter Data Availability").first.click()
    page.wait_for_selector("#installpointtype", timeout=30_000)


def set_filters(page, a):
    page.select_option("#installpointtype", label=a.attached)
    page.wait_for_timeout(800)          # dependent dropdowns may reload
    page.select_option("#statusresult", label=a.status)
    page.select_option("#metercheckmain", label=a.meter_type)
    page.select_option("#searchCriteria2", label="Meter Serial Number")


def search(page, serial):
    page.fill("#data", serial)
    # Search submits the form -> page navigates (seen in the recording; ~30 s for big results)
    with page.expect_navigation(timeout=300_000, wait_until="load"):
        page.click("#Search")
    try:
        page.wait_for_selector("text=/record\\(s\\) found/", timeout=60_000)
        txt = page.locator("text=/record\\(s\\) found/").first.inner_text()
        m = re.search(r"([\d,]+)\s*record", txt)
        return int(m.group(1).replace(",", "")) if m else None
    except PWTimeout:
        return 0


def generate_excel(page, downloads, out_dir, serial):
    page.select_option("#reportType", label="Excel")
    before = len(downloads)
    page.click("#Generate")
    print("  Generating Excel report (can take 1-2 minutes)...")
    waited = 0
    while len(downloads) == before:
        page.wait_for_timeout(500)
        waited += 0.5
        if waited > 600:
            raise TimeoutError("No download started within 10 minutes")
    dl = downloads[-1]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^A-Za-z0-9_-]", "_", serial)
    path = out_dir / f"{stamp}_{safe}_{dl.suggested_filename}"
    dl.save_as(str(path))
    return path


def main():
    ap = argparse.ArgumentParser(description="MDAS Meter Data Availability -> Excel report")
    ap.add_argument("--user", default="99991001", help="MDAS User ID")
    ap.add_argument("--serial", nargs="+", default=["x2"], help="meter serial value(s) to search")
    ap.add_argument("--attached", default="Transformer", help='"Attached to" option text')
    ap.add_argument("--status", default="ALL", help="Status option text: Active / InActive / PD/TD / ALL")
    ap.add_argument("--meter-type", default="Main Meter", help='"Main Meter" or "Check Meter"')
    ap.add_argument("--out", default="downloads", help="folder for the Excel files")
    a = ap.parse_args()

    password = os.environ.get("MDAS_PASSWORD") or getpass.getpass("MDAS password: ")
    out_dir = Path(a.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    errors = out_dir / "errors"
    errors.mkdir(exist_ok=True)
    log_path = out_dir / "run_log.csv"
    new_log = not log_path.exists()
    downloads = []

    with sync_playwright() as p, open(log_path, "a", newline="", encoding="utf-8") as lf:
        log = csv.writer(lf)
        if new_log:
            log.writerow(["time", "serial", "records_found", "status", "file_or_error"])
        browser, context = launch(p)

        def hook(pg):
            pg.on("download", lambda d: downloads.append(d))
            pg.on("dialog", lambda d: (print(f"  [site message] {d.message}"), d.accept()))

        context.on("page", hook)
        page = context.new_page()

        try:
            login(page, a.user, password)
            open_search_page(page)
            set_filters(page, a)

            for serial in a.serial:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print(f"\n== Serial: {serial}")
                try:
                    count = search(page, serial)
                    print(f"  Records found: {count}")
                    if not count:
                        log.writerow([now, serial, count, "SKIPPED", "no records"])
                        continue
                    path = generate_excel(page, downloads, out_dir, serial)
                    print(f"  Saved: {path}")
                    log.writerow([now, serial, count, "OK", str(path)])
                except Exception as e:
                    print(f"  FAILED: {e}")
                    try:
                        page.screenshot(path=str(errors / f"{re.sub(r'[^A-Za-z0-9_-]', '_', serial)}.png"))
                    except Exception:
                        pass
                    log.writerow([now, serial, "", "FAILED", str(e)[:200]])
                lf.flush()
        finally:
            try:
                browser.close()
            except Exception:
                pass

    print(f"\nDone. Files in: {out_dir.resolve()}   Log: {log_path.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
