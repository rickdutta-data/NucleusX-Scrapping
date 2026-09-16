"""
Stanza Living NucleusX -- Finance Ledger scraper.

RUN THIS FROM A NORMAL TERMINAL (Command Prompt / PowerShell / Anaconda Prompt),
NOT from inside a Jupyter notebook cell.

Why: Jupyter kernels always run their own asyncio event loop in the background.
Playwright's *sync* API refuses to run inside any process that already has a
running event loop -- it checks for this on purpose and raises
"Sync API inside the asyncio loop" (this is not fixable with nest_asyncio or
similar patches). A plain terminal process has no such loop, so the sync API
works normally there.

Usage:
    1. pip install playwright pandas openpyxl python-dateutil
       playwright install chromium
    2. Edit BOOKING_UUIDS below if needed.
    3. Open a terminal, cd into this folder, run:
           python scrape_nucleus.py
    4. A browser window will open. Log in manually the first time (your
       session is cached in ./playwright_nucleusx_profile for next time).
    5. When it finishes, it writes scrape_results.json next to this script.
       Then run the companion notebook's "load + build Excel" cells.
"""

import json
import re
import traceback
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

# --------------------------------------------------------------------------
# CONFIG -- edit as needed
# --------------------------------------------------------------------------

BOOKING_UUIDS = [
    "54573506-89a7-44d6-a9c2-bdb7675dad8a",
    "1c0b8303-aab8-4643-b18d-27f901efd7fb",
    "18d04069-ab26-4a53-a30c-a9bf058fcbe7",
    "c67e94a8-4d09-42bf-bc57-79946664e5e5",
    "0cae050f-0cca-4256-96da-de7439618ce2",
    "2f984fc1-6a64-4df2-a605-8a7271486f67",
    "0fd584ed-ed33-4e78-b338-e892ae17fa56",
    "eefd885f-b35b-4cc1-90f3-50a14f2fe696",
    "77f094e8-2a2a-41a6-88f5-55513d86586b",
    "a5a44a7d-ba86-4950-9bd2-1b31f31cf376",
]

URL_TEMPLATE = "https://stanzaliving.nucleusx.io/customer-360-listing/view/{uuid}?tab=CONTRACT_DETAILS"

# Folder everything gets written to (login cache + results json)
WORK_DIR = Path(r"C:/Users/Rick.Dutta/OneDrive - STANZA LIVING/Documents/VS Code/FI Scrapping")
WORK_DIR.mkdir(parents=True, exist_ok=True)

USER_DATA_DIR = str(WORK_DIR / "playwright_nucleusx_profile")  # caches your login
RESULTS_PATH = WORK_DIR / "scrape_results.json"

# Keep this False (visible browser) until you've confirmed everything works.
HEADLESS = False

# --------------------------------------------------------------------------
# Date / number helpers
# --------------------------------------------------------------------------

MONTH_ABBR_MAP = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}
MONTH_NAME_MAP = {
    "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
    "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12,
}
MONTH_ABBRS = set(MONTH_ABBR_MAP.keys())


def parse_date_ddMonYY(s):
    """Parses strings like "1 Oct '25" or "01 Aug'26" -> date object."""
    s = s.replace("'", " ")   # turn the apostrophe into a separator, not remove it
    s = re.sub(r"\s+", " ", s).strip()
    return datetime.strptime(s, "%d %b %y").date()


def to_float(s):
    if s is None:
        return None
    s = s.replace(",", "").replace("\u20b9", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Parsing: Contract Overview + Commercials Overview
# --------------------------------------------------------------------------

def parse_contract_overview(text):
    # ADJUST IF NEEDED: matches "Contract Tenure\n1 Oct '25 - 30 Sep '26 ( 12 months )"
    m = re.search(
        r"Contract Tenure\s*\n\s*(\d{1,2}\s+\w{3}\s*'?\d{2})\s*-\s*(\d{1,2}\s+\w{3}\s*'?\d{2})",
        text,
    )
    csd = ced = None
    if m:
        csd = parse_date_ddMonYY(m.group(1))
        ced = parse_date_ddMonYY(m.group(2))

    mi = re.search(r"Move-in Date\s*\n\s*(\d{1,2}\s+\w{3}\s*'?\d{2})", text)
    move_in = parse_date_ddMonYY(mi.group(1)) if mi else None

    return csd, ced, move_in


def parse_commercials(text):
    def grab(label):
        m = re.search(label + r"\s*\n\s*\u20b9?\s*([\d,]+(?:\.\d+)?)", text)
        return to_float(m.group(1)) if m else None

    monthly_rent = grab("Monthly Rent")
    sd = grab("Security Deposit")
    maintenance = grab("Maintenance Charges")
    onboarding = grab("Onboarding Charges")
    exit_fee = grab("Exit Processing Fee")

    return monthly_rent, sd, maintenance, onboarding, exit_fee


# --------------------------------------------------------------------------
# Parsing: Finance Ledger line items
#
# Expected pattern per entry (from the screenshot):
#   31
#   Aug
#   Payment Received
#   via ALISTE 6a95115a72a7cbdf898e7d38      <- optional sub-line, ignored
#                                                unless it's a (date range)
#   ₹500
#   Balance: -₹5,000.95
# or:
#   31
#   Aug
#   Electricity Charges - Aliste Consumption
#   (01 Aug'26 - 31 Aug'26)                  <- this sub-line fills From/To
#   ₹895.34
#   Balance: -₹4,500.95
# --------------------------------------------------------------------------

def parse_finance_ledger(text):
    lines = [l.strip() for l in text.split("\n") if l.strip() != ""]
    entries = []
    current_year = None
    i = 0
    while i < len(lines):
        line = lines[i]

        m = re.match(r"^([A-Z]+)\s+(\d{4})$", line)
        if m and m.group(1) in MONTH_NAME_MAP:
            current_year = int(m.group(2))
            i += 1
            continue

        if re.match(r"^\d{1,2}$", line) and i + 1 < len(lines) and lines[i + 1] in MONTH_ABBRS:
            day = int(line)
            mon_abbr = lines[i + 1]
            i += 2

            if i >= len(lines):
                break
            item_name = lines[i]
            i += 1

            sub_line = None
            if i < len(lines) and not lines[i].startswith("\u20b9") and not lines[i].startswith("Balance"):
                sub_line = lines[i]
                i += 1

            amt = None
            if i < len(lines):
                am = re.match(r"\u20b9?\s*([\d,]+(?:\.\d+)?)", lines[i])
                if am:
                    amt = to_float(am.group(1))
                    i += 1

            if i < len(lines) and lines[i].startswith("Balance"):
                i += 1

            entry_date = None
            if current_year is not None and mon_abbr in MONTH_ABBR_MAP:
                entry_date = date(current_year, MONTH_ABBR_MAP[mon_abbr], day)

            from_date = to_date = None
            if sub_line:
                rng = re.match(
                    r"\((\d{1,2}\s*\w{3}'?\d{2})\s*-\s*(\d{1,2}\s*\w{3}'?\d{2})\)", sub_line
                )
                if rng:
                    try:
                        from_date = parse_date_ddMonYY(rng.group(1))
                        to_date = parse_date_ddMonYY(rng.group(2))
                    except ValueError as e:
                        print(f"  [warn] could not parse date range {rng.group(0)!r}: {e}")

            entries.append({
                "item": item_name,
                "amt": amt,
                "date": entry_date,
                "from": from_date,
                "to": to_date,
            })
            continue

        i += 1

    return entries


# --------------------------------------------------------------------------
# Per-booking scrape
# --------------------------------------------------------------------------

def scrape_booking(page, uuid):
    url = URL_TEMPLATE.format(uuid=uuid)
    # NOTE: "networkidle" can hang/timeout forever on apps with background
    # polling or websockets (common on dashboards like this). "load" is
    # more reliable; we then explicitly wait for content we expect to see.
    page.goto(url, wait_until="load", timeout=60000)

    try:
        page.get_by_text("Contract Overview", exact=False).first.wait_for(
            state="visible", timeout=20000
        )
    except Exception:
        snippet = page.inner_text("body")[:500]
        raise RuntimeError(
            "Timed out waiting for 'Contract Overview' to appear. "
            "You are probably NOT logged in (session expired, or you pressed "
            "Enter before finishing login), or the page structure differs "
            "from what this script expects. First 500 chars of the page:\n"
            f"{snippet}"
        )

    page.wait_for_timeout(800)

    overview_text = page.inner_text("body")
    csd, ced, move_in = parse_contract_overview(overview_text)
    monthly_rent, sd, maintenance, onboarding, exit_fee = parse_commercials(overview_text)

    if csd is None or monthly_rent is None:
        print("  [warn] Contract/Commercials overview didn't fully parse for", uuid)
        print("  [warn] first 400 chars of page text for debugging:")
        print("  " + overview_text[:400].replace("\n", " | "))

    # ADJUST IF NEEDED: click the "Finance Ledger" sub-tab
    try:
        page.get_by_text("Finance Ledger", exact=True).first.click()
    except Exception:
        pass
    page.wait_for_timeout(1200)

    # Load every entry: scroll + click any "Load more" style button until content stops growing
    prev_len = -1
    for _ in range(40):
        try:
            more_btn = page.get_by_text(re.compile("load more", re.I))
            if more_btn.count() > 0:
                more_btn.first.click()
                page.wait_for_timeout(700)
        except Exception:
            pass
        page.mouse.wheel(0, 3000)
        page.wait_for_timeout(400)
        curr_len = len(page.inner_text("body"))
        if curr_len == prev_len:
            break
        prev_len = curr_len

    ledger_text = page.inner_text("body")
    ledger_entries = parse_finance_ledger(ledger_text)

    return {
        "uuid": uuid,
        "csd": csd,
        "ced": ced,
        "move_in": move_in,
        "monthly_rent": monthly_rent,
        "sd": sd,
        "maintenance": maintenance,
        "onboarding": onboarding,
        "exit_fee": exit_fee,
        "ledger": ledger_entries,
    }


# --------------------------------------------------------------------------
# JSON serialization helpers (date objects aren't JSON-native)
# --------------------------------------------------------------------------

def d2s(d):
    return d.isoformat() if isinstance(d, date) else None


def serialize_result(res):
    out = dict(res)
    out["csd"] = d2s(out["csd"])
    out["ced"] = d2s(out["ced"])
    out["move_in"] = d2s(out["move_in"])
    out["ledger"] = [
        {
            "item": e["item"],
            "amt": e["amt"],
            "date": d2s(e["date"]),
            "from": d2s(e["from"]),
            "to": d2s(e["to"]),
        }
        for e in res["ledger"]
    ]
    return out


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    all_results = []
    errors = []

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(USER_DATA_DIR, headless=HEADLESS)
        page = context.new_page()

        # First-time login: if you're not logged in yet, the page will show a login screen.
        page.goto("https://stanzaliving.nucleusx.io/", wait_until="load", timeout=60000)
        print()
        print("=" * 70)
        print("A browser window has opened.")
        print("In THAT window: enter your phone number, wait for the OTP,")
        print("enter the OTP, and wait until you actually see your normal")
        print("dashboard/home screen (not a login or OTP screen).")
        print("Only then come back here and press Enter.")
        print("=" * 70)
        input("Press Enter once you can see your dashboard in the browser window... ")

        for uuid in BOOKING_UUIDS:
            print("Scraping", uuid)
            try:
                result = scrape_booking(page, uuid)
                all_results.append(result)
                print("  ->", len(result["ledger"]), "ledger entries found")
            except Exception:
                print("  FAILED:")
                traceback.print_exc()
                errors.append(uuid)

        context.close()

    serialized = [serialize_result(r) for r in all_results]
    RESULTS_PATH.write_text(json.dumps(serialized, indent=2))

    print()
    print(f"Done. {len(all_results)} of {len(BOOKING_UUIDS)} bookings scraped successfully.")
    print("Results written to:", RESULTS_PATH)
    if errors:
        print(f"{len(errors)} booking(s) failed:", errors)


if __name__ == "__main__":
    main()
