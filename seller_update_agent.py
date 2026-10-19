# -*- coding: utf-8 -*-
"""
CartUp Seller Panel - Category Update Agent
===========================================

Automates, for every row of an Excel file (columns: Product Id, Category Path):

  sign in (ID + password filled automatically) -> OTP pasted in terminal
  -> Seller Information -> Seller List
  -> search Seller Code (2nd search bar, Pending / Approve / Reject tabs)
  -> Switch to Seller icon (data-pr-tooltip="Switch to seller") -> Yes
  -> Product -> Products -> Offline tab -> Rejected tab
  -> for each Product Id: Edit -> walk category cascade from Category Path
     (every column searched with 1-2 short words, click when the option
     shows) -> Main Materials* text bar: type 'Other' -> Update
     -> if Main Material / Warranty notifications appear, fill them
     ('Other' / 'No Warranty' in Warranty Type) and Update again -> next row

Prompts in this terminal:  OTP code (typed here, agent clicks Verify),
Seller Code, Excel path.

Run with run_agent.bat  (keeps the window open on errors)
"""

import csv
import os
import re
import sys
import time
import traceback

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ---------------------------------------------------------------------------
# Tunables - adjust here during calibration
# ---------------------------------------------------------------------------

START_URL = "https://seller-admin.cartup.com/auth/signin?callbackUrl=%2F"

# Seller panel credentials - entered automatically on the sign-in page.
# Plain text on purpose (local machine only); edit here if the password changes.
USERNAME = "Nafim_Hoque"
PASSWORD = "Nafim@123"

AGENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(AGENT_DIR, "user_data")     # persistent login profile
SHOT_DIR = os.path.join(AGENT_DIR, "screenshots")      # PNG per step / on error

HEADLESS = False            # keep the browser visible so you can watch it
SLOW_MO = 80                # ms delay between Playwright actions (0 = off)
NAV_TIMEOUT = 45_000        # page navigation timeout
DEFAULT_TIMEOUT = 20_000    # default wait timeout
SEARCH_TIMEOUT = 8_000      # waiting for a seller row after searching
TOAST_WAIT = 6_000          # waiting for the Update confirmation

# Sidebar / tab labels - change here if the site wording differs
LBL_SELLER_INFO = ["Seller Information", "Seller information"]
LBL_SELLER_LIST = ["Seller List", "Seller list", "Sellers"]
TABS_SHOP = [("Pending", "pending"), ("Approve", "approve"), ("Reject", "reject")]
TABS_OFFLINE = ["Offline"]
TABS_REJECTED = ["Rejected"]

MAIN_MATERIALS = "Main Materials"        # looked up as a regex, case-insensitive
MATERIALS_CHOICE = "Other"               # typed into the Main Material text bar

WARRANTY_TYPE = r"Warranty\s*Type"       # warranty type label (regex)
WARRANTY_VALUE = "No Warranty"           # typed into the Warranty Type text bar

# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def norm(text):
    """Lowercase, collapse whitespace, strip - for forgiving text matching."""
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip().lower()


def shot(page, name):
    """Save a screenshot; never crash because of it."""
    try:
        os.makedirs(SHOT_DIR, exist_ok=True)
        stamp = time.strftime("%H%M%S")
        path = os.path.join(SHOT_DIR, "%s_%s.png" % (stamp, name))
        page.screenshot(path=path, full_page=False)
        print("    [shot] %s" % os.path.basename(path))
    except Exception:
        pass


def ask(prompt):
    """input() that also works if stdin is closed (returns '')."""
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return ""


def click_text(scope, texts, exact=False, timeout=DEFAULT_TIMEOUT, desc=""):
    """Click the first visible element whose text matches one of `texts`."""
    deadline = time.time() + timeout / 1000
    last_err = None
    for attempt in range(60):
        if time.time() > deadline:
            break
        for t in texts:
            try:
                loc = scope.get_by_text(t, exact=exact)
                for i in range(loc.count()):
                    el = loc.nth(i)
                    try:
                        if el.is_visible():
                            el.click(timeout=4000)
                            return True
                    except Exception as e:
                        last_err = e
            except Exception as e:
                last_err = e
        scope.wait_for_timeout(400)
    raise RuntimeError("Could not click %r (tried: %s) -> %s"
                       % (desc or texts, texts, last_err))


# ---------------------------------------------------------------------------
# Excel reading: Product Id + Category Path
# ---------------------------------------------------------------------------

ID_KEYS = ("product id", "productID".lower(), "productid", "id",
           "product sku", "sku", "product code", "code")
CAT_KEYS = ("category path", "categorypath", "category", "cat path", "cat")


def read_excel_rows(path):
    """Return [(product_id, [category levels...]), ...] de-duplicated, in order."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    grid = []
    for row in ws.iter_rows():
        vals = ["" if c.value is None else str(c.value).strip() for c in row]
        grid.append(vals)
    wb.close()
    if not grid:
        raise RuntimeError("Excel sheet is empty")

    header_idx = id_col = cat_col = None
    for r, row in enumerate(grid[:20]):
        cols = [norm(v) for v in row]
        i = next((k for k, v in enumerate(cols) if v in ID_KEYS), None)
        c = next((k for k, v in enumerate(cols) if v in CAT_KEYS), None)
        if i is not None and c is not None:
            header_idx, id_col, cat_col = r, i, c
            break
    if header_idx is None:
        raise RuntimeError(
            "Could not find 'Product Id' and 'Category Path' columns. "
            "First rows: %s" % (grid[:3],))

    out, seen = [], set()
    for row in grid[header_idx + 1:]:
        pid = re.sub(r"\.0$", "", (row[id_col] if id_col < len(row) else "").strip())
        cat_raw = row[cat_col] if cat_col < len(row) else ""
        if not pid or not cat_raw:
            continue
        # keep the original text exactly as in the Excel (typed into the
        # site's search bar verbatim); only split on '>' and drop empties
        levels = [re.sub(r"\s+", " ", p).strip()
                  for p in cat_raw.split(">")]
        levels = [p for p in levels if p]
        if not levels:
            continue
        key = (pid, tuple(levels))
        if key in seen:
            continue
        seen.add(key)
        out.append((pid, levels))
    return out


def write_report(path, results):
    rep = os.path.join(os.path.dirname(path), "update_report.csv")
    with open(rep, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(["Product Id", "Result", "Detail"])
        w.writerows(results)
    return rep


# ---------------------------------------------------------------------------
# Step 1-2: open site, fill ID + password, submit, then OTP from terminal
# ---------------------------------------------------------------------------

def find_visible(page, selector):
    """First visible element matching `selector`, or None."""
    loc = page.locator(selector)
    for i in range(loc.count()):
        el = loc.nth(i)
        try:
            if el.is_visible():
                return el
        except Exception:
            continue
    return None


# username / ID input candidates, most specific first
USER_SEL = ("input[name*='user' i], input[name*='email' i], input[name*='name' i], "
            "input[id*='user' i], input[id*='email' i], input[id*='login' i], "
            "input[type='email'], input[placeholder*='user' i], "
            "input[placeholder*='email' i], input[placeholder*='id' i], "
            "input[type='text']")
OTP_SEL = ("input[autocomplete=one-time-code], input[inputmode=numeric], "
           "input[placeholder*='otp' i], input[placeholder*='code' i], "
           "input[aria-label*='otp' i]")

def sign_in(page):
    print("\n[1] Opening %s" % START_URL)
    page.goto(START_URL, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
    page.wait_for_timeout(1500)
    shot(page, "01_opened")

    # Session kept from an earlier run? (test env usually still asks for login)
    if "signin" not in (page.url or "").lower() and \
       not page.locator("input[type=password]").count():
        print("    Already signed in (session kept).")
        return

    # If a landing page shows just a Sign in button first, click it.
    if not find_visible(page, "input[type=password]"):
        print("    Clicking Sign in on the landing page...")
        try:
            click_text(page, ["Sign in", "Sign In", "SIGN IN"], timeout=8000,
                       desc="Sign in")
            page.wait_for_timeout(1200)
        except Exception:
            pass
        shot(page, "02_login_form")

    # Fill ID + password
    user = find_visible(page, USER_SEL)
    pwd = find_visible(page, "input[type=password]")
    if user is None or pwd is None:
        shot(page, "02_login_form_missing")
        raise RuntimeError("Login form not found (ID or password input missing) "
                           "- check screenshots/02_login_form_missing.png")
    print("    Signing in as %s ..." % USERNAME)
    user.fill(USERNAME)
    pwd.fill(PASSWORD)
    shot(page, "03_credentials_filled")

    # Submit
    try:
        page.locator("button[type=submit], input[type=submit]")\
            .first.click(timeout=6000)
    except Exception:
        click_text(page, ["Sign in", "Sign In", "SIGN IN", "Log in", "Login"],
                   timeout=6000, desc="Sign in submit")
    page.wait_for_timeout(2000)
    shot(page, "04_after_submit")

    # OTP or dashboard?
    try:
        page.wait_for_selector(OTP_SEL, timeout=20_000, state="visible")
        code = ask("Enter OTP in terminal: ")
        while not code:
            code = ask("  OTP cannot be empty. Enter OTP in terminal: ")
        boxes = page.locator(OTP_SEL)
        n = boxes.count()
        print("    Entering OTP (%d input box(es))..." % n)
        if n > 1:
            for i, ch in enumerate(code[:n]):
                boxes.nth(i).fill(ch)
        else:
            boxes.first.fill(code)
        page.wait_for_timeout(800)
        shot(page, "05_otp_filled")
        # Click Verify (fallback: Submit / Confirm, then Enter as last resort)
        clicked = False
        for texts in (["Verify"], ["Submit", "Confirm", "Continue"]):
            try:
                click_text(page, texts, timeout=5000, desc="Verify button")
                clicked = True
                break
            except Exception:
                continue
        if not clicked:
            boxes.first.press("Enter")
        print("    Clicked Verify - waiting for the seller panel to load...")
        page.wait_for_timeout(3000)
        shot(page, "06_after_verify")
    except PWTimeout:
        shot(page, "06_no_otp_direct")
        print("    No OTP asked - continuing (already signed in).")


# ---------------------------------------------------------------------------
# Step 3-5: Seller Information -> Seller List -> search shop -> switch
# ---------------------------------------------------------------------------

def goto_seller_list(page):
    print("\n[2] Opening Seller Information -> Seller List")
    for texts in (LBL_SELLER_INFO, LBL_SELLER_LIST):
        click_text(page, texts, desc=texts[0])
        page.wait_for_timeout(900)
    shot(page, "04_seller_list")


def search_seller(page, code):
    """Search the Seller Code in the 2nd search bar, trying
    Pending -> Approve -> Reject tabs. Returns (row_locator, tab_name).
    """
    print("\n[3] Searching Seller Code %r (2nd search bar)" % code)
    boxes = page.locator("input[type=text], input[type=search], "
                         "input:not([type])")
    vis = []
    for i in range(min(boxes.count(), 25)):
        el = boxes.nth(i)
        try:
            if el.is_visible():
                vis.append(el)
        except Exception:
            continue
    if len(vis) < 2:
        shot(page, "07_search_bars")
        raise RuntimeError("Expected at least 2 search bars on Seller List "
                           "(Shop Name + Seller Code); found %d" % len(vis))
    box = vis[1]                     # 2nd visible search bar = Seller Code

    for tab_label, _key in TABS_SHOP:
        try:
            click_text(page, [tab_label], exact=True, timeout=5000,
                       desc="tab %s" % tab_label)
            page.wait_for_timeout(900)
        except Exception:
            pass  # tab may not exist / already active

        box.click()
        box.fill("")
        box.type(code, delay=60)
        box.press("Enter")
        page.wait_for_timeout(1800)
        shot(page, "08_search_%s" % tab_label.lower())

        row = find_seller_row(page, code)
        if row:
            print("    Found on tab: %s" % tab_label)
            return row, tab_label
        print("    Not on tab: %s" % tab_label)
    return None, None


def find_seller_row(page, code):
    """Locate the table row whose cells contain the seller code."""
    target = norm(code)
    rows = page.locator("table tbody tr")
    for i in range(rows.count()):
        cells = rows.nth(i).locator("td")
        for c in range(min(cells.count(), 8)):
            if target and target in norm(cells.nth(c).inner_text()):
                return rows.nth(i)
    return None


def switch_to_seller(page, row):
    print("    Clicking Switch to Seller icon -> Yes")
    # Icon-only button:
    #   <svg class="custom-target-icon" data-pr-tooltip="Switch to seller">
    candidates = [
        row.locator("svg[data-pr-tooltip*='switch' i]"),
        row.locator("[data-pr-tooltip*='switch' i]"),
        row.locator("svg.custom-target-icon"),
        row.locator("[class*='custom-target-icon']"),
        row.locator("[title*='switch' i], [aria-label*='switch' i]"),
    ]
    clicked = False
    for cand in candidates:
        try:
            el = cand.first
            if el.count() and el.is_visible():
                try:
                    el.click(timeout=8000)
                except Exception:
                    el.click(timeout=4000, force=True)   # icon may be overlaid
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        # text fallback (in case the site later adds a label)
        try:
            row.get_by_text(re.compile("switch", re.I)).first.click(timeout=6000)
            clicked = True
        except Exception:
            pass
    if not clicked:
        shot(page, "09_switch_icon_missing")
        raise RuntimeError("Switch to Seller icon not found in the seller row "
                           "- check screenshots/09_switch_icon_missing.png")
    page.wait_for_timeout(1200)
    shot(page, "07_confirm_dialog")

    # Confirmation: styled modal button or native confirm()
    clicked = False
    for texts in (["Yes"], ["Confirm", "OK", "Proceed"]):
        try:
            click_text(page, texts, exact=True, timeout=4000, desc="Yes/Confirm")
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        page.on("dialog", lambda d: d.accept())
        page.wait_for_timeout(800)
    page.wait_for_timeout(3000)
    shot(page, "08_seller_page")


# ---------------------------------------------------------------------------
# Step 6: Product -> Products -> Offline -> Rejected
# ---------------------------------------------------------------------------

def goto_rejected_products(page):
    print("\n[4] Opening Product -> Products -> Offline -> Rejected")
    click_text(page, ["Product", "Products"], desc="Product menu")
    page.wait_for_timeout(900)
    click_text(page, ["Products"], desc="Products submenu")
    page.wait_for_timeout(1200)
    for t in TABS_OFFLINE:
        click_text(page, [t], desc=t)
        page.wait_for_timeout(1200)
    for t in TABS_REJECTED:
        click_text(page, [t], desc=t)
        page.wait_for_timeout(1500)
    shot(page, "09_rejected_list")


# ---------------------------------------------------------------------------
# Steps 8-14: per-product Edit -> category cascade -> Main Materials -> Update
# ---------------------------------------------------------------------------

def set_rows_per_page(page, value="50"):
    """Open the paginator's page-size dropdown (bottom-right, shows the
    current size e.g. '10', options 10/20/50) and pick `value`, so the Rejected
    list shows 50 rows per page before searching for a product id."""
    print("    Setting page size to %s rows/page..." % value)
    opened = False
    # PrimeVue paginator select (p-select shows the current page size)
    sel = page.locator(".p-paginator .p-select, .p-paginator .p-dropdown, "
                       ".p-paginator [class*='select' i]")
    for i in range(min(sel.count(), 8)):
        el = sel.nth(i)
        try:
            if el.is_visible():
                el.click(timeout=5000)
                opened = True
                break
        except Exception:
            continue
    if not opened:
        # Fallback: the paginator dropdown label is an exact number ('10')
        # sitting BELOW the table (the tabs' counts are above it).
        table_y = 0
        try:
            bb = page.locator("table tbody tr").first.bounding_box()
            table_y = bb["y"] if bb else 0
        except Exception:
            pass
        cand = page.get_by_text(re.compile(r"^\d+$"), exact=True)
        for i in range(min(cand.count(), 40)):
            el = cand.nth(i)
            try:
                if not el.is_visible():
                    continue
                bb = el.bounding_box()
                if not bb or bb["y"] < table_y:
                    continue                 # must sit at/below the table
                if bb["width"] > 200 or bb["height"] > 60:
                    continue                 # not a small dropdown label
                el.click(timeout=4000)
                opened = True
                break
            except Exception:
                continue
    if not opened:
        shot(page, "10_page_size_dropdown_missing")
        print("    !! Page-size dropdown not found - searching 10/page.")
        return False
    page.wait_for_timeout(700)
    # Pick the value from the option list (10 / 20 / 50)
    try:
        click_text(page, [value], exact=True, timeout=5000,
                   desc="page size %s" % value)
    except Exception:
        shot(page, "10_page_size_option_missing")
        print("    !! Option %r not shown in the dropdown." % value)
        page.keyboard.press("Escape")
        return False
    page.wait_for_timeout(1800)               # list reloads with 50 rows
    shot(page, "10_page_size_%s" % value)
    return True


def goto_next_page(page):
    """Click the paginator's next button; return False on the last page or
    if no next button is visible. Handles the chevron-style paginator
    ('<  1  2  3  4  5  >  >>') and detects its disabled/last state."""
    candidates = [
        page.locator("li.p-paginator-next:not(.p-disabled), button.p-paginator-next"),
        page.locator("a[aria-label*='next' i], button[aria-label*='next' i]"),
        page.locator("ul.pagination li.next:not(.disabled) a"),
        page.locator("nav[aria-label*=pagination i] li:last-child a"),
    ]
    for cand in candidates:
        for i in range(min(cand.count(), 6)):
            el = cand.nth(i)
            try:
                if not el.is_visible():
                    continue
                attrs = " ".join((el.get_attribute("class") or "",
                                  el.get_attribute("aria-disabled") or ""))
                if re.search(r"p-disabled|disabled|\btrue\b", attrs, re.I):
                    return False             # last page reached
                el.click(timeout=8000)
                page.wait_for_timeout(1800)  # rows reload
                return True
            except Exception:
                continue
    return False


def wait_edit_page_ready(page, timeout=45_000):
    """After clicking Edit the product edit page loads slowly (a spinner is
    shown first). Wait until the category field (an input whose value holds
    the current path, e.g. '... > ...') actually renders before continuing."""
    print("    Waiting for the edit page to load (category field)...")
    deadline = time.time() + timeout / 1000
    last_err = None
    while time.time() < deadline:
        try:
            inputs = page.locator("input")
            for i in range(min(inputs.count(), 40)):
                el = inputs.nth(i)
                try:
                    if el.is_visible() and ">" in (el.input_value() or ""):
                        print("    Edit page ready.")
                        return
                except Exception as e:
                    last_err = e
        except Exception as e:
            last_err = e
        page.wait_for_timeout(800)
    shot(page, "11_edit_page_timeout")
    raise RuntimeError("Edit page did not load within %ds (category field "
                       "never appeared) - check screenshots/"
                       "11_edit_page_timeout.png" % (timeout // 1000))


def open_product_edit(page, pid):
    """Find the product row by id (following pagination) and click Edit.

    The Edit control is an icon-only PrimeVue button (pencil icon, empty label):
      <button class="p-button ... p-button-icon-only p-button-text p-button-success">
        <svg><path d="m7 17.013 ..."/></svg>
    """
    print("\n[5] Product %s: locating row..." % pid)
    set_rows_per_page(page, "50")            # 50 rows/page, then walk pages
    # search to the LAST page
    for page_no in range(1, 41):
        rows = page.locator("table tbody tr")
        for i in range(rows.count()):
            row = rows.nth(i)
            if pid.lower() in norm(row.inner_text()):
                shot(page, "10_row_%s" % pid)
                # Icon-only edit button: match by its pencil-icon path data,
                # or by the green icon-only PrimeVue button class as fallback.
                pencil = ("m7 17.013 4.413-.015 9.632-9.54c.378-.378.586-.88."
                          "586-1.414s-.208-1.036-.586-1.414l-1.586-1.586")
                candidates = [
                    row.locator("button:has(svg path[d^='%s'])" % pencil),
                    row.locator("button.p-button-success"),
                    row.locator("button.p-button-icon-only"),
                    row.locator("button:has(svg[class*='fa-pencil' i]), "
                                "button:has(svg[class*='edit' i])"),
                    row.get_by_text(re.compile("^edit$", re.I)),
                ]
                clicked = False
                for cand in candidates:
                    try:
                        el = cand.first
                        if el.count() and el.is_visible():
                            try:
                                el.click(timeout=8000)
                            except Exception:
                                el.click(timeout=4000, force=True)
                            clicked = True
                            break
                    except Exception:
                        continue
                if not clicked:
                    shot(page, "11_edit_btn_%s" % pid)
                    raise RuntimeError("Edit button not found in row for %s "
                                       "- check screenshots/11_edit_btn_%s.png"
                                       % (pid, pid))
                wait_edit_page_ready(page)
                shot(page, "11_edit_%s" % pid)
                return True
        # try next page (stops automatically on the last page)
        if not goto_next_page(page):
            break
    shot(page, "10_row_%s_notfound" % pid)
    return False


def find_search_bars(page):
    """All visible Search inputs in the category popup, ordered LEFT to RIGHT.
    Returns [(x, y, locator), ...]. Each cascade column has its own search
    bar - a path can be up to 5 LEVELS deep, so up to 5 bars can exist and
    the popup then scrolls horizontally. The bar to type into is always the
    RIGHTMOST one (the newest column). Bars keep stale text from earlier
    picks, so callers must fill('') before typing."""
    seen, out = set(), []
    sels = ["input[placeholder='Search']", "input[placeholder='search']",
            "input.p-inputtext[placeholder*='search' i]"]
    for sel in sels:
        loc = page.locator(sel)
        for i in range(loc.count()):
            el = loc.nth(i)
            try:
                if not el.is_visible():
                    continue
                box = el.bounding_box()
                if not box:
                    continue
                key = (round(box["x"]), round(box["y"]))
                if key in seen:
                    continue
                seen.add(key)
                out.append((box["x"], box["y"], el))
            except Exception:
                continue
    out.sort(key=lambda t: (t[0], t[1]))
    return out


def find_search_bar(page):
    """Any visible Search input (used just to verify the popup opened)."""
    bars = find_search_bars(page)
    return bars[-1][2] if bars else None


def scroll_popup_right(page, anchor):
    """The cascade popup scrolls horizontally (a 5-level path shows 5
    columns but only ~3 fit on screen). Scroll the popup's scrollable
    ancestor to the far right so the newest column is fully in view."""
    try:
        handle = anchor.evaluate_handle(
            "node => { let n = node; while (n && n !== document.body) {"
            " const cs = getComputedStyle(n);"
            " if ((cs.overflowX === 'auto' || cs.overflowX === 'scroll')"
            "     && n.scrollWidth > n.clientWidth + 10) return n;"
            " n = n.parentElement; } return null; }")
        sc = handle.as_element()
        if sc is not None:
            sc.evaluate("node => node.scrollTo({left: node.scrollWidth,"
                        " behavior: 'instant'})")
            page.wait_for_timeout(400)
    except Exception:
        pass


def newest_search_bar(page, min_x=None, timeout_s=8.0):
    """The newest cascade column is always the RIGHTMOST Search bar, no
    matter how many levels (1-5) are already on screen. Wait up to timeout_s
    for a bar further right than min_x (the previous level's bar) to slide
    in; return (x, y, locator), falling back to the current rightmost bar."""
    deadline = time.time() + timeout_s
    while True:
        bars = find_search_bars(page)
        if bars:
            bx, by, bel = bars[-1]
            if min_x is None or bx > min_x + 40:
                return bx, by, bel
        if time.time() > deadline:
            return bars[-1] if bars else None
        page.wait_for_timeout(600)


def open_category_picker(page):
    """Click the filled Category input (value holds the current path with '>')
    so the popup with its Search bar appears."""
    print("    Opening Category picker...")
    field = None
    cand = page.locator("input.p-inputtext.p-filled")
    for i in range(cand.count()):
        el = cand.nth(i)
        try:
            if not el.is_visible():
                continue
            if ">" in (el.input_value() or ""):
                field = el
                break
        except Exception:
            continue
    if field is None:                      # any visible input holding a path
        allinp = page.locator("input")
        for i in range(allinp.count()):
            el = allinp.nth(i)
            try:
                if el.is_visible() and ">" in (el.input_value() or ""):
                    field = el
                    break
            except Exception:
                continue
    if field is None:
        shot(page, "12_category_field_missing")
        raise RuntimeError("Category input (value contains '>') not found - "
                           "check screenshots/12_category_field_missing.png")
    field.click(timeout=8000)
    page.wait_for_timeout(1000)
    if find_search_bar(page) is None:
        shot(page, "12_search_bar_missing")
        raise RuntimeError("Search bar did not appear after clicking the "
                           "Category field - check screenshots/"
                           "12_search_bar_missing.png")


def find_and_click_option(page, bar, centers, level, attempts=3):
    """Click the dropdown option equal to `level`, searching by TEXT anywhere
    on screen (the option row may be a plain div, not an li/[role=option]).

    The option must belong to THIS level's column: below its search bar and
    CLOSER to that bar's x than to any other column's bar (columns sit
    ~290px apart, so a fixed +/- window would leak into the neighbour once
    4-5 columns are on screen). Smallest visible match wins so we click the
    row itself, not a page-size wrapper."""
    for attempt in range(attempts):             # retry while dropdown filters
        sbox = bar.bounding_box() or {}
        sx, sy = sbox.get("x", 0), sbox.get("y", 0)
        others = [x for x in centers if abs(x - sx) > 40]
        for pass_no in (1, 2):                  # 1: exact, 2: contains
            try:
                matches = (page.get_by_text(level, exact=True)
                           if pass_no == 1 else
                           page.get_by_text(re.compile(re.escape(level), re.I)))
                best, best_area = None, None
                for i in range(min(matches.count(), 40)):
                    el = matches.nth(i)
                    try:
                        if not el.is_visible():
                            continue
                        box = el.bounding_box()
                        if not box or box["width"] > 900 or \
                           box["height"] > 220:
                            continue            # wrapper/page containers
                        cx, cy = box["x"], box["y"]
                        if cy < sy - 5:
                            continue            # options sit below the bar
                        dx = abs(cx - sx)
                        if dx > 145:            # ~half the column pitch
                            continue
                        if any(abs(cx - ox) < dx for ox in others):
                            continue            # belongs to another column
                        area = box["width"] * box["height"]
                        if best_area is None or area < best_area:
                            best, best_area = el, area
                    except Exception:
                        continue
                if best is not None:
                    try:
                        best.scroll_into_view_if_needed(timeout=3000)
                    except Exception:
                        pass
                    try:
                        best.click(timeout=4000)
                    except Exception:
                        best.click(timeout=4000, force=True)
                    return True
            except Exception:
                continue
        page.wait_for_timeout(900)              # give the filter a moment
    return False


def pick_category(page, levels):
    """Type each category level into the popup Search bar and pick it from
    the dropdown; each pick slides in the next list with a fresh Search bar.
    Finishes by confirming the final 'Yes, Continue' popup."""
    print("    Category: %s" % " > ".join(levels))
    open_category_picker(page)

    prev_x = None                       # x of the bar used by the last level
    for depth, level in enumerate(levels):
        # A path can be up to 5 levels deep -> up to 5 columns, with a
        # horizontal scrollbar. The newest column is always the RIGHTMOST
        # Search bar: wait for it to slide in right of the previous level's
        # bar, then scroll the popup right so it is fully in view.
        got = newest_search_bar(page, min_x=prev_x, timeout_s=8.0)
        if not got:
            shot(page, "12_search_bar_missing_d%d" % depth)
            raise RuntimeError("Search bar for category level %d not found "
                               "(no Search bar visible in the popup)"
                               % (depth + 1))
        bx, by, search = got
        try:
            search.scroll_into_view_if_needed()
        except Exception:
            pass
        scroll_popup_right(page, search)
        centers = [b[0] for b in find_search_bars(page)]
        sbox = search.bounding_box() or {"x": bx, "y": by}
        sx, sy = sbox["x"], sbox["y"]

        # Short queries: the dropdown filter often fails on long phrases, so
        # for EVERY level try the first word first, then the first two words,
        # then the full text - clicking the option as soon as it shows up
        # (e.g. type 'Outdoor' or 'Outdoor Gardening' and click
        # 'Outdoor Gardening Power Tools' from the filtered list).
        words = level.replace(",", " ").split()
        tries = [words[0]] if len(words) > 1 else []
        if len(words) > 2:
            tries.append(" ".join(words[:2]))
        tries.append(level)
        seen_q, ladder = set(), []
        for q in tries:
            if q not in seen_q:
                seen_q.add(q)
                ladder.append(q)

        clicked = False
        for typed in ladder:
            print("    Level %d/%d: bar at x=%.0f, typing %r"
                  % (depth + 1, len(levels), sx, typed))
            search.click()
            search.fill("")                      # clear stale text
            search.type(typed, delay=80)
            page.wait_for_timeout(1200)         # dropdown filters as you type
            if norm(search.input_value() or "") != norm(typed):
                search.fill("")                  # retry once if it did not stick
                search.type(typed, delay=80)
                page.wait_for_timeout(1200)
            shot(page, "12_cat_typed_d%d" % depth)
            clicked = find_and_click_option(page, search, centers, level)
            if clicked:
                break
            print("    Option %r not shown with %r - trying a longer query..."
                  % (level, typed))
        if not clicked:
            shot(page, "12_cat_fail_d%d" % depth)
            raise RuntimeError("Category option %r not visible/clickable in "
                               "dropdown" % level)
        page.wait_for_timeout(1200)             # next list slides in
        prev_x = sx                             # next column is right of this
        shot(page, "12_cat_d%d" % depth)

    # Final confirmation popup
    page.wait_for_timeout(1000)
    shot(page, "12_cat_confirm_popup")
    clicked = False
    for texts in (["Yes, Continue"], ["Yes Continue"], ["Yes"], ["Continue"]):
        try:
            click_text(page, texts, timeout=4000, desc="Yes, Continue")
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        shot(page, "12_confirm_missing")
        raise RuntimeError("Confirmation popup (Yes, Continue) not found")
    page.wait_for_timeout(1500)
    shot(page, "12_cat_done")


def fill_text_under_label(page, label_rx, value):
    """Find the visible label matching label_rx and type `value` into the
    nearest text input BELOW it (Main Material / Warranty Type are plain
    text bars on this panel, not dropdowns). Returns the input or None."""
    labels = page.get_by_text(re.compile(label_rx, re.I))
    label_boxes = []
    for i in range(min(labels.count(), 10)):
        el = labels.nth(i)
        try:
            if el.is_visible():
                lb = el.bounding_box()
                if lb:
                    label_boxes.append(lb)
        except Exception:
            continue
    if not label_boxes:
        return None
    inputs = []
    inp_loc = page.locator("input[type='text'], input:not([type]), "
                           "input.p-inputtext")
    for i in range(min(inp_loc.count(), 80)):
        el = inp_loc.nth(i)
        try:
            if not el.is_visible():
                continue
            bb = el.bounding_box()
            if bb and bb["height"] < 120:
                inputs.append((bb, el))
        except Exception:
            continue
    for lb in label_boxes:
        lx, ly = lb["x"] + lb["width"] / 2, lb["y"]
        best, best_d = None, None
        for bb, el in inputs:
            iy = bb["y"] + bb["height"] / 2
            ix = bb["x"] + bb["width"] / 2
            if iy < ly - 5:                     # must be at/below the label
                continue
            dy, dx = iy - ly, abs(ix - lx)
            if dy > 220 or dx > 520:
                continue
            d = dy + dx * 0.3
            if best_d is None or d < best_d:
                best, best_d = el, d
        if best is None:
            continue
        try:
            best.click(timeout=5000)
            best.fill("")
            best.type(value, delay=60)
            page.wait_for_timeout(600)
            if (best.input_value() or "").strip():
                return best
        except Exception:
            continue
    return None


def set_main_materials(page):
    """Main Materials* is a text bar under its label: click the bar and type
    'Other'. Returns True when the bar was filled, None if not shown."""
    try:
        mm = page.get_by_text(re.compile(MAIN_MATERIALS, re.I)).first
        mm.wait_for(state="visible", timeout=2500)
    except Exception:
        return None
    print("    Main Materials* found -> typing %r in its text bar"
          % MATERIALS_CHOICE)
    bar = fill_text_under_label(page, MAIN_MATERIALS, MATERIALS_CHOICE)
    if bar is None:
        shot(page, "14_main_material_bar_missing")
        return None
    # autocomplete suggestion list? click an exact 'Other' option that popped
    # up right below the bar (harmless if the field is a plain text input).
    try:
        bb = bar.bounding_box()
        sug = page.get_by_text(MATERIALS_CHOICE, exact=True)
        for i in range(min(sug.count(), 8)):
            el = sug.nth(i)
            sb = el.bounding_box()
            if not sb or sb["width"] > 420 or sb["height"] > 90:
                continue
            if bb and abs(sb["x"] - bb["x"]) < 320 and \
               0 <= sb["y"] - bb["y"] <= 150:
                el.click(timeout=3000)
                break
    except Exception:
        pass
    page.wait_for_timeout(400)
    return True


def set_warranty_type(page):
    """Warranty Type: click its text bar and type 'No Warranty'."""
    try:
        wt = page.get_by_text(re.compile(WARRANTY_TYPE, re.I)).first
        wt.wait_for(state="visible", timeout=2500)
    except Exception:
        return None
    print("    Warranty Type found -> typing %r" % WARRANTY_VALUE)
    bar = fill_text_under_label(page, WARRANTY_TYPE, WARRANTY_VALUE)
    if bar is None:
        shot(page, "14_warranty_bar_missing")
        return None
    return True


def visible_notification_text(page):
    """Combined text of visible toast / dialog / validation messages."""
    sels = (".p-toast-summary", ".p-toast-detail", ".p-toast-message",
            "[class*='toast' i] [class*='summary' i]",
            "[class*='toast' i] [class*='detail' i]",
            ".p-dialog-content", ".p-message-text", ".p-inline-message",
            ".p-error", "small[class*='error' i]",
            "[class*='validation' i]", "[class*='error' i] small")
    texts, seen = [], set()
    for sel in sels:
        try:
            loc = page.locator(sel)
        except Exception:
            continue
        for i in range(min(loc.count(), 10)):
            el = loc.nth(i)
            try:
                if el.is_visible():
                    t = norm(el.inner_text())
                    if t and t not in seen:
                        seen.add(t)
                        texts.append(t)
            except Exception:
                continue
    return " | ".join(texts)


def click_update(page):
    """Scroll to the bottom and click Update. Returns 'success', 'validation'
    (a Main Material / Warranty notification appeared) or 'unknown'."""
    print("    Clicking Update...")
    page.keyboard.press("End")
    page.wait_for_timeout(800)
    btn = page.get_by_role("button", name=re.compile("update", re.I))
    if not btn.count():
        btn = page.get_by_text(re.compile("^update$", re.I))
    btn.last.scroll_into_view_if_needed()
    btn.last.click(timeout=10_000)
    outcome = "unknown"
    deadline = time.time() + TOAST_WAIT / 1000
    while time.time() < deadline:
        notif = visible_notification_text(page)
        if re.search(r"main material|warranty", notif, re.I):
            print("    Validation notification: %s" % notif[:160])
            outcome = "validation"
            break
        if re.search(r"success|updated|done", notif, re.I):
            print("    Update confirmed by on-screen message.")
            outcome = "success"
            break
        page.wait_for_timeout(500)
    page.wait_for_timeout(1200)
    shot(page, "13_after_update")
    return outcome


def handle_post_update_notifications(page, max_rounds=3):
    """After Update the panel may complain that Main Material or Warranty
    Type/Period is missing. Fill the requested field ('Other' / 'No Warranty'
    in Warranty Type) and click Update again, up to max_rounds times."""
    for round_no in range(1, max_rounds + 1):
        notif = visible_notification_text(page)
        if re.search(r"main material", notif, re.I):
            print("    Notification: Main Material required.")
            filled = set_main_materials(page)
        elif re.search(r"warranty", notif, re.I):
            print("    Notification: Warranty Type/Period missing.")
            filled = set_warranty_type(page)
        else:
            return                                  # nothing asked for
        if not filled:
            shot(page, "14_fill_failed_r%d" % round_no)
            print("    !! Could not fill the requested field.")
            return
        shot(page, "14_validation_filled_r%d" % round_no)
        click_update(page)
        page.wait_for_timeout(800)


def back_to_rejected_list(page):
    """Return to the Rejected list for the next product - WITHOUT re-clicking
    Product / Products menus when we are already on the products page.

    Real page layout (from screenshots):
      top tabs : Online (176) | Pending (5) | Offline (88)
      sub-tabs : In Active (0) | Rejected (89) | Restricted (0)
    After Update the panel lands on the Online tab, so we must click
    'Offline (N)' then 'Rejected (N)'. Tab labels carry live counts, so
    matching must allow '(digits)' after the word.
    """
    rej_rx = re.compile(r"rejected\s*\(\s*\d+\s*\)", re.I)
    off_rx = re.compile(r"offline\s*\(\s*\d+\s*\)", re.I)

    def visible(rx):
        loc = page.get_by_text(rx)
        for i in range(min(loc.count(), 10)):
            try:
                if loc.nth(i).is_visible():
                    return loc.nth(i)
            except Exception:
                continue
        return None

    def wait_rows():
        try:
            page.wait_for_selector("table tbody tr", timeout=10_000,
                                   state="visible")
        except PWTimeout:
            pass
        page.wait_for_timeout(800)

    # 1. Already showing the Rejected sub-tab (active or not)?
    rej = visible(rej_rx)
    if rej is not None:
        try:
            cls = (rej.get_attribute("class") or "")
        except Exception:
            cls = ""
        if "active" in cls.lower():            # already the active tab
            print("    Already on the Rejected list - going to next row.")
            wait_rows()
            return
        rej.click(timeout=5000)                # visible but not active
        page.wait_for_timeout(1500)
        print("    Clicked the Rejected sub-tab.")
        wait_rows()
        return

    # 2. Products page on another top tab (e.g. Online after Update):
    #    click Offline (N), then Rejected (N).
    off = visible(off_rx)
    if off is not None:
        off.click(timeout=5000)
        page.wait_for_timeout(1500)
        rej = visible(rej_rx)
        if rej is not None:
            rej.click(timeout=5000)
            page.wait_for_timeout(1500)
            print("    Clicked Offline -> Rejected tabs in place.")
            wait_rows()
            return

    # 3. Truly elsewhere -> full sidebar navigation
    print("    Navigating Product -> Products -> Offline -> Rejected ...")
    goto_rejected_products(page)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 64)
    print("CartUp Category Update Agent")
    print("Folder: %s" % AGENT_DIR)
    print("=" * 64)

    code = ask("Enter Seller Code: ")
    while not code:
        code = ask("  Seller Code cannot be empty. Enter Seller Code: ")

    xlsx = ask("Path to Excel file (Product Id + Category Path): ").strip('"')
    while not os.path.exists(xlsx):
        xlsx = ask("  File not found. Path to Excel file: ").strip('"')
    rows = read_excel_rows(xlsx)
    print("Loaded %d product row(s) from %s\n" % (len(rows), os.path.basename(xlsx)))
    for pid, levels in rows:
        print("  %-12s -> %s" % (pid, " > ".join(levels)))
    print()

    results, done = [], set()
    os.makedirs(SHOT_DIR, exist_ok=True)

    with sync_playwright() as pw:
        browser = pw.chromium.launch_persistent_context(
            PROFILE_DIR, headless=HEADLESS, slow_mo=SLOW_MO,
            viewport={"width": 1440, "height": 900},
            args=["--start-maximized"])
        page = browser.pages[0] if browser.pages else browser.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)

        try:
            sign_in(page)
            goto_seller_list(page)

            seller_row, tab = search_seller(page, code)
            if not seller_row:
                shot(page, "08_seller_not_found")
                print("\n!! Seller Code %r was not found on Pending/Approve/Reject."
                      % code)
                print("   Check the code, then re-run.")
                return 1
            switch_to_seller(page, seller_row)
            goto_rejected_products(page)

            for idx, (pid, levels) in enumerate(rows, 1):
                if pid in done:
                    continue
                print("\n---- Product %d/%d: %s ----" % (idx, len(rows), pid))
                try:
                    back_to_rejected_list(page)
                    if not open_product_edit(page, pid):
                        print("    !! Row not found in Rejected list - skipped.")
                        results.append([pid, "NOT FOUND", "row missing in list"])
                        continue
                    pick_category(page, levels)
                    set_main_materials(page)   # fill the bar pre-emptively if shown
                    click_update(page)
                    handle_post_update_notifications(page)
                    results.append([pid, "OK", " > ".join(levels)])
                    done.add(pid)
                except Exception as e:
                    shot(page, "99_error_%s" % pid)
                    print("    !! %s" % e)
                    results.append([pid, "FAILED", str(e)[:180]])
                    if ask("    Continue with next product? [Y/n]: ").lower() == "n":
                        break

            rep = write_report(xlsx, results)
            ok = sum(1 for r in results if r[1] == "OK")
            print("\n" + "=" * 64)
            print("Done. %d/%d updated OK." % (ok, len(results)))
            for r in results:
                print("  %-12s %-9s %s" % (r[0], r[1], r[2]))
            print("Report: %s" % rep)
            return 0
        except Exception:
            traceback.print_exc()
            shot(page, "99_fatal")
            return 1
        finally:
            ask("\nPress Enter to close the browser...")
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
