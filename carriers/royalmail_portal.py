"""
Royal Mail OBA portal automation.

Handles the Royal Mail Online Business Account portal workflow:
1. Auto-launch Edge with remote debugging enabled
2. Navigate Edge to the OBA login page
3. Wait for user to log in and reach the OBA dashboard
4. Automate: select posting location, create order, fill form, confirm
5. Save confirmation as PDF and optionally print
6. Log out of OBA and close the automation browser

One order covers every destination in the posting: each destination and
format becomes its own line on the order form.

Akamai bot protection on royalmail.com blocks automated Chromium,
so we use the user's real Edge browser via CDP (Chrome DevTools Protocol).
Edge is launched automatically — the user only needs to log in.
"""

import os
import re
import asyncio
import subprocess
import time
from datetime import datetime
from typing import Callable, List, Optional
from dataclasses import dataclass, field

from .royalmail_countries import (
    SUPPORTED_COUNTRIES,
    CountryLine,
    normalise_text,
)
from core.credentials import get_royalmail_credentials


# Portal constants
OBA_LOGIN_URL = "https://www.royalmail.com/discounts-payment/credit-account/online-business-account"
CDP_PORT = 9222
CDP_URL = f"http://127.0.0.1:{CDP_PORT}"

# Product codes
PRODUCT_CODE_LETTERS = "PS5"
PRODUCT_CODE_FLATS = "PS7"

# Which OBA product each format is filed under
PRODUCT_CODES = {
    'Letters': PRODUCT_CODE_LETTERS,
    'Flats': PRODUCT_CODE_FLATS,
}

# Edge 136+ ignores --remote-debugging-port on the default user profile
# (Chromium security change), so CDP must use a dedicated profile directory.
# Persistent so the OBA login session is cached between runs.
EDGE_PROFILE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "MultiCarrierManifestTool", "edge-oba-profile"
)

# Edge executable paths (checked in order)
EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


def _leads_with(text: str, target: str) -> bool:
    """True if text starts with target on a word boundary."""
    if not text.startswith(target):
        return False
    rest = text[len(target):]
    return not rest or not rest[0].isalnum()


def _contains_word(text: str, target: str) -> bool:
    """True if target appears in text as a whole word."""
    return re.search(rf'(?<![a-z0-9]){re.escape(target)}(?![a-z0-9])', text) is not None


def match_option_value(options, wanted: str, what: str) -> str:
    """
    Pick the <option> value whose label matches `wanted`.

    We cannot verify OBA's exact dropdown wording from outside the portal —
    Ireland reads 'IRELAND (REPUBLIC OF)', and the others could carry
    suffixes of their own — so match against whatever the portal actually
    rendered instead of hardcoding a guess. Tried in order: an exact label,
    a label leading with the name, then a label containing it as a word.

    Raises LookupError naming the options that were available, so a wrong
    guess shows up as a clear failure rather than a wrong country on a
    confirmed order.
    """
    target = normalise_text(wanted)
    candidates = [(value, normalise_text(text), text) for value, text in options]
    candidates = [c for c in candidates if c[1]]

    for tier_name, predicate in (
        ('exact', lambda t: t == target),
        ('leading', lambda t: _leads_with(t, target)),
        ('containing', lambda t: _contains_word(t, target)),
    ):
        hits = [c for c in candidates if predicate(c[1])]
        if len(hits) == 1:
            return hits[0][0]
        if len(hits) > 1:
            labels = ', '.join(repr(h[2]) for h in hits)
            raise LookupError(
                f"Ambiguous {what} '{wanted}' on the OBA form — "
                f"{len(hits)} options match ({tier_name}): {labels}"
            )

    available = ', '.join(repr(c[2]) for c in candidates) or '(none)'
    raise LookupError(
        f"No {what} option matching '{wanted}' on the OBA form. "
        f"Available: {available}"
    )


@dataclass
class RoyalMailPortalInput:
    """
    Combined input for one OBA order, from one or several carrier sheets.

    Each CountryLine becomes one line on the order form, so a single order
    covers every destination in the posting.
    """
    po_number: str
    lines: List[CountryLine] = field(default_factory=list)

    @property
    def ordered_lines(self) -> List[CountryLine]:
        """
        Non-empty lines in the order they should be entered on the form:
        registry country order, Letters before Flats within each country.
        """
        country_order = list(SUPPORTED_COUNTRIES)
        format_order = ['Letters', 'Flats']

        def sort_key(line: CountryLine):
            country_rank = (
                country_order.index(line.country)
                if line.country in country_order else len(country_order)
            )
            format_rank = (
                format_order.index(line.format_type)
                if line.format_type in format_order else len(format_order)
            )
            return (country_rank, format_rank, line.country, line.format_type)

        return sorted((l for l in self.lines if l.items > 0), key=sort_key)

    @property
    def has_any(self) -> bool:
        return any(line.items > 0 for line in self.lines)

    def merge(self, other: 'RoyalMailPortalInput') -> None:
        """
        Fold another sheet's volumes into this order.

        Batch mode combines several carrier sheets into one OBA order, so a
        destination appearing on two sheets must sum rather than double-count.
        """
        by_key = {(l.country, l.format_type): l for l in self.lines}
        for incoming in other.lines:
            key = (incoming.country, incoming.format_type)
            existing = by_key.get(key)
            if existing:
                existing.items += incoming.items
                existing.weight_kg = round(existing.weight_kg + incoming.weight_kg, 3)
            else:
                new_line = CountryLine(
                    country=incoming.country,
                    format_type=incoming.format_type,
                    items=incoming.items,
                    weight_kg=incoming.weight_kg,
                )
                self.lines.append(new_line)
                by_key[key] = new_line

        if not self.po_number:
            self.po_number = other.po_number

    def describe(self) -> List[str]:
        """One human-readable summary line per destination, for the log."""
        return [
            f"{line.country} {line.format_type}: {line.items} items, "
            f"{line.avg_weight_grams}g avg ({line.weight_kg} kg total)"
            for line in self.ordered_lines
        ]


def _find_edge_executable() -> Optional[str]:
    """Find the Edge browser executable."""
    import shutil
    for path in EDGE_PATHS:
        if os.path.exists(path):
            return path
    # Fallback: try PATH
    found = shutil.which("msedge")
    return found


def _kill_edge_processes():
    """Kill all running Edge processes."""
    try:
        subprocess.run(
            ['taskkill', '/F', '/IM', 'msedge.exe'],
            capture_output=True, timeout=10
        )
        time.sleep(2)  # Wait for processes to fully exit
    except Exception:
        pass


def _find_profile_edge_pids() -> Optional[list]:
    """
    PIDs of Edge processes running under our dedicated OBA profile.

    Returns None if the process list could not be read, so callers can tell
    "no automation Edge running" apart from "could not look".
    """
    query = (
        "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | "
        f"Where-Object {{ $_.CommandLine -like '*{EDGE_PROFILE_DIR}*' }} | "
        "ForEach-Object { $_.ProcessId }"
    )
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-Command', query],
            capture_output=True, text=True, timeout=20
        )
        if result.returncode != 0:
            return None
        return [int(l.strip()) for l in result.stdout.splitlines() if l.strip().isdigit()]
    except Exception:
        return None


def _close_automation_edge(log=None) -> None:
    """
    Terminate the Edge instance this tool launched, leaving other Edge
    windows alone.

    Targets only processes running under EDGE_PROFILE_DIR so the user's own
    browsing survives. Falls back to the blunt all-Edge kill only if the
    process list cannot be read, and says so when it does.
    """
    def say(msg):
        if log:
            log(msg)

    pids = _find_profile_edge_pids()

    if pids is None:
        say("    Could not list Edge processes; closing all Edge windows")
        _kill_edge_processes()
        return

    if not pids:
        say("    Automation Edge already closed")
        return

    for pid in pids:
        try:
            subprocess.run(
                ['taskkill', '/F', '/T', '/PID', str(pid)],
                capture_output=True, timeout=10
            )
        except Exception:
            pass

    time.sleep(2)  # Let the processes actually exit before we report
    if _is_edge_cdp_available():
        say("    Edge still responding on the debug port after close")
    else:
        say(f"    Closed automation Edge ({len(pids)} process(es))")


async def _log_out_of_oba(browser, log) -> bool:
    """
    Click the OBA log-off control so the session does not stay open.

    The control lives in the portal masthead, which may be the top-level
    page or one of the SAP portal frames, so every frame is tried.
    """
    logout_pattern = re.compile(r'log\s?off|log\s?out|sign\s?out', re.IGNORECASE)

    for ctx in browser.contexts:
        for page in ctx.pages:
            try:
                if 'royalmail.com' not in page.url.lower():
                    continue
            except Exception:
                continue

            # Log off may raise a confirm dialog. Playwright auto-dismisses
            # dialogs when nothing is listening, which would cancel the very
            # thing we are trying to do, so accept them explicitly.
            page.on('dialog', lambda dialog: asyncio.ensure_future(dialog.accept()))

            for frame in list(page.frames):
                for role in ('link', 'button'):
                    try:
                        control = frame.get_by_role(role, name=logout_pattern).first
                        if await control.count() == 0:
                            continue
                        await control.click(timeout=5000)
                        log("    Logged out of OBA")
                        await page.wait_for_timeout(3000)
                        return True
                    except Exception:
                        continue

    log("    No log-off control found; closing the browser instead")
    return False


async def _logout_and_close_edge(log) -> None:
    """
    End the OBA session and close the automation browser.

    Runs on every path — success, failure or crash — so a run never leaves a
    logged-in tab behind. Reconnects over CDP rather than reusing the
    automation's connection, which keeps it independent of how the run ended.
    """
    log("  Closing Royal Mail session...")

    if _is_edge_cdp_available():
        try:
            from playwright.async_api import async_playwright
            async with async_playwright() as p:
                browser = await p.chromium.connect_over_cdp(CDP_URL)
                try:
                    await _log_out_of_oba(browser, log)
                finally:
                    try:
                        await browser.close()
                    except Exception:
                        pass
        except Exception as e:
            log(f"    Log-off skipped ({e})")

    _close_automation_edge(log)


def _is_edge_cdp_available() -> bool:
    """Check if Edge is already running with CDP on the expected port."""
    import urllib.request
    try:
        req = urllib.request.urlopen(f"http://127.0.0.1:{CDP_PORT}/json/version", timeout=2)
        req.close()
        return True
    except Exception:
        return False


def launch_edge_for_royalmail(log_callback=None) -> tuple[bool, str]:
    """
    Launch Edge with remote debugging and navigate to OBA login page.

    Returns:
        (success, message)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    # Check if Edge CDP is already available
    if _is_edge_cdp_available():
        log("  Edge already running with remote debugging")
        return True, "Edge already available"

    # Find Edge executable
    edge_path = _find_edge_executable()
    if not edge_path:
        return False, "Microsoft Edge not found. Please install Edge browser."

    # Kill existing Edge processes
    log("  Closing existing Edge windows...")
    _kill_edge_processes()

    # Launch Edge with remote debugging
    log("  Launching Edge with remote debugging...")
    try:
        os.makedirs(EDGE_PROFILE_DIR, exist_ok=True)
        subprocess.Popen(
            [
                edge_path,
                f"--remote-debugging-port={CDP_PORT}",
                f"--user-data-dir={EDGE_PROFILE_DIR}",
                "--no-first-run",
                "--no-default-browser-check",
                OBA_LOGIN_URL,
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        return False, f"Failed to launch Edge: {e}"

    # Wait for Edge to start and CDP to become available
    for i in range(15):  # Wait up to 15 seconds
        time.sleep(1)
        if _is_edge_cdp_available():
            log("    Edge launched successfully")
            return True, "Edge launched"

    return False, "Edge launched but remote debugging not responding. Try closing Edge manually and retry."


def _open_tab_urls(browser) -> List[str]:
    """URLs of every tab in the browser; unreadable tabs are skipped."""
    urls = []
    for ctx in browser.contexts:
        for page in ctx.pages:
            try:
                urls.append(page.url)
            except Exception:
                continue
    return urls


async def _find_royalmail_page(browser, log, wait_seconds=15, poll_seconds=1):
    """
    Return a royalmail.com tab to log in on, opening one if none arrives.

    Edge is launched with the login URL, but the debug port can answer before
    that tab exists or has navigated — and now that every run starts a fresh
    Edge, a slow machine hits that window every time. Edge may also show its
    own pages first (sync confirmation, session restore). So wait for the
    tab, and if it never appears, open the login page ourselves.

    Returns (page, error); the error lists the tabs seen, so a failure on a
    machine we cannot watch still says what Edge was showing.
    """
    deadline = time.time() + wait_seconds
    while True:
        for ctx in browser.contexts:
            for page in ctx.pages:
                try:
                    if 'royalmail.com' in page.url.lower():
                        return page, ""
                except Exception:
                    continue
        if time.time() >= deadline:
            break
        await asyncio.sleep(poll_seconds)

    seen = ', '.join(_open_tab_urls(browser)) or '(no tabs)'
    log(f"    No Royal Mail tab in Edge (tabs: {seen}); opening the login page")

    if not browser.contexts:
        return None, "Could not find Royal Mail page in Edge — Edge has no browser window open"

    try:
        page = await browser.contexts[0].new_page()
        await page.goto(OBA_LOGIN_URL, wait_until='domcontentloaded', timeout=30000)
        return page, ""
    except Exception as e:
        return None, f"Could not find Royal Mail page in Edge (tabs: {seen}); opening it failed: {e}"


async def _auto_login_to_oba(browser, log, timeout_ms=30000):
    """
    Automatically log in to Royal Mail OBA via the royalmail.com login page.

    Fills credentials, clicks login, clicks Access OBA, and waits for
    the OBA dashboard to load. Returns the OBA page on success.
    """
    creds = get_royalmail_credentials()
    if not creds.is_valid():
        return None, "Royal Mail credentials not configured. Set ROYALMAIL_EMAIL and ROYALMAIL_PASSWORD in .env"

    login_page, find_error = await _find_royalmail_page(browser, log)
    if not login_page:
        return None, find_error

    log("  Automating login...")

    try:
        # Wait for page to be ready
        await login_page.wait_for_timeout(3000)

        # Accept cookies if banner is present
        try:
            accept_btn = login_page.locator('button:has-text("Accept all")').first
            if await accept_btn.is_visible(timeout=3000):
                await accept_btn.click()
                log("    Cookies accepted")
                await login_page.wait_for_timeout(1000)
        except Exception:
            pass

        # Check if already on the OBA dashboard
        if 'oba.royalmail.com/irj/portal/oba' in login_page.url.lower():
            log("    Already on OBA dashboard")
            return login_page, ""

        # Detect page state: already logged in, or need to fill login form?
        try:
            page_content = await login_page.content()
        except Exception:
            # Page mid-navigation; treat as not logged in — the form-visibility
            # check below sorts out the real state.
            page_content = ''
        already_logged_in = 'You are logged in as' in page_content or 'Access OBA services' in page_content

        if not already_logged_in:
            # Need to fill login form
            try:
                email_field = login_page.get_by_role('textbox', name='Email address')
                if await email_field.is_visible(timeout=5000):
                    await email_field.fill(creds.email)
                    password_field = login_page.get_by_role('textbox', name='Password')
                    await password_field.fill(creds.password)
                    log("    Credentials entered")

                    # Click Log in
                    await login_page.locator('input[type="submit"][value="Log in"]').click()
                    log("    Clicked Log in, waiting...")
                    await login_page.wait_for_timeout(8000)

                    # Check for login errors in the VISIBLE text only — the raw
                    # HTML always contains 'invalid' (aria-invalid attributes),
                    # which used to cause false 'invalid email or password'
                    # reports. Report the actual on-screen error line.
                    try:
                        visible_text = await login_page.locator('body').inner_text()
                    except Exception:
                        visible_text = ''  # page navigating away = login accepted
                    error_keywords = (
                        'invalid', 'incorrect',
                        'check that you have a valid online business account',
                    )
                    for line in visible_text.splitlines():
                        stripped = line.strip()
                        if stripped and any(kw in stripped.lower() for kw in error_keywords):
                            return None, f"Login failed: {stripped[:200]}"
                else:
                    log("    Login form not visible, checking for other options...")
            except Exception as e:
                return None, f"Could not fill login form: {e}"

        # At this point we're either logged in with a cached session or just logged in.
        # Navigate through the two-step OBA access flow:
        #   Step 1: Click "Access OBA services" (if present)
        #   Step 2: Click "Access OBA" on the services page

        # Step 1: "Access OBA services" button (shown when already logged in)
        try:
            oba_services_btn = login_page.locator('a:has-text("Access OBA services")').first
            if await oba_services_btn.is_visible(timeout=5000):
                await oba_services_btn.click()
                log("    Clicked Access OBA services")
                await login_page.wait_for_timeout(5000)
        except Exception:
            pass

        # Step 2: "Access OBA" link (on the OBA Services page)
        try:
            access_link = login_page.locator('a:has-text("Access OBA")').first
            if await access_link.is_visible(timeout=8000):
                await access_link.click()
                log("    Clicked Access OBA, waiting for SSO redirect...")
                await login_page.wait_for_timeout(15000)
            else:
                return None, "Access OBA link not found"
        except Exception as e:
            return None, f"Failed to click Access OBA: {e}"

        # Wait for OBA dashboard to appear
        oba_page = await _wait_for_oba_dashboard(browser, log, timeout_seconds=60)
        if oba_page:
            log("    Login successful")
            return oba_page, ""

        return None, "Timed out waiting for OBA dashboard after login"

    except Exception as e:
        return None, f"Auto-login failed: {e}"


async def _find_content_frame(page):
    """Find the main content frame in the SAP portal."""
    # Check by URL pattern first (most reliable)
    for frame in page.frames:
        url = frame.url
        if 'getdealerfamily' in url or 'showbasket' in url or 'itemconfiguration' in url:
            return frame
    # Fallback: look for frames with order-related or navigation content
    for frame in page.frames:
        try:
            body = frame.locator('body')
            if await body.count() > 0:
                text = await body.inner_text()
                if any(kw in text for kw in [
                    'Create new Order', 'Emanifest ID', 'Choose an option',
                    'Your customer accounts', 'Your accounts',
                    'Manage your orders', 'Confirmed Sales Order',
                ]):
                    return frame
        except Exception:
            continue
    return None


async def _find_oba_page(browser):
    """Find the OBA dashboard page (oba.royalmail.com, not royalmail.com/oba)."""
    for ctx in browser.contexts:
        for page in ctx.pages:
            url = page.url.lower()
            # Must be on the actual OBA portal, not the royalmail.com intermediate pages
            if 'oba.royalmail.com/irj/portal/oba' in url:
                return page
    return None


async def _wait_for_oba_dashboard(browser, log, timeout_seconds=300):
    """
    Wait for the user to log in and reach the OBA dashboard.

    Polls every 3 seconds checking if any page has reached the OBA portal.
    Returns the page once the user is on the OBA dashboard.
    """
    log("  Waiting for OBA login...")
    start = time.time()

    while time.time() - start < timeout_seconds:
        # Check all pages for OBA dashboard indicators
        for ctx in browser.contexts:
            for page in ctx.pages:
                try:
                    url = page.url.lower()
                    if 'oba.royalmail.com/irj/portal/oba' in url:
                        log("    OBA dashboard detected")
                        return page
                except Exception:
                    continue

        await asyncio.sleep(3)

    return None


async def _read_select_options(select_locator):
    """Return [(value, label)] for a <select>, as the portal rendered it."""
    return await select_locator.evaluate(
        "el => Array.from(el.options).map(o => [o.value, o.textContent])"
    )


async def _select_matching_option(frame, field_suffix: str, wanted: str, what: str):
    """
    Select the option matching `wanted` on the item-configuration form.

    Fields are located by name suffix rather than the '1.1.' prefix the
    configuration page happens to use, so a different line prefix cannot
    silently skip a fill.
    """
    select_locator = frame.locator(f'select[name$=".{field_suffix}"]').first
    options = await _read_select_options(select_locator)
    value = match_option_value(options, wanted, what)
    await select_locator.select_option(value)


async def _fill_by_suffix(frame, field_suffix: str, value: str):
    """Fill an item-configuration input located by its name suffix."""
    field = frame.locator(f'input[name$=".{field_suffix}"]').first
    await field.click()
    await field.fill(value)


async def _configure_line(page, frame, line: CountryLine, index: int, log):
    """
    Open line `index`'s configuration and file it against its destination.

    Returns the content frame after the configuration is accepted.
    """
    spec = SUPPORTED_COUNTRIES[line.country]
    product_code = PRODUCT_CODES[line.format_type]

    log(f"  Configuring line {index + 1}: {product_code} {line.country} ({line.format_type})...")
    config_links = frame.locator('a[onclick*="itemconfig"]')
    link_count = await config_links.count()
    if link_count <= index:
        raise RuntimeError(
            f"Expected at least {index + 1} configurable order lines but the "
            f"form shows {link_count}. The order form may have rejected a "
            f"duplicate product code."
        )
    await config_links.nth(index).click()
    await page.wait_for_timeout(3000)

    frame = await _find_content_frame(page)
    if not frame:
        raise RuntimeError(f"Lost the content frame opening line {index + 1}'s configuration")

    await _select_matching_option(frame, 'ZZOBA_INTL_REGION', spec.region, 'region')
    await _select_matching_option(frame, 'ZZOBA_CNTRY_DESP', line.country, 'country')
    await _fill_by_suffix(frame, 'ZZOBA_TOTAL_ITEM_QTY', str(line.items))
    await _fill_by_suffix(frame, 'ZZUNITWGT', str(line.avg_weight_grams))
    log(f"    {line.country} {line.format_type}: {line.items} items, "
        f"{line.avg_weight_grams}g avg, region {spec.region}")

    await frame.locator('text=Accept').first.click()
    await page.wait_for_timeout(3000)

    frame = await _find_content_frame(page)
    if not frame:
        raise RuntimeError(f"Lost the content frame after accepting line {index + 1}")
    return frame


async def _create_order(page, frame, portal_input: RoyalMailPortalInput, log, timeout_ms: int):
    """
    Create one order covering every destination in the posting.

    Each CountryLine becomes one line on the order form: its product code
    goes in a product row, then its configuration carries the destination
    country, OBA region, item count and average item weight.
    """
    lines = portal_input.ordered_lines
    if not lines:
        return False, "No Royal Mail volumes to submit"

    # Check every line is fileable before touching the form. Without this a
    # stray destination or format surfaces as a bare KeyError — and KeyError
    # is a LookupError, so it would be caught below and reported as just the
    # offending value with no explanation.
    for line in lines:
        if line.country not in SUPPORTED_COUNTRIES:
            return False, (
                f"Cannot file destination '{line.country}' through OBA. "
                f"Supported: {', '.join(SUPPORTED_COUNTRIES)}."
            )
        if line.format_type not in PRODUCT_CODES:
            return False, (
                f"Cannot file format '{line.format_type}' through OBA. "
                f"Supported: {', '.join(PRODUCT_CODES)}."
            )

    # Click 'Create new Order' link
    log("  Creating new order...")
    order_link = frame.locator('a:has-text("Order")').first
    await order_link.click()
    await page.wait_for_timeout(5000)

    # Re-find the content frame (may have changed after navigation)
    frame = await _find_content_frame(page)
    if not frame:
        return False, "Could not find order form after clicking Create Order"

    # Fill PO number
    po_field = frame.locator('#poNumber')
    await po_field.click()
    await po_field.fill(portal_input.po_number)
    log(f"    PO number: {portal_input.po_number}")

    # One product row per destination/format line
    for index, line in enumerate(lines, start=1):
        product_code = PRODUCT_CODES[line.format_type]
        prod_field = frame.locator(f'input[name="product[{index}]"]')
        if await prod_field.count() == 0:
            return False, (
                f"The order form has no product row {index}; it cannot hold "
                f"{len(lines)} lines. Reduce the destinations in this posting "
                f"or split it across orders."
            )
        await prod_field.click()
        await prod_field.fill('')
        await prod_field.type(product_code)
        await prod_field.press('Tab')
        log(f"    Product [{index}]: {product_code} — {line.country} {line.format_type}")
        await page.wait_for_timeout(500)

    # Click Update order to validate product codes
    log("  Updating order...")
    update_link = frame.locator('text=Update').first
    await update_link.click()
    await page.wait_for_timeout(5000)

    # Re-find frame
    frame = await _find_content_frame(page)
    if not frame:
        return False, "Could not find order form after Update"

    # Configure each line in the same order the product rows were entered
    try:
        for index, line in enumerate(lines):
            frame = await _configure_line(page, frame, line, index, log)
    except (LookupError, RuntimeError) as e:
        return False, str(e)

    return True, ""


async def _confirm_and_save(page, frame, portal_input, output_dir, log, timeout_ms):
    """Confirm the order and save the confirmation page as PDF."""

    log("  Confirming order...")

    # Handle any confirmation dialog
    page.on('dialog', lambda dialog: asyncio.ensure_future(dialog.accept()))

    confirm_link = frame.locator('a:has-text("Confirm order")').first
    await confirm_link.click()
    await page.wait_for_timeout(8000)

    # Re-find frame — check for confirmation text
    frame = await _find_content_frame(page)
    if not frame:
        for f in page.frames:
            try:
                text = await f.locator('body').inner_text()
                if 'Confirmed Sales Order' in text or 'Thank you' in text:
                    frame = f
                    break
            except Exception:
                continue

    if not frame:
        return False, "Could not find confirmation page", None

    # Verify confirmation
    body_text = await frame.locator('body').inner_text()
    if 'Confirmed Sales Order' not in body_text and 'Thank you' not in body_text:
        screenshot_path = os.path.join(output_dir, "rm_debug_confirm.png")
        await page.screenshot(path=screenshot_path)
        return False, "Order confirmation not detected. Check rm_debug_confirm.png", None

    log("    Order confirmed")

    # Save the sales order confirmation as PDF.
    # The confirmation lives inside a SAP portal frame. We extract the frame's
    # full HTML (including styles), load it into a standalone page, and print
    # that to PDF — producing the same clean output as the portal's Print button.
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_filename = f"Royal_Mail_{portal_input.po_number}_{timestamp}.pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)

    try:
        # Extract the frame's complete HTML (document + inline styles)
        # Inject a <base> tag so relative image URLs resolve against the OBA domain
        frame_html = await frame.evaluate("""() => {
            const html = document.documentElement.cloneNode(true);
            const head = html.querySelector('head') || html;
            const base = document.createElement('base');
            base.href = 'https://www.oba.royalmail.com/';
            head.insertBefore(base, head.firstChild);
            return html.outerHTML;
        }""")

        # Open a new page and set its content to the extracted HTML
        new_page = await page.context.new_page()
        await new_page.set_content(frame_html, wait_until='networkidle')
        await new_page.wait_for_timeout(1000)

        # Use CDP printToPDF on the standalone page
        import base64
        cdp = await new_page.context.new_cdp_session(new_page)
        result = await cdp.send("Page.printToPDF", {
            "printBackground": True,
            "preferCSSPageSize": False,
            "landscape": False,
            "paperWidth": 8.27,   # A4
            "paperHeight": 11.69, # A4
            "marginTop": 0.4,
            "marginBottom": 0.4,
            "marginLeft": 0.4,
            "marginRight": 0.4,
        })
        await cdp.detach()

        pdf_bytes = base64.b64decode(result["data"])
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)
        log(f"    Saved: {pdf_filename}")

        await new_page.close()
    except Exception as e:
        log(f"    PDF generation failed ({e}), saving screenshot fallback...")
        try:
            png_filename = f"Royal_Mail_{portal_input.po_number}_{timestamp}.png"
            pdf_path = os.path.join(output_dir, png_filename)
            await page.screenshot(path=pdf_path, full_page=True)
            log(f"    Saved screenshot: {png_filename}")
        except Exception:
            pass

    return True, "Order confirmed and saved", pdf_path


async def _submit_to_royalmail_portal_impl(
    portal_input: RoyalMailPortalInput,
    output_dir: str,
    auto_print: bool = True,
    log_callback: Optional[Callable[[str], None]] = None,
    timeout_ms: int = 30000,
) -> tuple[bool, str]:
    """
    Submit order to Royal Mail OBA portal.

    Connects to Edge via CDP (Edge must already be running with debugging).
    The user must be logged in to the OBA dashboard.
    """
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return False, "Playwright not installed. Run: pip install playwright && playwright install chromium"

    def log(msg):
        if log_callback:
            log_callback(msg)

    downloaded_file = None

    try:
        async with async_playwright() as p:
            # Connect to Edge via CDP
            log("  Connecting to Edge browser...")
            try:
                browser = await p.chromium.connect_over_cdp(CDP_URL)
            except Exception as e:
                return False, f"Could not connect to Edge browser: {e}"

            oba_page = None
            try:
                # Check if already on OBA dashboard
                oba_page = await _find_oba_page(browser)

                if not oba_page:
                    # Try automatic login
                    oba_page, login_error = await _auto_login_to_oba(browser, log, timeout_ms)

                    if not oba_page:
                        log(f"    Auto-login failed: {login_error}")
                        return False, f"Auto-login failed: {login_error}"

                log(f"    Connected to OBA: {oba_page.url}")

                # Find the content frame
                frame = await _find_content_frame(oba_page)

                if not frame:
                    for idx, f in enumerate(oba_page.frames):
                        try:
                            log(f"    Frame {idx}: {f.url[:80] if f.url else '(no url)'}")
                        except Exception:
                            pass
                    screenshot_path = os.path.join(output_dir, "rm_debug_noframe.png")
                    await oba_page.screenshot(path=screenshot_path)
                    return False, "Could not find OBA content frame. Check rm_debug_noframe.png"

                # Handle 'Choose an option' page: 'Your accounts' (exact — the page
                # also has a 'Your customer accounts' link) goes straight to the
                # 'Manage your orders' page since the July 2026 portal change.
                body_text = await frame.locator('body').inner_text()
                if 'Choose an option' in body_text:
                    log("  Selecting 'Your accounts'...")
                    accounts_link = frame.get_by_role("link", name="Your accounts", exact=True)
                    if await accounts_link.count() == 0:
                        accounts_link = frame.locator('a:has-text("Your accounts")')
                    await accounts_link.first.click()
                    await oba_page.wait_for_timeout(5000)
                    frame = await _find_content_frame(oba_page)

                if not frame:
                    screenshot_path = os.path.join(output_dir, "rm_debug_navigate.png")
                    await oba_page.screenshot(path=screenshot_path)
                    return False, "Could not navigate to order form. Check rm_debug_navigate.png"

                # Create and configure the order
                success, error = await _create_order(oba_page, frame, portal_input, log, timeout_ms)
                if not success:
                    screenshot_path = os.path.join(output_dir, "rm_debug_order.png")
                    await oba_page.screenshot(path=screenshot_path)
                    return False, f"Order creation failed: {error}"

                # Re-find frame for confirmation
                frame = await _find_content_frame(oba_page)
                if not frame:
                    return False, "Lost content frame before confirmation"

                # Confirm and save
                success, message, pdf_path = await _confirm_and_save(
                    oba_page, frame, portal_input, output_dir, log, timeout_ms
                )

                if not success:
                    return False, message

                downloaded_file = pdf_path

                # Print if enabled — reuse the app's print function (SumatraPDF > Adobe /t > fallback)
                if auto_print and downloaded_file and os.path.exists(downloaded_file):
                    log("  Printing confirmation...")
                    from gui import print_pdf_file
                    print_success, print_msg = print_pdf_file(downloaded_file)
                    if print_success:
                        log(f"  {print_msg}")
                        return True, "Order confirmed, saved and printed successfully"
                    else:
                        log(f"  Print issue: {print_msg}")
                        return True, f"Order confirmed and saved. Print failed: {print_msg}"
                elif downloaded_file:
                    return True, "Order confirmed and saved successfully"
                else:
                    return True, message

            except Exception as e:
                # Capture crash evidence while the CDP connection is still open
                # (outside this block Playwright has shut down and the page is gone)
                message = f"Portal automation failed: {e}"
                if oba_page:
                    try:
                        screenshot_path = os.path.join(output_dir, "rm_debug_crash.png")
                        await oba_page.screenshot(path=screenshot_path)
                        for idx, f in enumerate(oba_page.frames):
                            try:
                                log(f"    Frame {idx}: {f.url[:100] if f.url else '(no url)'}")
                            except Exception:
                                pass
                        message += " (see rm_debug_crash.png)"
                    except Exception:
                        pass
                return False, message

    except Exception as e:
        return False, f"Portal automation failed: {str(e)}"


async def submit_to_royalmail_portal(
    portal_input: RoyalMailPortalInput,
    output_dir: str,
    auto_print: bool = True,
    log_callback: Optional[Callable[[str], None]] = None,
    timeout_ms: int = 30000,
    retry_count: int = 1
) -> tuple[bool, str]:
    """
    Submit order to Royal Mail OBA portal with retry logic.

    Always ends by logging out of OBA and closing the automation browser, so
    a run never leaves a logged-in Edge tab behind. That teardown sits
    outside the retry loop, so retries still share one browser session.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    last_error = None

    try:
        for attempt in range(retry_count + 1):
            if attempt > 0:
                log(f"\n  Retry attempt {attempt} of {retry_count}...")

            success, message = await _submit_to_royalmail_portal_impl(
                portal_input, output_dir, auto_print, log_callback, timeout_ms
            )

            if success:
                return success, message

            last_error = message
            if "Timeout" in message or "timeout" in message:
                if attempt < retry_count:
                    log("  Timeout occurred, will retry...")
                    continue
            else:
                break

        return False, last_error or "Portal automation failed after retries"

    finally:
        # Teardown must not turn a completed order into a reported failure
        try:
            await _logout_and_close_edge(log)
        except Exception as e:
            log(f"  Session cleanup issue: {e}")


def run_royalmail_upload(
    portal_input: RoyalMailPortalInput,
    output_dir: str,
    auto_print: bool = True,
    log_callback: Optional[Callable[[str], None]] = None
) -> tuple[bool, str]:
    """Synchronous wrapper for the async Royal Mail portal function."""
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(
                submit_to_royalmail_portal(
                    portal_input, output_dir, auto_print, log_callback
                )
            )
        finally:
            loop.close()
    except Exception as e:
        return False, f"Upload error: {str(e)}"
