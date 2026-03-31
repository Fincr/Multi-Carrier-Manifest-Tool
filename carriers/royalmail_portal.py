"""
Royal Mail OBA portal automation.

Handles the Royal Mail Online Business Account portal workflow:
1. Auto-launch Edge with remote debugging enabled
2. Navigate Edge to the OBA login page
3. Wait for user to log in and reach the OBA dashboard
4. Automate: select posting location, create order, fill form, confirm
5. Save confirmation as PDF and optionally print

Akamai bot protection on royalmail.com blocks automated Chromium,
so we use the user's real Edge browser via CDP (Chrome DevTools Protocol).
Edge is launched automatically — the user only needs to log in.
"""

import os
import asyncio
import subprocess
import time
from datetime import datetime
from typing import Callable, Optional
from dataclasses import dataclass

from .royalmail import RoyalMailData
from core.credentials import get_royalmail_credentials


# Portal constants
OBA_LOGIN_URL = "https://www.royalmail.com/discounts-payment/credit-account/online-business-account"
POSTING_LOCATION = "9000227875"
CDP_PORT = 9222
CDP_URL = f"http://127.0.0.1:{CDP_PORT}"

# Product codes
PRODUCT_CODE_LETTERS = "PS5"
PRODUCT_CODE_FLATS = "PS7"

# Portal form values
REGION = "EUROPEAN UNION"
COUNTRY = "IRELAND (REPUBLIC OF)"

# Edge executable paths (checked in order)
EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


@dataclass
class RoyalMailPortalInput:
    """Combined input for the Royal Mail portal from one or both carrier sheets."""
    po_number: str
    flats_items: int = 0
    flats_weight_kg: float = 0.0
    letters_items: int = 0
    letters_weight_kg: float = 0.0

    @property
    def has_letters(self) -> bool:
        return self.letters_items > 0

    @property
    def has_flats(self) -> bool:
        return self.flats_items > 0

    @property
    def avg_letter_weight_grams(self) -> int:
        if self.letters_items > 0:
            return round(self.letters_weight_kg * 1000 / self.letters_items)
        return 0

    @property
    def avg_flat_weight_grams(self) -> int:
        if self.flats_items > 0:
            return round(self.flats_weight_kg * 1000 / self.flats_items)
        return 0


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
        subprocess.Popen(
            [
                edge_path,
                f"--remote-debugging-port={CDP_PORT}",
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


async def _auto_login_to_oba(browser, log, timeout_ms=30000):
    """
    Automatically log in to Royal Mail OBA via the royalmail.com login page.

    Fills credentials, clicks login, clicks Access OBA, and waits for
    the OBA dashboard to load. Returns the OBA page on success.
    """
    creds = get_royalmail_credentials()
    if not creds.is_valid():
        return None, "Royal Mail credentials not configured. Set ROYALMAIL_EMAIL and ROYALMAIL_PASSWORD in .env"

    # Find the login page (Edge was launched with the login URL)
    login_page = None
    for ctx in browser.contexts:
        for page in ctx.pages:
            if 'royalmail.com' in page.url.lower():
                login_page = page
                break
        if login_page:
            break

    if not login_page:
        return None, "Could not find Royal Mail page in Edge"

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
        page_content = await login_page.content()
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

                    # Check for login errors
                    content = await login_page.content()
                    if 'invalid' in content.lower() or 'incorrect' in content.lower():
                        return None, "Login failed - invalid email or password"
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
                    'Choose the site', 'Your customer accounts', 'Your accounts',
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


async def _create_order(page, frame, portal_input: RoyalMailPortalInput, log, timeout_ms: int):
    """Create and configure an order with the given products."""

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

    # Enter product codes
    product_row = 1
    if portal_input.has_letters:
        prod_field = frame.locator(f'input[name="product[{product_row}]"]')
        await prod_field.click()
        await prod_field.fill('')
        await prod_field.type(PRODUCT_CODE_LETTERS)
        await prod_field.press('Tab')
        log(f"    Product [{product_row}]: {PRODUCT_CODE_LETTERS} (Letters)")
        product_row += 1
        await page.wait_for_timeout(500)

    if portal_input.has_flats:
        prod_field = frame.locator(f'input[name="product[{product_row}]"]')
        await prod_field.click()
        await prod_field.fill('')
        await prod_field.type(PRODUCT_CODE_FLATS)
        await prod_field.press('Tab')
        log(f"    Product [{product_row}]: {PRODUCT_CODE_FLATS} (Flats)")
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

    # Configure each product
    if portal_input.has_letters:
        log("  Configuring PS5 (Letters)...")
        config_links = frame.locator('a[onclick*="itemconfig"]')
        await config_links.nth(0).click()
        await page.wait_for_timeout(3000)

        frame = await _find_content_frame(page)
        await frame.locator('select[name="1.1.ZZOBA_INTL_REGION"]').select_option(REGION)
        await frame.locator('select[name="1.1.ZZOBA_CNTRY_DESP"]').select_option(COUNTRY)
        items_field = frame.locator('input[name="1.1.ZZOBA_TOTAL_ITEM_QTY"]')
        await items_field.click()
        await items_field.fill(str(portal_input.letters_items))
        weight_field = frame.locator('input[name="1.1.ZZUNITWGT"]')
        await weight_field.click()
        await weight_field.fill(str(portal_input.avg_letter_weight_grams))
        log(f"    Letters: {portal_input.letters_items} items, {portal_input.avg_letter_weight_grams}g avg")

        await frame.locator('text=Accept').first.click()
        await page.wait_for_timeout(3000)
        frame = await _find_content_frame(page)

    if portal_input.has_flats:
        log("  Configuring PS7 (Flats)...")
        config_links = frame.locator('a[onclick*="itemconfig"]')
        # If both products exist, PS7 config is the last itemconfig link
        config_idx = (await config_links.count()) - 1
        await config_links.nth(config_idx).click()
        await page.wait_for_timeout(3000)

        frame = await _find_content_frame(page)
        await frame.locator('select[name="1.1.ZZOBA_INTL_REGION"]').select_option(REGION)
        await frame.locator('select[name="1.1.ZZOBA_CNTRY_DESP"]').select_option(COUNTRY)
        items_field = frame.locator('input[name="1.1.ZZOBA_TOTAL_ITEM_QTY"]')
        await items_field.click()
        await items_field.fill(str(portal_input.flats_items))
        weight_field = frame.locator('input[name="1.1.ZZUNITWGT"]')
        await weight_field.click()
        await weight_field.fill(str(portal_input.avg_flat_weight_grams))
        log(f"    Flats: {portal_input.flats_items} items, {portal_input.avg_flat_weight_grams}g avg")

        await frame.locator('text=Accept').first.click()
        await page.wait_for_timeout(3000)
        frame = await _find_content_frame(page)

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
                # Navigate through the dashboard — try to find posting location or navigation links
                log("  Navigating to account...")

                # Debug: list all frames to help diagnose
                for idx, f in enumerate(oba_page.frames):
                    try:
                        furl = f.url[:80] if f.url else '(no url)'
                        log(f"    Frame {idx}: {furl}")
                    except Exception:
                        pass

                # Try clicking posting location in any frame
                for f in oba_page.frames:
                    try:
                        link = f.locator(f'a:has-text("{POSTING_LOCATION}")').first
                        if await link.is_visible(timeout=2000):
                            await link.click()
                            log(f"    Clicked posting location {POSTING_LOCATION}")
                            await oba_page.wait_for_timeout(5000)
                            frame = await _find_content_frame(oba_page)
                            break
                    except Exception:
                        continue

            if not frame:
                # Take debug screenshot
                screenshot_path = os.path.join(output_dir, "rm_debug_noframe.png")
                await oba_page.screenshot(path=screenshot_path)
                return False, f"Could not find OBA content frame. Check rm_debug_noframe.png"

            # Handle 'Choose an option' page
            body_text = await frame.locator('body').inner_text()
            if 'Choose an option' in body_text:
                log("  Selecting 'Your accounts'...")
                accounts_link = frame.locator('a:has-text("Your accounts")').first
                await accounts_link.click()
                await oba_page.wait_for_timeout(5000)
                frame = await _find_content_frame(oba_page)
                body_text = await frame.locator('body').inner_text() if frame else ''

            # Handle 'Choose the site' page
            if 'Choose the site' in body_text:
                log(f"  Selecting posting location {POSTING_LOCATION}...")
                for f in oba_page.frames:
                    try:
                        loc_link = f.locator(f'a:has-text("{POSTING_LOCATION}")').first
                        if await loc_link.is_visible(timeout=2000):
                            await loc_link.click()
                            await oba_page.wait_for_timeout(5000)
                            frame = await _find_content_frame(oba_page)
                            break
                    except Exception:
                        continue

            if not frame:
                screenshot_path = os.path.join(output_dir, "rm_debug_navigate.png")
                await oba_page.screenshot(path=screenshot_path)
                return False, f"Could not navigate to order form. Check rm_debug_navigate.png"

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
        return False, f"Portal automation failed: {str(e)}"


async def submit_to_royalmail_portal(
    portal_input: RoyalMailPortalInput,
    output_dir: str,
    auto_print: bool = True,
    log_callback: Optional[Callable[[str], None]] = None,
    timeout_ms: int = 30000,
    retry_count: int = 1
) -> tuple[bool, str]:
    """Submit order to Royal Mail OBA portal with retry logic."""
    def log(msg):
        if log_callback:
            log_callback(msg)

    last_error = None

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
