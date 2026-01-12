"""
Spring Global Delivery Solutions Portal Automation

Robust portal automation with comprehensive error handling, retry logic,
and graceful degradation for unreliable portal behaviour.

Known portal issues handled:
- Post-login hang (page doesn't fully load after authentication)
- "Upload Multiple Orders" button not found (dynamic UI loading)
- "Unexpected error" after CSV upload
- "Unexpected error" when downloading PDF manifest
"""

import os
import asyncio
from datetime import datetime
from dataclasses import dataclass
from enum import Enum
from typing import Optional, Callable, Tuple


class SpringPortalStage(Enum):
    """Stages of the Spring portal workflow."""
    INIT = "initialisation"
    LOGIN = "login"
    POST_LOGIN = "post_login_navigation"
    FIND_UPLOAD = "finding_upload_button"
    UPLOAD_FILE = "uploading_file"
    VIEW_ORDERS = "viewing_uploaded_orders"
    SELECT_ORDER = "selecting_order"
    DOWNLOAD_PDF = "downloading_pdf"
    COMPLETE = "complete"


@dataclass
class SpringPortalResult:
    """Result of a Spring portal operation."""
    success: bool
    message: str
    stage_reached: SpringPortalStage
    pdf_downloaded: bool = False
    pdf_path: Optional[str] = None
    requires_manual_intervention: bool = False
    
    @property
    def partial_success(self) -> bool:
        """True if upload succeeded but PDF download failed."""
        return self.stage_reached in (
            SpringPortalStage.VIEW_ORDERS,
            SpringPortalStage.SELECT_ORDER,
        )


class SpringPortalConfig:
    """Configuration for Spring portal automation."""
    
    def __init__(
        self,
        timeout_ms: int = 30000,
        retry_count: int = 2,
        stage_retry_count: int = 2,
        inter_stage_delay_ms: int = 2000,
        post_login_wait_ms: int = 5000,
        post_upload_wait_ms: int = 5000,
        pre_print_delay_ms: int = 4000,  # Delay before clicking Print to avoid "unexpected error"
    ):
        self.timeout_ms = timeout_ms
        self.retry_count = retry_count  # Full workflow retries
        self.stage_retry_count = stage_retry_count  # Per-stage retries
        self.inter_stage_delay_ms = inter_stage_delay_ms
        self.post_login_wait_ms = post_login_wait_ms
        self.post_upload_wait_ms = post_upload_wait_ms
        self.pre_print_delay_ms = pre_print_delay_ms


async def _wait_for_page_stable(page, timeout_ms: int = 5000, check_interval_ms: int = 500) -> bool:
    """
    Wait for page to become stable (no more network activity).
    More reliable than wait_for_load_state("networkidle") for flaky portals.
    
    Returns True if page stabilised, False if timeout.
    """
    try:
        # First try the standard networkidle
        await page.wait_for_load_state("networkidle", timeout=timeout_ms)
        return True
    except Exception:
        pass

    # Fallback: manual stability check
    start_time = asyncio.get_event_loop().time()
    max_time = start_time + (timeout_ms / 1000)

    while asyncio.get_event_loop().time() < max_time:
        try:
            # Try a shorter networkidle wait
            await page.wait_for_load_state("networkidle", timeout=check_interval_ms)
            return True
        except Exception:
            await page.wait_for_timeout(check_interval_ms)

    return False


async def _safe_click(page, selectors: list, description: str, timeout_ms: int = 10000, log: Callable = None) -> bool:
    """
    Safely click an element using multiple selector fallbacks.
    
    Args:
        page: Playwright page
        selectors: List of selectors to try in order
        description: Human-readable description for logging
        timeout_ms: Timeout for each selector attempt
        log: Optional logging callback
    
    Returns:
        True if click succeeded, False otherwise
    """
    for selector in selectors:
        try:
            element = page.locator(selector).first
            if await element.is_visible(timeout=timeout_ms // len(selectors)):
                await element.click()
                if log:
                    log(f"    ✓ Found {description} with selector: {selector[:50]}...")
                return True
        except Exception:
            continue
    return False


async def _check_for_portal_error(page) -> Tuple[bool, str]:
    """
    Check if the portal is displaying an error message.
    
    Returns:
        (has_error: bool, error_message: str)
    """
    error_selectors = [
        '.error-message',
        '.alert-danger',
        '.notification-error',
        '[class*="error"]',
        '[role="alert"]',
    ]
    
    error_keywords = [
        'unexpected error',
        'something went wrong',
        'please try again',
        'error occurred',
        'unable to process',
        'server error',
    ]
    
    # Check for error elements
    for selector in error_selectors:
        try:
            element = page.locator(selector).first
            if await element.is_visible(timeout=1000):
                text = await element.text_content()
                if text and any(kw in text.lower() for kw in error_keywords):
                    return True, text.strip()
        except Exception:
            continue
    
    # Check page content for error messages
    try:
        page_text = await page.inner_text('body')
        for keyword in error_keywords:
            if keyword in page_text.lower():
                # Try to extract context around the error
                idx = page_text.lower().find(keyword)
                start = max(0, idx - 50)
                end = min(len(page_text), idx + 100)
                context = page_text[start:end].strip()
                return True, context
    except Exception:
        pass
    
    return False, ""


async def _stage_login(
    page,
    email: str,
    password: str,
    config: SpringPortalConfig,
    log: Callable,
) -> Tuple[bool, str]:
    """
    Perform login to Spring portal with retry logic.
    
    Returns:
        (success: bool, error_message: str)
    """
    LOGIN_URL = "https://my.spring-gds.com/"
    
    for attempt in range(config.stage_retry_count + 1):
        try:
            if attempt > 0:
                log(f"    ⟳ Login retry {attempt}...")
                await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=config.timeout_ms)
                await page.wait_for_timeout(2000)
            
            # Enter email
            email_selectors = [
                'input[type="email"]',
                'input[name="email"]',
                'input[placeholder*="mail"]',
                'input[id*="email"]',
            ]
            
            email_entered = False
            for selector in email_selectors:
                try:
                    element = page.locator(selector).first
                    if await element.is_visible(timeout=config.timeout_ms // 2):
                        await element.fill(email)
                        email_entered = True
                        break
                except Exception:
                    continue
            
            if not email_entered:
                continue
            
            # Click Next
            next_clicked = await _safe_click(
                page,
                ['button:has-text("Next")', 'button:has-text("Continue")', 'button[type="submit"]'],
                "Next button",
                config.timeout_ms // 2,
                log,
            )
            
            if not next_clicked:
                continue
            
            # Wait for password field with stability check
            await _wait_for_page_stable(page, config.timeout_ms // 2)
            
            # Enter password
            try:
                await page.wait_for_selector('input[type="password"]', timeout=config.timeout_ms // 2)
                await page.fill('input[type="password"]', password)
            except Exception:
                continue
            
            # Click Sign in
            signin_clicked = await _safe_click(
                page,
                ['button:has-text("Sign in")', 'button:has-text("Login")', 'button[type="submit"]'],
                "Sign in button",
                config.timeout_ms // 2,
                log,
            )
            
            if not signin_clicked:
                continue
            
            # Wait for post-login navigation - this is where hangs often occur
            # Use a more resilient approach: wait for any of several indicators
            log("    Waiting for dashboard...")
            
            post_login_success = False
            
            # Strategy 1: Wait for networkidle (may hang)
            try:
                await page.wait_for_load_state("networkidle", timeout=config.post_login_wait_ms)
                post_login_success = True
            except Exception:
                pass
            
            # Strategy 2: Check for known dashboard elements
            if not post_login_success:
                dashboard_indicators = [
                    'text="Upload Multiple Orders"',
                    'text="Dashboard"',
                    'text="My Orders"',
                    'text="Welcome"',
                    '[href*="upload"]',
                    '[href*="order"]',
                ]
                
                for indicator in dashboard_indicators:
                    try:
                        element = page.locator(indicator).first
                        if await element.is_visible(timeout=2000):
                            post_login_success = True
                            break
                    except Exception:
                        continue
            
            # Strategy 3: Just wait and check we're not on login page
            if not post_login_success:
                await page.wait_for_timeout(config.post_login_wait_ms)
                current_url = page.url
                if "login" not in current_url.lower() and "signin" not in current_url.lower():
                    post_login_success = True
            
            if post_login_success:
                log("    ✓ Login successful")
                return True, ""
            
        except Exception as e:
            if attempt == config.stage_retry_count:
                return False, f"Login failed after retries: {str(e)}"
    
    return False, "Login failed - could not authenticate"


async def _stage_navigate_to_upload(
    page,
    config: SpringPortalConfig,
    log: Callable,
) -> Tuple[bool, str]:
    """
    Navigate to Upload Multiple Orders page with retry logic.
    
    Returns:
        (success: bool, error_message: str)
    """
    upload_selectors = [
        'text="Upload Multiple Orders"',
        'text="Upload multiple orders"',
        'a:has-text("Upload Multiple")',
        'a:has-text("Multiple Orders")',
        'text="Upload Orders"',
        'text="Bulk Upload"',
        'text="Import Orders"',
        'a:has-text("Upload")',
        'button:has-text("Upload")',
        '[href*="upload"]',
        '[href*="Upload"]',
        '[href*="multiple"]',
        '[href*="bulk"]',
    ]
    
    for attempt in range(config.stage_retry_count + 1):
        if attempt > 0:
            log(f"    ⟳ Navigation retry {attempt}...")
            await page.wait_for_timeout(config.inter_stage_delay_ms)
            
            # Try refreshing the page
            try:
                await page.reload(wait_until="networkidle", timeout=config.timeout_ms // 2)
            except Exception:
                await page.wait_for_timeout(2000)
        
        clicked = await _safe_click(
            page,
            upload_selectors,
            "Upload button",
            config.timeout_ms,
            log,
        )
        
        if clicked:
            await _wait_for_page_stable(page, config.timeout_ms // 2)
            await page.wait_for_timeout(config.inter_stage_delay_ms)
            
            # Verify we're on the upload page
            try:
                file_input = page.locator('input[type="file"]').first
                if await file_input.is_visible(timeout=5000):
                    log("    ✓ Upload page loaded")
                    return True, ""
            except Exception:
                pass
    
    return False, "Could not find or navigate to Upload Multiple Orders page"


async def _stage_upload_file(
    page,
    file_path: str,
    config: SpringPortalConfig,
    log: Callable,
) -> Tuple[bool, str]:
    """
    Upload the manifest file with retry logic and error detection.
    
    Returns:
        (success: bool, error_message: str)
    """
    for attempt in range(config.stage_retry_count + 1):
        if attempt > 0:
            log(f"    ⟳ Upload retry {attempt}...")
            # Navigate back to upload page
            await page.go_back()
            await page.wait_for_timeout(config.inter_stage_delay_ms)
        
        try:
            # Find and use file input
            file_input = page.locator('input[type="file"]').first
            await file_input.set_input_files(file_path)
            log(f"    ✓ File selected: {os.path.basename(file_path)}")
            
            # Wait for upload processing - portal needs time to process
            log("    Waiting for file validation (7 seconds)...")
            await page.wait_for_timeout(7000)
            
            # Check for portal errors
            has_error, error_msg = await _check_for_portal_error(page)
            if has_error:
                log(f"    ⚠ Portal error detected: {error_msg[:100]}...")
                if attempt < config.stage_retry_count:
                    continue
                return False, f"Portal error during upload: {error_msg[:200]}"
            
            # Look for success indicators
            success_indicators = [
                'text="View uploaded orders"',
                'text="Upload successful"',
                'text="Orders uploaded"',
                'button:has-text("View")',
                '.success',
                '.alert-success',
            ]
            
            for indicator in success_indicators:
                try:
                    element = page.locator(indicator).first
                    if await element.is_visible(timeout=2000):
                        log("    ✓ Upload validation passed")
                        return True, ""
                except Exception:
                    continue
            
            # If no explicit success/error, assume success and continue
            log("    ✓ Upload completed (no validation errors detected)")
            return True, ""
            
        except Exception as e:
            if attempt == config.stage_retry_count:
                return False, f"File upload failed: {str(e)}"
    
    return False, "File upload failed after retries"


async def _stage_view_and_select_order(
    page,
    po_number: str,
    config: SpringPortalConfig,
    log: Callable,
) -> Tuple[bool, str, int]:
    """
    Navigate to view orders and select ALL uploaded orders matching the PO number
    that were created TODAY.
    
    Spring creates separate orders for STANDARD MAIL SORTED (Economy) and 
    PREMIUM MAIL SORTED (Priority), so we need to select all rows with the PO.
    
    We filter by today's date to avoid selecting old orders with the same PO.
    
    If "View uploaded orders" button is not found, falls back to navigating
    to "Order confirmation" in the left menu.
    
    Returns:
        (success: bool, error_message: str, orders_selected: int)
    """
    # Try clicking "View uploaded orders" first
    view_selectors = [
        'button:has-text("View uploaded orders")',
        'a:has-text("View uploaded orders")',
        'button:has-text("View orders")',
        'a:has-text("View orders")',
        'button:has-text("Continue")',
    ]
    
    clicked = await _safe_click(
        page,
        view_selectors,
        "View orders button",
        config.timeout_ms // 2,  # Shorter timeout, we have a fallback
        log,
    )
    
    if not clicked:
        # Fallback: Navigate to "Order confirmation" in left menu
        log("    View orders button not found, trying Order confirmation menu...")
        
        order_confirm_selectors = [
            'text="Order confirmation"',
            'a:has-text("Order confirmation")',
            'a:has-text("Order Confirmation")',
            '[href*="confirmation"]',
            '[href*="confirm"]',
            'nav a:has-text("confirmation")',
            '.sidebar a:has-text("confirmation")',
            '.menu a:has-text("confirmation")',
        ]
        
        clicked = await _safe_click(
            page,
            order_confirm_selectors,
            "Order confirmation menu",
            config.timeout_ms // 2,
            log,
        )
        
        if not clicked:
            return False, "Could not find 'View uploaded orders' button or 'Order confirmation' menu", 0
    
    await _wait_for_page_stable(page, config.timeout_ms // 2)
    await page.wait_for_timeout(config.inter_stage_delay_ms)
    
    # Check for portal errors
    has_error, error_msg = await _check_for_portal_error(page)
    if has_error:
        return False, f"Portal error when viewing orders: {error_msg[:200]}", 0
    
    # Now select the orders
    return await _select_orders_on_page(page, po_number, config, log)


async def _navigate_to_order_confirmation(
    page,
    config: SpringPortalConfig,
    log: Callable,
) -> Tuple[bool, str]:
    """
    Navigate directly to Order confirmation page via left menu.
    Used as fallback when View uploaded orders fails.
    
    Returns:
        (success: bool, error_message: str)
    """
    log("  Navigating to Order confirmation...")
    
    order_confirm_selectors = [
        'text="Order confirmation"',
        'a:has-text("Order confirmation")',
        'a:has-text("Order Confirmation")',
        '[href*="confirmation"]',
        '[href*="confirm"]',
        'nav a:has-text("confirmation")',
        '.sidebar a:has-text("confirmation")',
        '.menu a:has-text("confirmation")',
    ]
    
    clicked = await _safe_click(
        page,
        order_confirm_selectors,
        "Order confirmation menu",
        config.timeout_ms,
        log,
    )
    
    if not clicked:
        return False, "Could not find Order confirmation menu"
    
    await _wait_for_page_stable(page, config.timeout_ms // 2)
    await page.wait_for_timeout(config.inter_stage_delay_ms)
    
    return True, ""


async def _select_orders_on_page(
    page,
    po_number: str,
    config: SpringPortalConfig,
    log: Callable,
) -> Tuple[bool, str, int]:
    """
    Select orders on the current page (either View orders or Order confirmation page).
    
    Returns:
        (success: bool, error_message: str, orders_selected: int)
    """
    # Find and select ALL orders by PO number
    if not po_number:
        log("    ⚠ No PO number provided, cannot select specific order")
        return True, "", 0  # Partial success - upload worked
    
    log(f"    Looking for orders with PO: {po_number}")
    
    # Get today's date in the format shown in the portal (DD-MM-YYYY)
    today_str = datetime.now().strftime("%d-%m-%Y")
    log(f"    Filtering for orders created today: {today_str}")
    
    orders_selected = 0
    
    try:
        # Find ALL rows containing BOTH the PO number AND today's date
        # This ensures we only select the most recent upload, not old ones with same PO
        row_selectors = [
            f'tr:has-text("{po_number}"):has-text("{today_str}")',
            f'[role="row"]:has-text("{po_number}"):has-text("{today_str}")',
        ]
        
        for row_selector in row_selectors:
            try:
                # Get ALL matching rows, not just the first one
                matching_rows = page.locator(row_selector)
                row_count = await matching_rows.count()
                
                if row_count > 0:
                    # Cap at 2 orders max (one Standard, one Premium per upload)
                    max_to_select = min(row_count, 2)
                    log(f"    Found {row_count} order(s) matching PO {po_number}, selecting top {max_to_select}")
                    
                    # Click checkbox in each matching row (up to max)
                    for i in range(max_to_select):
                        try:
                            row = matching_rows.nth(i)
                            if await row.is_visible(timeout=2000):
                                # Find checkbox in this row
                                checkbox = row.locator('input[type="checkbox"], [role="checkbox"]').first
                                if await checkbox.is_visible(timeout=1000):
                                    # Check if already selected
                                    try:
                                        is_checked = await checkbox.is_checked()
                                    except Exception:
                                        is_checked = False
                                    
                                    if not is_checked:
                                        await checkbox.click()
                                        orders_selected += 1
                                        
                                        # Try to get the product type for logging
                                        try:
                                            row_text = await row.text_content()
                                            if 'STANDARD' in row_text.upper():
                                                log("    ✓ Selected: STANDARD MAIL SORTED (Economy)")
                                            elif 'PREMIUM' in row_text.upper():
                                                log("    ✓ Selected: PREMIUM MAIL SORTED (Priority)")
                                            else:
                                                log(f"    ✓ Selected order row {i + 1}")
                                        except Exception:
                                            log(f"    ✓ Selected order row {i + 1}")
                                        
                                        await page.wait_for_timeout(500)  # Brief pause between selections
                                    else:
                                        orders_selected += 1  # Already selected counts too
                        except Exception as row_err:
                            log(f"    ⚠ Could not select row {i + 1}: {row_err}")
                            continue
                    
                    if orders_selected > 0:
                        log(f"    ✓ Total orders selected: {orders_selected}")
                        await page.wait_for_timeout(1000)
                        return True, "", orders_selected
                    
            except Exception:
                continue
        
        # Fallback: if no orders found with today's date, try without date filter
        # This handles edge cases like portal showing different date format
        log("    No orders found with today's date, trying without date filter...")
        
        fallback_selectors = [
            f'tr:has-text("{po_number}")',
            f'[role="row"]:has-text("{po_number}")',
        ]
        
        for row_selector in fallback_selectors:
            try:
                matching_rows = page.locator(row_selector)
                row_count = await matching_rows.count()
                
                if row_count > 0:
                    # Only select the FIRST 2 rows (most recent Standard + Premium)
                    # Orders are typically shown newest first
                    max_to_select = min(row_count, 2)
                    log(f"    Found {row_count} order(s), selecting top {max_to_select}")
                    
                    for i in range(max_to_select):
                        try:
                            row = matching_rows.nth(i)
                            if await row.is_visible(timeout=2000):
                                checkbox = row.locator('input[type="checkbox"], [role="checkbox"]').first
                                if await checkbox.is_visible(timeout=1000):
                                    try:
                                        is_checked = await checkbox.is_checked()
                                    except Exception:
                                        is_checked = False
                                    
                                    if not is_checked:
                                        await checkbox.click()
                                        orders_selected += 1
                                        
                                        try:
                                            row_text = await row.text_content()
                                            if 'STANDARD' in row_text.upper():
                                                log("    ✓ Selected: STANDARD MAIL SORTED (Economy)")
                                            elif 'PREMIUM' in row_text.upper():
                                                log("    ✓ Selected: PREMIUM MAIL SORTED (Priority)")
                                            else:
                                                log(f"    ✓ Selected order row {i + 1}")
                                        except Exception:
                                            log(f"    ✓ Selected order row {i + 1}")
                                        
                                        await page.wait_for_timeout(500)
                                    else:
                                        orders_selected += 1
                        except Exception:
                            continue
                    
                    if orders_selected > 0:
                        log(f"    ✓ Total orders selected: {orders_selected}")
                        await page.wait_for_timeout(1000)
                        return True, "", orders_selected
                        
            except Exception:
                continue
        
        log(f"    ⚠ Could not find order rows for PO: {po_number}")
        return True, "", 0  # Partial success
        
    except Exception as e:
        log(f"    ⚠ Error selecting orders: {e}")
        return True, "", orders_selected  # Partial success


async def _hard_refresh_page(page, log: Callable, timeout_ms: int = 30000) -> bool:
    """
    Perform a hard refresh that mimics manually clicking the browser refresh button.

    This is more thorough than page.reload() because it:
    1. Uses JavaScript location.reload(true) for cache-bypassing reload
    2. Falls back to full URL navigation if that fails
    3. Waits for actual page elements to confirm the page is truly loaded

    Returns:
        True if refresh succeeded and page is ready, False otherwise
    """
    current_url = page.url

    # Strategy 1: JavaScript hard reload (cache-bypass)
    # This is closest to pressing F5/Ctrl+R in the browser
    try:
        log("    Performing hard refresh...")

        # Use JavaScript to force a true reload from server, not cache
        # location.reload(true) is deprecated but still works in most browsers
        # and forces a cache-bypassing reload
        await page.evaluate("() => { window.location.reload(true); }")

        # Wait for navigation to complete
        try:
            await page.wait_for_load_state("load", timeout=timeout_ms)
        except Exception:
            pass  # Continue even if this times out

        # Additional wait for dynamic content
        await page.wait_for_timeout(2000)

        # Verify page is actually loaded by checking for expected elements
        page_ready = await _verify_page_ready(page, timeout_ms // 2)
        if page_ready:
            log("    ✓ Hard refresh successful")
            return True

    except Exception as e:
        log(f"    ⚠ JavaScript reload failed: {e}")

    # Strategy 2: Full navigation to current URL (fresh request)
    # This completely tears down and rebuilds the page context
    try:
        log("    Trying full navigation refresh...")

        # Navigate away briefly to force complete teardown
        await page.goto("about:blank", wait_until="load", timeout=5000)
        await page.wait_for_timeout(500)

        # Navigate back to the portal
        await page.goto(current_url, wait_until="domcontentloaded", timeout=timeout_ms)

        # Wait for networkidle after DOM is loaded
        try:
            await page.wait_for_load_state("networkidle", timeout=timeout_ms // 2)
        except Exception:
            pass  # Continue even if networkidle times out

        await page.wait_for_timeout(2000)

        # Verify page is ready
        page_ready = await _verify_page_ready(page, timeout_ms // 2)
        if page_ready:
            log("    ✓ Navigation refresh successful")
            return True

    except Exception as e:
        log(f"    ⚠ Navigation refresh failed: {e}")

    # Strategy 3: Last resort - standard reload with extended wait
    try:
        log("    Trying standard reload as fallback...")
        await page.reload(wait_until="load", timeout=timeout_ms)
        await page.wait_for_timeout(3000)
        log("    ✓ Standard reload completed")
        return True
    except Exception as e:
        log(f"    ⚠ Standard reload failed: {e}")

    return False


async def _verify_page_ready(page, timeout_ms: int = 10000) -> bool:
    """
    Verify that the Spring portal page is actually loaded and ready.

    Checks for presence of expected UI elements that indicate
    the page has fully rendered and is interactive.
    """
    # Elements that should be present on a properly loaded portal page
    ready_indicators = [
        'table',  # Order table
        'tr',     # Table rows
        'button:has-text("Print")',
        'button:has-text("Download")',
        '[role="row"]',
        'input[type="checkbox"]',
    ]

    try:
        # Wait a moment for initial render
        await page.wait_for_timeout(1000)

        # Check if any of the expected elements are visible
        for selector in ready_indicators:
            try:
                element = page.locator(selector).first
                if await element.is_visible(timeout=timeout_ms // len(ready_indicators)):
                    return True
            except Exception:
                continue

        # Even if specific elements aren't found, check if body has content
        body_content = await page.locator("body").inner_text()
        if body_content and len(body_content.strip()) > 100:
            return True

    except Exception:
        pass

    return False


async def _dismiss_error_modal_and_refresh(page, log: Callable, timeout_ms: int = 30000) -> bool:
    """
    Dismiss the "unexpected error" modal and perform a hard refresh.

    The Spring portal shows this modal when you progress too fast.
    Solution: Click the X button, then do a proper hard refresh that
    mimics manually clicking the browser's refresh button.

    Returns:
        True if error was dismissed and page refreshed, False otherwise
    """
    # Look for the error modal close button (X)
    close_selectors = [
        'button:has(svg)',  # X button with SVG icon
        '.modal button:has-text("×")',
        '.modal button:has-text("X")',
        '[class*="close"]',
        '[aria-label="Close"]',
        '[aria-label="close"]',
        'button[class*="close"]',
        '.error button',
        'div:has-text("Error") button',
        # The blue X button visible in the screenshot
        'button:near(:text("Error"))',
    ]

    dismissed = False
    for selector in close_selectors:
        try:
            element = page.locator(selector).first
            if await element.is_visible(timeout=1000):
                await element.click()
                dismissed = True
                log("    ✓ Dismissed error modal")
                await page.wait_for_timeout(500)
                break
        except Exception:
            continue

    # If we couldn't find the close button, try pressing Escape
    if not dismissed:
        try:
            await page.keyboard.press("Escape")
            dismissed = True
            log("    ✓ Dismissed modal with Escape key")
            await page.wait_for_timeout(500)
        except Exception:
            pass

    if dismissed:
        # Perform a hard refresh that mimics browser refresh button
        return await _hard_refresh_page(page, log, timeout_ms)

    return False


async def _stage_download_pdf(
    page,
    po_number: str,
    output_dir: str,
    config: SpringPortalConfig,
    log: Callable,
) -> Tuple[bool, str, Optional[str]]:
    """
    Download the manifest PDF with retry logic.
    
    Handles the "unexpected error" modal that appears when progressing
    too fast through the portal - dismisses it and refreshes before retry.
    
    Returns:
        (success: bool, error_message: str, pdf_path: Optional[str])
    """
    for attempt in range(config.stage_retry_count + 1):
        if attempt > 0:
            log(f"    ⟳ Download retry {attempt}...")
            await page.wait_for_timeout(config.inter_stage_delay_ms)
        
        try:
            # Before clicking Print, check if orders are still selected
            # After a refresh, we may need to re-select them
            if attempt > 0 and po_number:
                log(f"    Re-selecting orders for PO {po_number}...")
                try:
                    row_selectors = [
                        f'tr:has-text("{po_number}")',
                        f'[role="row"]:has-text("{po_number}")',
                    ]
                    reselected_count = 0
                    for row_selector in row_selectors:
                        try:
                            matching_rows = page.locator(row_selector)
                            row_count = await matching_rows.count()
                            
                            for i in range(row_count):
                                try:
                                    row = matching_rows.nth(i)
                                    if await row.is_visible(timeout=2000):
                                        checkbox = row.locator('input[type="checkbox"], [role="checkbox"]').first
                                        if await checkbox.is_visible(timeout=1000):
                                            try:
                                                is_checked = await checkbox.is_checked()
                                            except Exception:
                                                is_checked = False
                                            
                                            if not is_checked:
                                                await checkbox.click()
                                                reselected_count += 1
                                                await page.wait_for_timeout(500)
                                except Exception:
                                    continue
                            
                            if reselected_count > 0:
                                log(f"    ✓ Re-selected {reselected_count} order(s)")
                                await page.wait_for_timeout(1000)
                                break
                        except Exception:
                            continue
                except Exception as e:
                    log(f"    ⚠ Could not re-select orders: {e}")
            
            # Add a delay before clicking Print to avoid "unexpected error"
            # The portal throws errors when you progress too fast
            if config.pre_print_delay_ms > 0:
                log(f"    Waiting {config.pre_print_delay_ms}ms before Print...")
                await page.wait_for_timeout(config.pre_print_delay_ms)
            
            # Set up download handler
            async with page.expect_download(timeout=config.timeout_ms) as download_info:
                # Click Print button
                print_clicked = await _safe_click(
                    page,
                    [
                        'button:has-text("Print")',
                        'a:has-text("Print")',
                        'button:has-text("Download")',
                        'a:has-text("Download")',
                        '[title="Print"]',
                        '[title="Download"]',
                    ],
                    "Print/Download button",
                    config.timeout_ms // 2,
                    log,
                )
                
                if not print_clicked:
                    continue
            
            download = await download_info.value
            
            # Save the downloaded file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Spring_{po_number}_{timestamp}.pdf" if po_number else f"Spring_manifest_{timestamp}.pdf"
            pdf_path = os.path.join(output_dir, filename)
            
            await download.save_as(pdf_path)
            log(f"    ✓ Downloaded: {filename}")
            
            return True, "", pdf_path
            
        except Exception as e:
            # Check for the "unexpected error" modal
            has_error, error_msg = await _check_for_portal_error(page)
            if has_error:
                log(f"    ⚠ Portal error: {error_msg[:80]}...")
                
                if attempt < config.stage_retry_count:
                    # Dismiss the error modal and refresh the page
                    dismissed = await _dismiss_error_modal_and_refresh(
                        page, log, config.timeout_ms
                    )
                    if dismissed:
                        log("    Will retry after refresh...")
                        continue
                    else:
                        log("    ⚠ Could not dismiss error modal")
                
                return False, f"Portal error during PDF download: {error_msg[:200]}", None
            
            if attempt == config.stage_retry_count:
                return False, f"PDF download failed: {str(e)}", None
    
    return False, "PDF download failed after retries", None


async def upload_to_spring_portal_robust(
    file_path: str,
    po_number: str = "",
    output_dir: str = "",
    auto_print: bool = True,
    config: Optional[SpringPortalConfig] = None,
    log_callback: Optional[Callable] = None,
) -> SpringPortalResult:
    """
    Robust Spring portal upload with comprehensive error handling.
    
    This function handles known portal issues:
    - Post-login hangs with multiple wait strategies
    - Missing UI elements with selector fallbacks
    - Portal "unexpected errors" with retry logic
    - PDF download failures with graceful degradation
    
    Args:
        file_path: Path to the manifest file to upload
        po_number: PO/Customer reference for order selection
        output_dir: Directory for downloaded PDF
        auto_print: Whether to print the downloaded PDF
        config: Portal configuration (uses defaults if None)
        log_callback: Function for logging messages
    
    Returns:
        SpringPortalResult with detailed status information
    """
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return SpringPortalResult(
            success=False,
            message="Playwright not installed. Run: pip install playwright && playwright install chromium",
            stage_reached=SpringPortalStage.INIT,
        )
    
    # Load credentials
    from core.credentials import get_spring_credentials
    creds = get_spring_credentials()
    if not creds.is_valid():
        return SpringPortalResult(
            success=False,
            message="Spring credentials not configured. Set SPRING_EMAIL and SPRING_PASSWORD in .env file.",
            stage_reached=SpringPortalStage.INIT,
        )
    
    config = config or SpringPortalConfig()
    output_dir = output_dir or os.path.dirname(file_path)
    
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    current_stage = SpringPortalStage.INIT
    
    try:
        async with async_playwright() as p:
            log("  Launching browser...")
            browser = await p.chromium.launch(headless=False)
            context = await browser.new_context()
            page = await context.new_page()
            
            try:
                # Navigate to login
                log("  Navigating to Spring portal...")
                await page.goto("https://my.spring-gds.com/", wait_until="domcontentloaded", timeout=config.timeout_ms)
                
                # Stage: Login
                current_stage = SpringPortalStage.LOGIN
                log("  Logging in...")
                success, error = await _stage_login(page, creds.email, creds.password, config, log)
                if not success:
                    await browser.close()
                    return SpringPortalResult(
                        success=False,
                        message=error,
                        stage_reached=current_stage,
                    )
                
                # Stage: Navigate to Upload
                current_stage = SpringPortalStage.FIND_UPLOAD
                log("  Navigating to upload page...")
                success, error = await _stage_navigate_to_upload(page, config, log)
                if not success:
                    await browser.close()
                    return SpringPortalResult(
                        success=False,
                        message=error,
                        stage_reached=current_stage,
                        requires_manual_intervention=True,
                    )
                
                # Stage: Upload File
                current_stage = SpringPortalStage.UPLOAD_FILE
                log(f"  Uploading file: {os.path.basename(file_path)}")
                success, error = await _stage_upload_file(page, file_path, config, log)
                if not success:
                    await browser.close()
                    return SpringPortalResult(
                        success=False,
                        message=error,
                        stage_reached=current_stage,
                    )
                
                # Stage: View and Select Order
                current_stage = SpringPortalStage.VIEW_ORDERS
                log("  Viewing uploaded orders...")
                success, error, orders_selected = await _stage_view_and_select_order(page, po_number, config, log)
                if not success:
                    await browser.close()
                    # Upload succeeded but couldn't navigate to orders
                    return SpringPortalResult(
                        success=True,  # Partial success
                        message=f"Upload completed but order selection failed: {error}",
                        stage_reached=current_stage,
                        pdf_downloaded=False,
                    )
                
                # Stage: Download PDF
                current_stage = SpringPortalStage.DOWNLOAD_PDF
                log("  Downloading manifest PDF...")
                success, error, pdf_path = await _stage_download_pdf(
                    page, po_number, output_dir, config, log
                )
                
                await browser.close()
                
                if not success:
                    return SpringPortalResult(
                        success=True,  # Upload worked, PDF failed
                        message=f"Upload completed but PDF download failed: {error}",
                        stage_reached=current_stage,
                        pdf_downloaded=False,
                    )
                
                # Print if requested
                if auto_print and pdf_path:
                    log("  Printing manifest...")
                    from gui import print_pdf_file
                    print_success, print_msg = print_pdf_file(pdf_path)
                    if print_success:
                        log(f"    ✓ {print_msg}")
                    else:
                        log(f"    ⚠ Print failed: {print_msg}")
                
                return SpringPortalResult(
                    success=True,
                    message="Upload, download and print completed successfully",
                    stage_reached=SpringPortalStage.COMPLETE,
                    pdf_downloaded=True,
                    pdf_path=pdf_path,
                )
                
            except Exception as e:
                try:
                    await browser.close()
                except Exception:
                    pass
                raise e
                
    except Exception as e:
        return SpringPortalResult(
            success=False,
            message=f"Unexpected error at {current_stage.value}: {str(e)}",
            stage_reached=current_stage,
        )


async def upload_with_full_retry(
    file_path: str,
    po_number: str = "",
    output_dir: str = "",
    auto_print: bool = True,
    max_retries: int = 2,
    log_callback: Optional[Callable] = None,
    config: Optional[SpringPortalConfig] = None,
) -> SpringPortalResult:
    """
    Upload with full workflow retry on critical failures.
    
    This wraps upload_to_spring_portal_robust with additional retry logic
    for complete workflow failures (not just stage failures).
    
    On retry after VIEW_ORDERS failure, skips re-upload and goes directly
    to Order confirmation page to select and print the already-uploaded orders.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    last_result = None
    
    for attempt in range(max_retries + 1):
        if attempt > 0:
            log(f"\n  ⟳ Full workflow retry {attempt} of {max_retries}...")
            await asyncio.sleep(2)  # Brief pause between retries
            
            # If the last failure was at VIEW_ORDERS stage, the upload already succeeded
            # Skip re-upload and go directly to Order confirmation
            if last_result and last_result.stage_reached == SpringPortalStage.VIEW_ORDERS:
                log("  Upload already completed, going to Order confirmation to select orders...")
                result = await _retry_via_order_confirmation(
                    po_number=po_number,
                    output_dir=output_dir,
                    auto_print=auto_print,
                    config=config,
                    log_callback=log_callback,
                )
                if result.success and result.pdf_downloaded:
                    return result
                # If that also failed, continue to next retry attempt
                last_result = result
                continue
        
        result = await upload_to_spring_portal_robust(
            file_path=file_path,
            po_number=po_number,
            output_dir=output_dir,
            auto_print=auto_print,
            config=config,
            log_callback=log_callback,
        )
        
        last_result = result
        
        if result.success and result.pdf_downloaded:
            return result
        
        # Decide whether to retry based on the failure stage
        if result.stage_reached in (
            SpringPortalStage.INIT,  # Credentials/setup issue - don't retry
            SpringPortalStage.COMPLETE,  # Should never happen
        ):
            break
        
        # For other stages, retry is worthwhile
        if attempt < max_retries:
            log(f"  Failed at {result.stage_reached.value}, will retry...")
    
    return last_result


async def _retry_via_order_confirmation(
    po_number: str,
    output_dir: str,
    auto_print: bool,
    config: Optional[SpringPortalConfig],
    log_callback: Optional[Callable],
) -> SpringPortalResult:
    """
    Retry by going directly to Order confirmation page.
    
    Used when upload succeeded but View uploaded orders failed.
    Logs in, navigates to Order confirmation, selects orders, and downloads PDF.
    """
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return SpringPortalResult(
            success=False,
            message="Playwright not installed",
            stage_reached=SpringPortalStage.INIT,
        )
    
    from core.credentials import get_spring_credentials
    creds = get_spring_credentials()
    if not creds.is_valid():
        return SpringPortalResult(
            success=False,
            message="Spring credentials not configured",
            stage_reached=SpringPortalStage.INIT,
        )
    
    config = config or SpringPortalConfig()
    
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    try:
        async with async_playwright() as p:
            log("  Launching browser...")
            browser = await p.chromium.launch(headless=False)
            context = await browser.new_context()
            page = await context.new_page()
            
            try:
                # Navigate and login
                log("  Navigating to Spring portal...")
                await page.goto("https://my.spring-gds.com/", wait_until="domcontentloaded", timeout=config.timeout_ms)
                
                log("  Logging in...")
                success, error = await _stage_login(page, creds.email, creds.password, config, log)
                if not success:
                    await browser.close()
                    return SpringPortalResult(
                        success=False,
                        message=f"Login failed: {error}",
                        stage_reached=SpringPortalStage.LOGIN,
                    )
                
                # Go directly to Order confirmation
                success, error = await _navigate_to_order_confirmation(page, config, log)
                if not success:
                    await browser.close()
                    return SpringPortalResult(
                        success=False,
                        message=f"Could not navigate to Order confirmation: {error}",
                        stage_reached=SpringPortalStage.VIEW_ORDERS,
                    )
                
                # Select orders
                log("  Selecting orders...")
                success, error, orders_selected = await _select_orders_on_page(page, po_number, config, log)
                if not success or orders_selected == 0:
                    await browser.close()
                    return SpringPortalResult(
                        success=True,  # Partial - upload worked earlier
                        message=f"Could not select orders: {error}",
                        stage_reached=SpringPortalStage.SELECT_ORDER,
                        pdf_downloaded=False,
                    )
                
                # Download PDF
                log("  Downloading manifest PDF...")
                success, error, pdf_path = await _stage_download_pdf(
                    page, po_number, output_dir, config, log
                )
                
                await browser.close()
                
                if not success:
                    return SpringPortalResult(
                        success=True,
                        message=f"Orders selected but PDF download failed: {error}",
                        stage_reached=SpringPortalStage.DOWNLOAD_PDF,
                        pdf_downloaded=False,
                    )
                
                # Print if requested
                if auto_print and pdf_path:
                    log("  Printing manifest...")
                    from gui import print_pdf_file
                    print_success, print_msg = print_pdf_file(pdf_path)
                    if print_success:
                        log(f"    ✓ {print_msg}")
                    else:
                        log(f"    ⚠ Print failed: {print_msg}")
                
                return SpringPortalResult(
                    success=True,
                    message="Order confirmation retry successful - PDF downloaded and printed",
                    stage_reached=SpringPortalStage.COMPLETE,
                    pdf_downloaded=True,
                    pdf_path=pdf_path,
                )
                
            except Exception as e:
                try:
                    await browser.close()
                except Exception:
                    pass
                raise e
                
    except Exception as e:
        return SpringPortalResult(
            success=False,
            message=f"Order confirmation retry failed: {str(e)}",
            stage_reached=SpringPortalStage.VIEW_ORDERS,
        )


def run_spring_upload_robust(
    file_path: str,
    po_number: str = "",
    output_dir: str = "",
    auto_print: bool = True,
    log_callback: Optional[Callable] = None,
) -> Tuple[bool, str, bool]:
    """
    Synchronous wrapper for the robust upload function.
    
    Maintains compatibility with existing calling code.
    
    Returns:
        (success: bool, message: str, pdf_downloaded: bool)
    """
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            from core.config import get_config
            app_config = get_config()
            
            # Build portal config from app settings
            portal_config = SpringPortalConfig(
                timeout_ms=app_config.portal_timeout_ms,
                retry_count=app_config.portal_retry_count,
                stage_retry_count=getattr(app_config, 'portal_stage_retry_count', 2),
            )
            
            result = loop.run_until_complete(
                upload_with_full_retry(
                    file_path=file_path,
                    po_number=po_number,
                    output_dir=output_dir,
                    auto_print=auto_print,
                    max_retries=app_config.portal_retry_count,
                    log_callback=log_callback,
                    config=portal_config,
                )
            )
            
            return result.success, result.message, result.pdf_downloaded
            
        finally:
            loop.close()
    except Exception as e:
        return False, f"Upload error: {str(e)}", False
