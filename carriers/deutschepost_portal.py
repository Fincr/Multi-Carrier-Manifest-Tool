"""
Deutsche Post portal automation.

Handles the Deutsche Post business portal workflow:
1. Login to portal
2. Click Ship in navigation
3. Click "Prepare Airway Bills" button
4. Click "Print Airway Bill" button
5. Fill form with manifest details
6. Click "Create" and download manifest PDF
"""

import os
import asyncio
from datetime import datetime
from typing import Callable, Optional

from core.credentials import get_deutschepost_credentials


async def _upload_to_deutschepost_portal_impl(
    po_number: str,
    total_weight: float,
    item_format: str,
    output_dir: str,
    auto_print: bool = True,
    log_callback: Optional[Callable[[str], None]] = None,
    timeout_ms: int = 30000
) -> tuple[bool, str]:
    """
    Internal implementation of Deutsche Post portal registration.
    
    Args:
        item_format: One of 'P', 'G', 'E', or 'mixed (P/G/E)'
    """
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return False, "Playwright not installed. Run: pip install playwright && playwright install chromium"
    
    # Load credentials from environment/.env
    creds = get_deutschepost_credentials()
    if not creds.is_valid():
        return False, "Deutsche Post credentials not configured. Set DEUTSCHEPOST_EMAIL and DEUTSCHEPOST_PASSWORD in .env file."
    
    EMAIL = creds.email
    PASSWORD = creds.password
    CONTACT_NAME = creds.contact_name
    LOGIN_URL = "https://packet.deutschepost.com/webapp/index.xhtml"
    
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    downloaded_file = None
    
    try:
        async with async_playwright() as p:
            # Launch browser
            log("  Launching browser...")
            browser = await p.chromium.launch(headless=False)
            context = await browser.new_context(accept_downloads=True)
            page = await context.new_page()
            
            try:
                # Navigate to login page
                log("  Navigating to Deutsche Post portal...")
                await page.goto(LOGIN_URL, wait_until="networkidle", timeout=timeout_ms)
                
                # Handle cookie consent banner if present
                log("  Checking for cookie consent...")
                cookie_selectors = [
                    'button:has-text("Accept")',
                    'button:has-text("Accept all")',
                    'button:has-text("Alle akzeptieren")',
                    '#onetrust-accept-btn-handler',
                ]
                
                for selector in cookie_selectors:
                    try:
                        element = page.locator(selector).first
                        if await element.is_visible(timeout=2000):
                            await element.click()
                            log("    ✓ Cookie consent accepted")
                            await page.wait_for_timeout(1000)
                            break
                    except Exception:
                        continue
                
                await page.wait_for_timeout(2000)
                
                # Enter email
                log("  Entering credentials...")
                email_selectors = [
                    'input[type="email"]',
                    'input[name="email"]',
                    'input[id*="email"]',
                    'input[name="username"]',
                ]
                
                for selector in email_selectors:
                    try:
                        element = page.locator(selector).first
                        if await element.is_visible(timeout=2000):
                            await element.fill(EMAIL)
                            log("    ✓ Email entered")
                            break
                    except Exception:
                        continue
                
                # Enter password
                try:
                    password_field = page.locator('input[type="password"]').first
                    if await password_field.is_visible(timeout=2000):
                        await password_field.fill(PASSWORD)
                        log("    ✓ Password entered")
                except Exception:
                    pass
                
                # Click Login button
                log("  Clicking Login...")
                login_selectors = [
                    'button:has-text("Login")',
                    'button:has-text("Log in")',
                    'input[type="submit"]',
                    'button[type="submit"]',
                ]
                
                clicked = False
                for selector in login_selectors:
                    try:
                        element = page.locator(selector).first
                        if await element.is_visible(timeout=2000):
                            await element.click()
                            clicked = True
                            break
                    except Exception:
                        continue
                
                if not clicked:
                    screenshot_path = os.path.join(output_dir, "dp_debug_login.png")
                    await page.screenshot(path=screenshot_path)
                    await browser.close()
                    return False, "Could not find Login button. Check dp_debug_login.png"
                
                await page.wait_for_load_state("networkidle", timeout=timeout_ms)
                await page.wait_for_timeout(3000)
                
                log("  ✓ Logged in successfully")
                
                # Step 1: Click "Ship" in navigation menu
                log("  Clicking Ship menu...")
                try:
                    ship_link = page.locator('a:has-text("Ship")').first
                    if await ship_link.is_visible(timeout=5000):
                        await ship_link.click()
                        log("    ✓ Clicked Ship")
                except Exception as e:
                    log(f"    Note: Ship menu click issue - {e}")
                
                await page.wait_for_load_state("networkidle", timeout=timeout_ms)
                await page.wait_for_timeout(2000)
                
                # Step 2: Click "Prepare Airway Bills" button
                log("  Clicking Prepare Airway Bills...")
                awb_button_selectors = [
                    'input[value="Prepare Airway Bills"]',
                    'button:has-text("Prepare Airway Bills")',
                    'a:has-text("Prepare Airway Bills")',
                ]
                
                clicked = False
                for selector in awb_button_selectors:
                    try:
                        elements = page.locator(selector)
                        count = await elements.count()
                        for i in range(count):
                            el = elements.nth(i)
                            if await el.is_visible(timeout=1000):
                                tag = await el.evaluate("el => el.tagName.toLowerCase()")
                                if tag in ['input', 'button', 'a']:
                                    await el.click()
                                    clicked = True
                                    log("    ✓ Clicked Prepare Airway Bills")
                                    break
                        if clicked:
                            break
                    except Exception:
                        continue

                if not clicked:
                    screenshot_path = os.path.join(output_dir, "dp_debug_prepare_awb.png")
                    await page.screenshot(path=screenshot_path)
                    await browser.close()
                    return False, "Could not find Prepare Airway Bills button. Check dp_debug_prepare_awb.png"
                
                await page.wait_for_load_state("networkidle", timeout=timeout_ms)
                await page.wait_for_timeout(2000)
                
                # Step 3: Click "Print Airway Bill" button
                log("  Clicking Print Airway Bill...")
                print_awb_selectors = [
                    'input[value="Print Airway Bill"]',
                    'button:has-text("Print Airway Bill")',
                    'a:has-text("Print Airway Bill")',
                ]
                
                clicked = False
                for selector in print_awb_selectors:
                    try:
                        elements = page.locator(selector)
                        count = await elements.count()
                        for i in range(count):
                            el = elements.nth(i)
                            if await el.is_visible(timeout=1000):
                                await el.click()
                                clicked = True
                                log("    ✓ Clicked Print Airway Bill")
                                break
                        if clicked:
                            break
                    except Exception:
                        continue

                if not clicked:
                    screenshot_path = os.path.join(output_dir, "dp_debug_print_awb.png")
                    await page.screenshot(path=screenshot_path)
                    await browser.close()
                    return False, "Could not find Print Airway Bill button. Check dp_debug_print_awb.png"
                
                await page.wait_for_load_state("networkidle", timeout=timeout_ms)
                await page.wait_for_timeout(2000)
                
                # Step 4: Fill in the form
                # The form is a table with rows. Each row has a label and an input.
                # We need to find inputs by their position/order since labels may not have proper associations
                log("  Filling in form...")
                
                # Get all visible text inputs in order
                all_text_inputs = page.locator('input[type="text"]:visible')
                input_count = await all_text_inputs.count()
                log(f"    Found {input_count} text inputs")
                
                # Based on the form structure:
                # Input 0: Contact name
                # Input 1: Your job reference
                # Input 2: Total weight in kg
                
                # Fill Contact name (first text input)
                log(f"    Contact name: {CONTACT_NAME}")
                try:
                    contact_input = all_text_inputs.nth(0)
                    await contact_input.click()
                    await page.wait_for_timeout(200)
                    await contact_input.fill(CONTACT_NAME)
                    await page.wait_for_timeout(200)
                    log("      ✓ Contact filled")
                except Exception as e:
                    log(f"      ⚠ Contact error: {e}")
                
                # Fill Job reference (second text input)
                log(f"    Job reference: {po_number}")
                try:
                    ref_input = all_text_inputs.nth(1)
                    await ref_input.click()
                    await page.wait_for_timeout(200)
                    await ref_input.fill(po_number)
                    await page.wait_for_timeout(200)
                    log("      ✓ Reference filled")
                except Exception as e:
                    log(f"      ⚠ Reference error: {e}")
                
                # Item format dropdown - options are P, G, E, mixed (P/G/E)
                log(f"    Item format: {item_format}")
                try:
                    # Find the Item Format select - it's the 3rd select (after Product and Service Level)
                    all_selects = page.locator('select:visible')
                    select_count = await all_selects.count()
                    log(f"      Found {select_count} dropdowns")
                    
                    # Item Format should be the 3rd select (index 2)
                    if select_count >= 3:
                        format_select = all_selects.nth(2)
                        await format_select.click()
                        await page.wait_for_timeout(200)
                        await format_select.select_option(label=item_format)
                        log(f"      ✓ Selected: {item_format}")
                except Exception as e:
                    log(f"      Format selection error: {e}")
                
                # Total weight (third text input, after contact and reference)
                log(f"    Total weight: {total_weight} kg")
                try:
                    weight_input = all_text_inputs.nth(2)
                    await weight_input.click()
                    await page.wait_for_timeout(200)
                    await weight_input.fill(str(total_weight))
                    await page.wait_for_timeout(200)
                    log("      ✓ Weight filled")
                except Exception as e:
                    log(f"      Weight error: {e}")
                
                await page.wait_for_timeout(1000)
                
                # Check for validation errors before clicking Create
                page_content = await page.content()
                if "please enter a customer reference" in page_content.lower():
                    log("  ⚠ Validation error: Job reference not filled")
                    await browser.close()
                    return False, "Failed to fill job reference field. Check dp_debug_form_filled.png"
                
                # Step 5: Click Create and wait for download
                log("  Clicking Create...")
                
                try:
                    async with page.expect_download(timeout=timeout_ms) as download_info:
                        create_button = page.locator('input[value="Create"], button:has-text("Create")').first
                        if await create_button.is_visible(timeout=3000):
                            await create_button.click()
                            log("    ✓ Clicked Create")
                        else:
                            screenshot_path = os.path.join(output_dir, "dp_debug_create.png")
                            await page.screenshot(path=screenshot_path)
                            await browser.close()
                            return False, "Could not find Create button. Check dp_debug_create.png"
                    
                    # Get the download
                    download = await download_info.value
                    
                    # Save the downloaded file
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    downloaded_filename = f"Deutsche_Post_{po_number}_{timestamp}.pdf"
                    downloaded_path = os.path.join(output_dir, downloaded_filename)
                    await download.save_as(downloaded_path)
                    downloaded_file = downloaded_path
                    log(f"  ✓ Downloaded: {downloaded_filename}")
                    
                except Exception as e:
                    log(f"  ⚠ Error after clicking Create: {e}")
                    
                    screenshot_path = os.path.join(output_dir, "dp_debug_after_create.png")
                    await page.screenshot(path=screenshot_path)
                    
                    page_content = await page.content()
                    if "please enter" in page_content.lower():
                        await browser.close()
                        return False, "Form validation failed - required field missing. Check dp_debug_after_create.png"
                    
                    await browser.close()
                    return False, f"Failed to download manifest. Check dp_debug_after_create.png. Error: {e}"
                
                await browser.close()
                
                # Print if enabled and file was downloaded
                if auto_print and downloaded_file:
                    log("  Printing downloaded manifest...")
                    import ctypes
                    import subprocess
                    import time
                    try:
                        result = ctypes.windll.shell32.ShellExecuteW(
                            None, "print", downloaded_file, None, None, 0
                        )
                        print_success = result > 32
                        print_msg = "Sent to default printer" if print_success else f"Print failed with code {result}"
                        
                        # Close Adobe after printing (7 second delay)
                        if print_success:
                            def close_adobe_later():
                                time.sleep(7)
                                try:
                                    subprocess.run(['taskkill', '/F', '/IM', 'Acrobat.exe'], 
                                                 capture_output=True, timeout=5)
                                    subprocess.run(['taskkill', '/F', '/IM', 'AcroRd32.exe'],
                                                 capture_output=True, timeout=5)
                                except Exception:
                                    pass
                            import threading
                            threading.Thread(target=close_adobe_later, daemon=True).start()
                            
                    except Exception as print_err:
                        print_success = False
                        print_msg = str(print_err)
                    if print_success:
                        log(f"  ✓ {print_msg}")
                        return True, "Manifest created, downloaded and printed successfully"
                    else:
                        log(f"  ⚠ {print_msg}")
                        return True, f"Manifest created and downloaded. Print failed: {print_msg}"
                elif downloaded_file:
                    return True, "Manifest created and downloaded successfully"
                else:
                    return False, "Failed to download manifest"
                
            except Exception as e:
                try:
                    await browser.close()
                except Exception:
                    pass
                raise e
                
    except Exception as e:
        return False, f"Portal automation failed: {str(e)}"


async def upload_to_deutschepost_portal(
    po_number: str,
    total_weight: float,
    item_format: str,
    output_dir: str,
    auto_print: bool = True,
    log_callback: Optional[Callable[[str], None]] = None,
    timeout_ms: int = 30000,
    retry_count: int = 1
) -> tuple[bool, str]:
    """
    Register a manifest on the Deutsche Post portal using Playwright.
    Includes automatic retry on timeout errors.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    last_error = None
    
    for attempt in range(retry_count + 1):
        if attempt > 0:
            log(f"\n  ⟳ Retry attempt {attempt} of {retry_count}...")
        
        success, message = await _upload_to_deutschepost_portal_impl(
            po_number, total_weight, item_format, output_dir, 
            auto_print, log_callback, timeout_ms
        )
        
        if success:
            return success, message
        
        last_error = message
        if "Timeout" in message or "timeout" in message:
            if attempt < retry_count:
                log("  ⚠ Timeout occurred, will retry...")
                continue
        else:
            break
    
    return False, last_error or "Portal automation failed after retries"


def run_deutschepost_upload(
    po_number: str,
    total_weight: float,
    item_format: str,
    output_dir: str,
    auto_print: bool = True,
    log_callback: Optional[Callable[[str], None]] = None
) -> tuple[bool, str]:
    """
    Synchronous wrapper for the async Deutsche Post portal function.
    """
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(
                upload_to_deutschepost_portal(
                    po_number, total_weight, item_format, output_dir,
                    auto_print, log_callback
                )
            )
        finally:
            loop.close()
    except Exception as e:
        return False, f"Upload error: {str(e)}"
