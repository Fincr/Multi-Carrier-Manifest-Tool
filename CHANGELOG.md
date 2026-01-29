# Changelog

All notable changes to the Multi-Carrier Manifest Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.4.0] - 2026-01-23

### Added
- **Pre-Alert Email Automation** - New "Pre-Alerts" tab for sending manifest notification emails
  - Dedicated tab in the main UI alongside "Manifest Processing"
  - Automatically registers processed manifests for pre-alert carriers (Asendia, Deutsche Post, PostNord, United Business)
  - Configurable recipients (TO and CC) per carrier via GUI settings dialog
  - Customizable email subject templates with placeholders: {carrier}, {date}, {po_number}
  - Professional HTML email template with dispatch date and PO number
  - Send tracking prevents duplicate emails (remembers what was sent today)
  - "Already sent" warning with option to re-send if needed
  - Batch send with summary: select multiple manifests and send all at once
  - Detailed send log with success/failure status
  - Uses Outlook desktop client via COM automation (emails appear in Sent folder)
  - Configuration persists to `pre_alert_config.json`
  - Send history tracked in `pre_alert_log.json`

### Changed
- Main UI now uses tabbed interface (Notebook) for better organization
- Window layout updated: title bar at top, tabs below

### Technical Details
- New `pre_alerts/` module with:
  - `config_manager.py` - Carrier email configuration with JSON persistence
  - `email_sender.py` - Outlook COM integration for sending HTML emails
  - `send_tracker.py` - Daily send tracking with automatic cleanup
  - `pre_alert_tab.py` - Full Tkinter tab UI with configuration dialogs
- Email template stored in `templates/pre_alert_email.html`
- Carrier name matching handles variations (e.g., "Asendia 2026" → "Asendia")

## [1.3.1] - 2026-01-17

### Added
- **Batch Processing Improvements**
  - Duplicate carrier detection - prevents processing multiple sheets of the same carrier type in one batch
  - Jersey Post exclusion - files containing "Jersey Post" in carrier name are automatically skipped
  - Scrollable "Supported Carriers" tab in About dialog for better display on smaller screens

### Fixed
- **Upload file cleanup in batch mode** - Spring (.xlsx) and Landmark (.csv) upload files are now deleted after successful portal upload with PDF download, matching non-batch behavior
- Upload files are retained if PDF download fails, for debugging purposes

## [1.3.0] - 2026-01-17

### Added
- **Batch Processing Mode** - Process multiple carrier sheets from a folder in one operation
  - New "Batch Process Folder..." button in the GUI
  - Auto-detects valid carrier sheets in selected folder (Excel files with recognized carrier in B3)
  - Processes files sequentially, top to bottom (alphabetical order)
  - Shows scan preview with detected carriers before processing
  - Skips invalid files with clear error messages in log
  - Applies output folder, auto-print, and auto-upload settings to all files
  - Displays summary dialog with success/failure counts on completion
  - Full error handling - continues to next file if one fails

## [1.2.7] - 2026-01-17

### Added
- **United Business NZP ETOE Carrier Support**
  - New carrier for T&D Priority Manifest (Untracked Priority Mail)
  - Format columns: Letters (P), Flats (G), Packets (E)
  - Simple single-row-per-country structure (rows 6-50)
  - Template: `UBL_CP_Pre_Alert_T_D-ETOE.xlsx`
  - Auto-detects from B3 containing "NZP", "ETOE", or "T&D"
  - Country mapping for Czech Republic → Czechia, Taiwan → Taiwan, China

## [1.2.6] - 2026-01-17

### Fixed
- **Deutsche Post portal automation** - Product dropdown now correctly selects "Business Mail" instead of defaulting to "Packet"
  - Portal form has three dropdowns: Product, Service Level, and Item Format
  - Product was previously left at default "Packet" value
  - Now explicitly sets Product to "Business Mail" before form submission
  - Ensures manifests are created with correct service classification

## [1.2.5] - 2026-01-16

### Added
- **Landmark ISO country code mappings** for additional countries:
  - British Virgin Islands (VG)
  - Democratic Republic of Congo (CD)
  - Namibia (NA)
  - The Republic of Zambia / Zambia (ZM)

## [1.2.4] - 2026-01-12

### Added
- **Auto-update on launch** - The launcher (`Run Manifest Tool.vbs`) now automatically pulls the latest changes from Git before starting the application
  - Updates happen silently in the background
  - If update fails (no internet, Git not available, etc.), the app still launches normally
  - Update failures are logged to `update_errors.log` for troubleshooting
  - No more need to manually run `update.bat` - just double-click to run and you're always up to date

## [1.2.3] - 2026-01-08

### Fixed
- **Spring portal: "View uploaded orders" fallback** - If button not found after upload, now automatically falls back to "Order confirmation" menu
- **Spring portal: Smart retry logic** - On retry after VIEW_ORDERS failure, skips re-upload and goes directly to Order confirmation page to select already-uploaded orders
- **PDF printing now single-sided** - SumatraPDF is now tried first with `-print-settings simplex` to force single-sided printing
- **Max 2 orders selected** - Caps order selection to 2 (one Standard, one Premium) to avoid selecting old orders with same PO

### Changed
- PDF print priority changed: SumatraPDF (with simplex) → Adobe → Windows shell
- File validation wait increased from 4 to 7 seconds for slower portal responses

## [1.2.2] - 2026-01-08

### Fixed
- **Spring portal now selects ALL orders matching the PO number** before printing
  - Previously only selected one order, missing either Economy (STANDARD MAIL SORTED) or Priority (PREMIUM MAIL SORTED)
  - Now correctly identifies and ticks all rows with the same Customer ref (PO number)
  - Downloads a single combined PDF manifest containing both service levels
- Enhanced logging shows which product types are being selected (STANDARD vs PREMIUM)

## [1.2.1] - 2026-01-08

### Added
- Enhanced diagnostic logging for Landmark and Spring processing to help debug multi-service-level issues
- Detailed file collection logging showing exactly which files are queued for upload/print
- Logging now shows Priority vs Economy breakdown for Spring order lines
- Logging now shows primary file vs additional files split for Landmark CSVs

## [1.2.0] - 2026-01-08

### Added
- **Robust Spring Portal Automation** (`carriers/spring_portal.py`)
  - Comprehensive error handling for unreliable portal behaviour
  - Per-stage retry logic with configurable retry counts
  - Multiple wait strategies for post-login hangs
  - Graceful degradation when PDF download fails (upload still succeeds)
  - Detailed stage-based error reporting
  - Debug screenshots on failures

### Changed
- Spring portal now uses new robust automation module instead of inline implementation
- Increased default retry count from 1 to 2 for better reliability
- Added `portal_stage_retry_count` config option (default: 2) for per-stage retries

### Fixed
- Post-login hang: Multiple wait strategies instead of single networkidle wait
- "Upload Multiple Orders" button not found: Extended selector fallbacks and page refresh retry
- "Unexpected error" after CSV upload: Portal error detection with retry logic
- "Unexpected error" when clicking Print: Dismisses error modal, refreshes page, re-selects order, and retries
- Added 4-second delays before Print and after CSV upload to reduce "unexpected error" occurrences (portal dislikes fast progression)

### Technical Details
- New `SpringPortalStage` enum tracks workflow stages (LOGIN, FIND_UPLOAD, UPLOAD_FILE, etc.)
- New `SpringPortalResult` dataclass provides detailed status including partial success states
- New `SpringPortalConfig` dataclass consolidates portal configuration
- Modular stage functions (`_stage_login`, `_stage_upload_file`, etc.) with individual retry logic
- `_wait_for_page_stable()` helper for more reliable page load detection
- `_safe_click()` helper with multiple selector fallbacks
- `_check_for_portal_error()` helper to detect portal error states

## [1.1.1] - 2026-01-02

### Added
- Zero record detection: Processing now cancels gracefully if carrier sheet contains no data rows, with clear warning message to user

### Fixed
- Czech Republic country mapping for Asendia 2026 template (was incorrectly mapping to "Czechia" when template uses "Czech Republic")

## [1.1.0] - 2025-12-31

### Added
- **United Business ADS Carrier Support**
  - Single service type: Untracked Economy Mail
  - Weight-band based row selection for China (5 bands), Russia (2), Ukraine (2)
  - Format columns: Letters, Flats, Packets
  - Country name mapping for manifest typos (Afganistan, Azerbajan, Kyrgystan)
  - Template: `United_Business.xlsx`

### Fixed
- Batch launcher pointing to incorrect project directory

## [1.0.0] - 2024-12-30

### Added
- **Carrier Support**
  - Asendia UK Business Mail (2025 and 2026 templates)
  - PostNord Business Mail
  - Spring Global Delivery Solutions (with portal automation)
  - Landmark Global (with portal automation)
  - Deutsche Post (with portal automation)
  - Air Business (Ireland-only)
  - Mail Americas/Africa

- **Portal Automation**
  - Automatic upload to Spring portal with manifest download and printing
  - Automatic upload to Landmark portal (Economy and Priority CSV files)
  - Automatic registration on Deutsche Post portal
  - Configurable timeout and retry settings
  - Screenshot capture for debugging failed uploads

- **Printing**
  - Auto-print manifests to configured network printer
  - Support for Adobe Acrobat, SumatraPDF, and Windows shell printing
  - Fit-to-page printing for Excel workbooks
  - Configurable PDF close delay

- **Settings Dialog**
  - Printer selection dropdown (enumerates available Windows printers)
  - Portal timeout configuration (5-120 seconds)
  - Portal retry count (0-5 retries)
  - PDF close delay (1-30 seconds)
  - Max errors before stop (1-50)
  - Settings persist to config.json

- **User Interface**
  - Process Manifest button with progress indicator
  - Print Last Manifest button
  - Upload to Portal button (for manual upload)
  - Processing log with detailed status messages
  - Auto-print and auto-upload checkboxes (enabled by default)

- **Core Features**
  - Auto-detection of carrier from cell B3 in carrier sheets
  - Country code mapping and validation
  - Service type and format classification
  - Error handling with configurable max errors threshold
  - Output files saved to configurable directory

### Notes
- Default output directory: `U:\Erith\Hailey Road\International Ops\Pre-Alerts\Dispatch #1`
- Default printer: `\\print01.citipost.co.uk\KT02`
- Requires Python with openpyxl, pandas, playwright, and pywin32
