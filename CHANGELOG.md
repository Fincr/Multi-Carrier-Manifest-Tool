# Changelog

All notable changes to the Multi-Carrier Manifest Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
