# Multi-Carrier Manifest Tool

Automated population of carrier manifests from internal carrier sheets, with integrated portal automation, printing, and file management.

**Version:** 1.2.4
**Author:** Finlay Crawley

## What's New in v1.2.x

**v1.2.4** - Auto-update on launch:
- The launcher now automatically pulls the latest changes from Git before starting
- Updates happen silently in the background - no manual `update.bat` required
- If update fails (no internet, etc.), the app launches normally anyway
- Update failures are logged to `update_errors.log` for troubleshooting

**v1.2.3** - Enhanced reliability and printing improvements:
- Spring portal now gracefully falls back between menu paths ("View uploaded orders" → "Order confirmation")
- Smart retry logic skips re-upload if file already uploaded successfully
- **Single-sided (simplex) printing** via SumatraPDF to reduce paper usage
- PDF print priority changed: SumatraPDF now preferred over Adobe
- Order selection limited to 2 most recent to avoid old order conflicts

**v1.2.2** - Multi-service level support:
- Spring portal now selects ALL matching orders (both Economy and Priority)
- Combined PDF manifest includes both STANDARD and PREMIUM service levels
- Enhanced logging for service type selection

**v1.2.0** - Robust portal automation:
- Comprehensive error handling for unreliable Spring portal behavior
- Per-stage retry logic with configurable retry counts
- Multiple wait strategies for post-login hangs
- Graceful degradation when PDF download fails
- Debug screenshots on portal failures

## Supported Carriers

| Carrier | Template File | Manifest Type | Portal Integration |
|---------|---------------|---------------|-------------------|
| Asendia UK Business Mail 2026 | `Asendia_UK_Business_2026_Mail_Manifest.xlsx` | Matrix (EU/ROW sections) | ❌ |
| Asendia UK Business Mail 2025 | `Asendia_UK_Business_Mail_2025.xlsx` | Matrix (EU/ROW sections) | ❌ |
| PostNord Business Mail | `PostNord.xlsx` | Matrix (Main Europe/ROW) | ❌ |
| Spring Global Delivery Solutions | `MailOrderTemplate.xlsx` | Flat order lines (CSV-style) | ✅ Auto-upload & PDF download |
| Landmark Global | `UploadCodeList_-_Citipost.xls` | CSV generation (pipe-separated) | ✅ Auto-upload |
| Deutsche Post | *(uses carrier sheet directly)* | Form-based portal submission | ✅ Auto-registration |
| Air Business | `Air_Business_Ireland.xlsx` | Fixed rows (Ireland only) | ❌ |
| Mail Americas/Africa | `Mail_America_Africa_2025.xlsx` | Weight-break based (3 sheets) | ❌ |
| United Business ADS | `United_Business.xlsx` | Weight-band based (single sheet) | ❌ |

## Features

### Core Functionality
- Auto-detection of carrier from cell B3 in carrier sheets
- Country code mapping with extensive variations handled
- Service type normalisation (Priority/Economy)
- Format classification (Letters/Flats/Packets)
- Configurable error threshold with detailed logging

### Portal Automation
- **Spring Global**: Robust automated upload with multi-order selection, combined PDF download, and printing
  - Handles both Economy (STANDARD) and Priority (PREMIUM) service levels in single manifest
  - Comprehensive error handling with per-stage retry logic
  - Multiple wait strategies for unreliable portal behavior
  - Graceful fallback between "View uploaded orders" and "Order confirmation" menu paths
  - Smart retry logic (skips re-upload if already successful)
- **Landmark Global**: Automatic CSV upload (separate files for Economy/Priority)
- **Deutsche Post**: Form-based manifest registration
- Configurable timeouts, retry counts (global and per-stage), and screenshot debugging

### Printing
- Auto-print manifests to network printer
- Three-tier PDF printing: SumatraPDF (simplex) → Adobe Acrobat → Windows Shell
- Single-sided (simplex) printing to reduce paper usage
- Fit-to-page printing for Excel workbooks
- Configurable PDF close delay

### Settings
- Printer selection from available Windows printers
- Portal timeout (5-120 seconds)
- Portal retry count (0-5 attempts for entire upload workflow)
- Portal stage retry count (0-5 attempts per individual stage, e.g., login, upload, print)
- PDF close delay (1-30 seconds)
- Max errors before stop (1-50)
- All settings persist to `config.json`

## Installation

### Requirements
- Python 3.8+
- Windows (for COM automation and printing)

### Dependencies
```bash
pip install openpyxl pandas pywin32 playwright python-dotenv
playwright install chromium
```

Alternatively, use the provided setup script:
```bash
first_time_setup.bat
```

### Setup
1. Clone or copy the project folder
2. Place carrier manifest templates in `templates/`
3. (Optional) Place `SumatraPDF.exe` in `tools/` for reliable PDF printing
4. (Optional) Create `.env` file with portal credentials for automated uploads:
   ```
   SPRING_EMAIL=your_email@example.com
   SPRING_PASSWORD=your_password
   LANDMARK_EMAIL=your_email@example.com
   LANDMARK_PASSWORD=your_password
   DEUTSCHEPOST_EMAIL=your_email@example.com
   DEUTSCHEPOST_PASSWORD=your_password
   DEUTSCHEPOST_CONTACT=Your Name
   ```
   See `.env.example` for template.
5. Run `python gui.py` or double-click `Run Manifest Tool.vbs` (auto-updates from Git)

### PDF Printing (For Deployment)

The tool uses this priority for PDF printing:

| Priority | Method | Notes |
|----------|--------|-------|
| 1 | SumatraPDF | Fast, silent, portable, **single-sided printing** (place in `tools/`) |
| 2 | Adobe Acrobat/Reader | Best quality (if installed) |
| 3 | Windows Shell | Last resort (may show dialogs) |

**Recommended Setup:** Download [SumatraPDF Portable](https://www.sumatrapdfreader.org/download-free-pdf-viewer) and place `SumatraPDF.exe` in the `tools/` folder. SumatraPDF is now the preferred method as it supports single-sided (simplex) printing, reducing paper usage.

## Usage

1. **Select Carrier Sheet**: Browse to your internal carrier sheet (`.xlsx`)
2. **Select Output Folder**: Choose where to save populated manifests
3. **Click Process**: The tool will:
   - Detect the carrier from cell B3
   - Load the appropriate template
   - Populate data into correct cells/generate CSV
   - Save output files with timestamp
   - (Optional) Auto-print and upload to portal

### Quick Actions
- **Print Last Manifest**: Re-print the most recent output file
- **Upload to Portal**: Manually trigger portal upload for the last processed manifest

## Carrier Sheet Format

The internal carrier sheet must have:

| Cell | Content |
|------|---------|
| B3 | Carrier Name (e.g., "Asendia 2026", "PostNord", "Spring") |
| B4 | PO Number |

Data starts at row 9 with headers at row 8:

| Column | Header |
|--------|--------|
| A | Country |
| C | Service |
| D | Format |
| E | Items |
| F | Weight (KG) |

### Service Types
- `Untracked Priority` / `Untracked Priority Mail` → Priority
- `Untracked Economy` / `Untracked Economy Mail` → Economy

### Format Types
- `Letters`
- `Flats` (Large Letters / Boxable)
- `Packets` (Non-boxable)

## Carrier Details

### Asendia (2025/2026)
- **Structure**: Matrix layout with EU section (format columns) and ROW section (left/right split)
- **Sheets**: Priority Manifest, Non-Priority Manifest
- **Countries**: Full international coverage with extensive name mapping

### PostNord
- **Structure**: Matrix layout with Main Europe (Boxable/Nonboxable columns) and ROW section
- **Sheets**: Priority Manifest, Non-Priority Manifest
- **Countries**: European focus with Scandinavian regional groupings

### Spring Global Delivery Solutions
- **Structure**: Flat order lines grouped by product code (1MI/2MI)
- **Output**: Excel file uploaded to Spring portal
- **Portal**: my.spring-gds.com - robust automated workflow
  - Uploads Excel file to portal
  - Selects ALL matching orders (both STANDARD and PREMIUM service levels)
  - Downloads combined PDF manifest containing both Economy and Priority shipments
  - Automatically prints single-sided manifest
  - Smart retry logic with per-stage error handling
  - Graceful fallback between menu paths if portal layout changes
- **Codes**: EU uses B/L/N format codes; ROW uses P/G/E codes

### Landmark Global
- **Structure**: CSV generation (pipe-separated)
- **Output**: Separate files for Economy (12SL03) and Priority (12SL02)
- **Portal**: bpost business portal - automatic upload
- **Headers**: `CONTRACT_NR|PRODUCT_CODE|DEPOSIT_DATE|DEPOSIT_DAY_PART||PO|`

### Deutsche Post
- **Structure**: Uses carrier sheet directly (removes EMB Manifest sheet)
- **Portal**: packet.deutschepost.com - form-based registration
- **Fields**: Contact name, Job reference (PO), Item format, Total weight

### Air Business
- **Structure**: Fixed rows for Ireland only
- **Rows**: Letters (16), Flats (19), Packets (22)
- **Coverage**: Ireland/Eire/IE only

### Mail Americas/Africa
- **Structure**: Weight-break based with three regional sheets
- **Sheets**: Mail Africa 2025, Mail Americas 2025, Europe & ROW 2025
- **Europe & ROW**: Uses format columns instead of weight breaks

### United Business ADS
- **Structure**: Single sheet with weight-band rows for select countries
- **Service**: Untracked Economy Mail only (all records → Economy)
- **Formats**: Letters, Flats, Packets (columns C-H)
- **Weight Bands**: China (5 bands), Russia (2), Ukraine (2) - uses average item weight
- **Coverage**: Eastern Europe and Asia regions

## Adding New Carriers

1. Create new module in `carriers/` (use existing carrier as template)
2. Inherit from `BaseCarrier` and implement required methods:
   - `build_country_index()` - Map countries to sheet/row/column positions
   - `get_cell_positions()` - Get items/weight columns for format
   - `set_metadata()` - Set PO number and date
   - `place_record()` - Override for custom placement logic
3. Add country name mappings for any variations
4. Register in `carriers/__init__.py`:
   - Import the class
   - Add to `CARRIER_REGISTRY` dict
   - Add detection logic in `get_carrier()` function
5. Add template to `templates/` folder (if applicable)

## Configuration

Settings are stored in `config.json`:

```json
{
  "printer_name": "\\\\print01.citipost.co.uk\\KT02",
  "portal_timeout_ms": 30000,
  "portal_retry_count": 2,
  "portal_stage_retry_count": 2,
  "pdf_close_delay_seconds": 7,
  "max_errors_before_stop": 5,
  "default_output_dir": "U:\\Erith\\Hailey Road\\International Ops\\Pre-Alerts\\Dispatch #1"
}
```

### Default Paths
- **Output Directory**: `U:\Erith\Hailey Road\International Ops\Pre-Alerts\Dispatch #1`
- **Printer**: `\\print01.citipost.co.uk\KT02`

## Error Handling

- Processing continues until error threshold is reached
- Unmapped countries are logged as errors
- Output files are generated with successful records even if some fail
- Portal failures include screenshot capture for debugging

## Building Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --add-data "templates:templates" --add-data "tools:tools" gui.py
```

The executable will be in `dist/gui.exe`.

## Project Structure

```
Multi Carrier Manifest Automation/
├── carriers/
│   ├── __init__.py          # Carrier registry and auto-detection
│   ├── base.py              # Abstract base class
│   ├── asendia.py           # Asendia 2025/2026 handler
│   ├── postnord.py          # PostNord handler
│   ├── spring.py            # Spring Global handler (order lines)
│   ├── spring_portal.py     # Spring Global robust portal automation
│   ├── landmark.py          # Landmark Global handler (CSV generation)
│   ├── deutschepost.py      # Deutsche Post handler
│   ├── deutschepost_portal.py  # Deutsche Post portal automation
│   ├── airbusiness.py       # Air Business Ireland handler
│   ├── mail_americas.py     # Mail Americas/Africa handler
│   └── unitedbusiness.py    # United Business ADS handler
├── core/
│   ├── __init__.py
│   ├── engine.py            # Processing engine
│   ├── config.py            # Configuration management
│   └── credentials.py       # Portal credential loading (.env)
├── templates/
│   ├── Asendia_UK_Business_2026_Mail_Manifest.xlsx
│   ├── Asendia_UK_Business_Mail_2025.xlsx
│   ├── PostNord.xlsx
│   ├── MailOrderTemplate.xlsx
│   ├── Air_Business_Ireland.xlsx
│   ├── Mail_America_Africa_2025.xlsx
│   ├── United_Business.xlsx
│   └── UploadCodeList_-_Citipost.xls
├── tools/
│   ├── SumatraPDF.exe       # PDF printer (download separately)
│   └── README.md            # Setup instructions
├── gui.py                   # Main application
├── config.json              # User settings (auto-generated)
├── Run Manifest Tool.vbs    # Quick launcher (auto-updates from Git)
├── update.bat               # Manual update script (git pull)
├── first_time_setup.bat     # Initial setup script
├── CHANGELOG.md             # Version history
└── README.md                # This file
```

## Troubleshooting

### Portal Timeouts
- Increase timeout in Settings dialog
- Check network connectivity
- Screenshots saved to output directory on failure

### PDF Printing Issues
- Ensure SumatraPDF.exe is in `tools/` folder
- Check printer is online and accessible
- Try increasing PDF close delay in Settings

### Country Not Found
- Check spelling matches carrier sheet exactly
- Add mapping to carrier's `country_mapping` dict
- Check for regional groupings (e.g., "Rest of Europe")

### Carrier Not Detected
- Ensure cell B3 contains carrier name
- Check detection logic in `carriers/__init__.py`
- Common patterns: "Asendia 2026", "PostNord", "Spring"

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history.
