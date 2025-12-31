# Multi-Carrier-Manifest-Tool
Python tool automating carrier manifest population and portal submission for Citipost's international operations. Supports 7+ carriers (Asendia, PostNord, Spring, Deutsche Post, etc.) with plugin architecture. Features automated portal workflows, PDF printing, robust error handling, and configurable settings for high-volume shipping environments.

## Supported Carriers

| Carrier | Template File |
|---------|---------------|
| Asendia UK Business Mail 2026 | `Asendia_UK_Business_2026_Mail_Manifest.xlsx` |
| PostNord Business Mail | `PostNord.xlsx` |

## Installation

### Requirements
- Python 3.8+
- Required packages: `openpyxl`, `pandas`, `pywin32`, `playwright`

```bash
pip install openpyxl pandas pywin32 playwright
playwright install chromium
```

### Setup
1. Place carrier manifest templates in the `templates/` folder
2. Run `python gui.py`

### PDF Printing (For Deployment)

The tool prints PDFs using this priority:
1. **Adobe Acrobat/Reader** - Best quality (if installed)
2. **SumatraPDF** - Fast, silent, portable fallback
3. **Windows Shell** - Last resort (may show dialogs)

**For machines without Adobe:** Download [SumatraPDF Portable](https://www.sumatrapdfreader.org/download-free-pdf-viewer) and place `SumatraPDF.exe` in the `tools/` folder. See `tools/README.md` for details.

## Usage

1. **Select Carrier Sheet**: Browse to your internal carrier sheet (.xlsx)
2. **Select Output Folder**: Choose where to save populated manifests
3. **Click Process**: The tool will:
   - Detect the carrier from cell B3
   - Load the appropriate template
   - Populate data into correct cells
   - Save with filename: `{Carrier}_{PO}_{timestamp}.xlsx`

## Carrier Sheet Format

The internal carrier sheet must have:

| Cell | Content |
|------|---------|
| B3 | Carrier Name (e.g., "Asendia 2026", "PostNord") |
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
- `Flats` (maps to Boxable for Main Europe in PostNord)
- `Packets` (maps to Nonboxable for Main Europe in PostNord)

## Adding New Carriers

1. Create new module in `carriers/` (copy `postnord.py` as template)
2. Implement required methods:
   - `build_country_index()` - Map countries to sheet/row/column positions
   - `get_cell_positions()` - Get items/weight columns for format
   - `set_metadata()` - Set PO number and date
3. Add country mapping for any naming differences
4. Register in `carriers/__init__.py`
5. Add template to `templates/` folder

## Error Handling

- Processing stops after 5 errors per carrier
- Unmapped countries are logged as errors
- Output file is still generated with successful records

## Building Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --add-data "templates:templates" gui.py
```

The executable will be in `dist/gui.exe`.

## Project Structure

```
manifest_tool/
├── carriers/
│   ├── __init__.py      # Carrier registry
│   ├── base.py          # Abstract base class
│   ├── asendia.py       # Asendia handler
│   ├── postnord.py      # PostNord handler
│   ├── spring.py        # Spring Global handler
│   ├── landmark.py      # Landmark Global handler
│   ├── deutschepost.py  # Deutsche Post handler
│   ├── airbusiness.py   # Air Business handler
│   └── mailamericas.py  # Mail Americas handler
├── core/
│   ├── __init__.py
│   └── engine.py        # Processing engine
├── templates/           # Manifest templates
│   ├── Asendia_UK_Business_2026_Mail_Manifest.xlsx
│   ├── PostNord.xlsx
│   ├── MailOrderTemplate.xlsx
│   └── ...
├── tools/               # Portable tools for deployment
│   ├── SumatraPDF.exe   # PDF printer (download separately)
│   └── README.md        # Setup instructions
├── gui.py               # Main application
├── Run Manifest Tool.bat
└── README.md
```
