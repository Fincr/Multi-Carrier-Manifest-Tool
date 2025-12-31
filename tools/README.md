# Tools Folder

This folder contains portable tools bundled with the Manifest Tool for deployment.

## SumatraPDF (Recommended for PDF Printing)

**SumatraPDF** is used as the fallback PDF printer when Adobe Acrobat/Reader is not installed.

### Why SumatraPDF?
- **Free and open source** - No licensing costs
- **Portable** - Single .exe file, no installation required
- **Silent printing** - No dialogs or user interaction needed
- **Fast** - Lightweight and quick to launch
- **Reliable** - Good barcode and layout rendering for carrier manifests

### Download

Download **SumatraPDF Portable** (64-bit recommended):

1. Go to: https://www.sumatrapdfreader.org/download-free-pdf-viewer
2. Download: **SumatraPDF-x.x.x-64.zip** (Portable version)
3. Extract `SumatraPDF.exe` to this folder

Or direct link (check for latest version):
https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64.zip

### File Location

Place the executable here:
```
Multi Carrier Manifest Automation/
├── tools/
│   ├── SumatraPDF.exe    <-- Place here
│   └── README.md
```

### Printing Priority

The manifest tool tries PDF printers in this order:

1. **Adobe Acrobat/Reader** - Best quality, used if installed
2. **SumatraPDF** - Fast and silent fallback (this folder)
3. **Windows Shell** - Last resort, may show print dialogs

### Verify Installation

After placing `SumatraPDF.exe` in this folder, the log output will show:
```
✓ Sent to printer via SumatraPDF
```

Instead of:
```
✓ Sent to default printer via Windows shell (may require interaction)
```
