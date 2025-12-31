# Changelog

All notable changes to the Multi-Carrier Manifest Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
