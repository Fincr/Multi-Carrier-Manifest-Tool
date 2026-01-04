@echo off
REM ============================================
REM First Time Setup - Run this ONCE
REM ============================================
REM Installs Python dependencies and Playwright
REM ============================================

echo ============================================
echo Manifest Tool - First Time Setup
echo ============================================
echo.

REM Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo IMPORTANT: Tick "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

REM Check Git is installed
git --version >nul 2>&1
if errorlevel 1 (
    echo WARNING: Git is not installed.
    echo You won't be able to receive updates automatically.
    echo Install from https://git-scm.com/downloads if needed.
    echo.
)

echo [1/2] Installing Python dependencies...
pip install openpyxl pandas playwright pywin32 python-dotenv
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)
echo.

echo [2/2] Installing Playwright Chromium browser...
echo This downloads ~150MB, please wait...
playwright install chromium
if errorlevel 1 (
    echo ERROR: Failed to install Playwright.
    pause
    exit /b 1
)
echo.

echo ============================================
echo Setup complete!
echo ============================================
echo.
echo Next steps:
echo 1. Rename ".env.example" to ".env"
echo 2. Edit ".env" and add your portal credentials
echo 3. Double-click "run.bat" to start the tool
echo.
pause
