@echo off
REM ============================================
REM Update Manifest Tool
REM ============================================
REM Pulls the latest changes from GitHub
REM ============================================

cd /d "%~dp0"

echo ============================================
echo Updating Manifest Tool...
echo ============================================
echo.

git pull
if errorlevel 1 (
    echo.
    echo ERROR: Update failed. 
    echo Make sure you have internet access and Git installed.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================
echo Update complete!
echo ============================================
echo.
pause
