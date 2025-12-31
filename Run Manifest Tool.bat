@echo off
cd /d "C:\Users\fcrawley\OneDrive - Citipost Ltd\Tools and Automation\Automations\Manifest Automation\Multi-Carrier-Manifest-Tool"
python gui.py
if %errorlevel% neq 0 pause
