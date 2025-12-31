@echo off
cd /d "C:\Users\fcrawley\OneDrive - Citipost Ltd\Projects\Active\Manifest Automation\Multi Carrier Manifest Automation"
python gui.py
if %errorlevel% neq 0 pause
