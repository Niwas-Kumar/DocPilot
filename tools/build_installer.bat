@echo off
REM Double-click friendly wrapper around tools\build_installer.py
cd /d "%~dp0\.."
python tools\build_installer.py %*
pause
