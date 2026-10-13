@echo off
cd /d "%~dp0"
echo Starting DocPilot on http://localhost:8930 ...
start "" http://localhost:8930
python -m uvicorn backend.app:app --host 127.0.0.1 --port 8930
