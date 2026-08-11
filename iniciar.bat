@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt -q
py app.py
pause
