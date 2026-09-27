@echo off
REM Builds a single-file Windows .exe. Run from this folder in a terminal with Python 3.10+ installed.
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconsole --onefile --name WABulkSender app.py
echo.
echo Done. Your app is in dist\WABulkSender.exe
