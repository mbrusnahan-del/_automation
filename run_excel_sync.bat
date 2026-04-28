@echo off
REM =============================================================
REM Kingdom Structural — Excel Client List sync (standalone)
REM
REM Standalone tool — NOT part of the KS Automation Sweep.
REM Runs sync_clients_from_excel.py in --live mode and appends
REM its output to excel_sync.log.
REM
REM Put this into Windows Task Scheduler on its own schedule
REM (daily is fine) or run it manually when you update the
REM Master Business Plan Excel.
REM
REM Setup:
REM   Program:   C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\_automation\run_excel_sync.bat
REM   Start in:  C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\_automation
REM =============================================================

setlocal
cd /d "%~dp0"

echo. >> excel_sync.log
echo === %DATE% %TIME% === >> excel_sync.log

"C:\Users\MichaelBrusnahan\AppData\Local\Programs\Python\Python312\python.exe" sync_clients_from_excel.py --live >> excel_sync.log 2>&1

exit /b %ERRORLEVEL%
