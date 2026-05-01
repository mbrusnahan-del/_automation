@echo off
REM =============================================================
REM Kingdom Structural — Master Business Plan → Notion sync
REM
REM Standalone tool — NOT part of the KS Automation Sweep.
REM Runs both Excel→Notion syncs in --live mode against the
REM Master Business Plan workbook and appends each one's output
REM to its own log file:
REM
REM   1. sync_clients_from_excel.py  → excel_sync.log
REM      Client List → Notion Clients + Contacts DB
REM   2. sync_jobs_from_excel.py     → jobs_sync.log
REM      Jobs 26KS → Notion Jobs 26KS DB (totals only)
REM
REM Both scripts mtime-cache so re-runs against an unchanged
REM workbook exit fast.
REM
REM Setup:
REM   Program:   C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\_automation\run_excel_sync.bat
REM   Start in:  C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\_automation
REM =============================================================

setlocal
cd /d "%~dp0"

set PYTHON="C:\Users\MichaelBrusnahan\AppData\Local\Programs\Python\Python312\python.exe"

REM -- Client List sync --
echo. >> excel_sync.log
echo === %DATE% %TIME% === >> excel_sync.log
%PYTHON% sync_clients_from_excel.py --live >> excel_sync.log 2>&1
set CLIENT_RC=%ERRORLEVEL%

REM -- Jobs 26KS sync --
echo. >> jobs_sync.log
echo === %DATE% %TIME% === >> jobs_sync.log
%PYTHON% sync_jobs_from_excel.py --live >> jobs_sync.log 2>&1
set JOBS_RC=%ERRORLEVEL%

REM Exit non-zero if EITHER sync had a problem so Task Scheduler
REM flags the run accurately.
if %CLIENT_RC% neq 0 exit /b %CLIENT_RC%
exit /b %JOBS_RC%
