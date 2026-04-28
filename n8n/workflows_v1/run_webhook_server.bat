@echo off
REM =============================================================
REM KS Webhook Server launcher
REM
REM Put this into shell:startup to auto-start the server when
REM you log in to Windows. Or run it manually from this folder.
REM =============================================================

setlocal
cd /d "%~dp0"
python webhook_server.py
pause
