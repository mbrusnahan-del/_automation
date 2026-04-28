@echo off
cd /d "%~dp0"

echo. >> sweep.log
echo === %DATE% %TIME% === >> sweep.log

"C:\Users\MichaelBrusnahan\AppData\Local\Programs\Python\Python312\python.exe" sweep.py >> sweep.log 2>&1

exit /b %ERRORLEVEL%
