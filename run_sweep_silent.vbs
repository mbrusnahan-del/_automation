' run_sweep_silent.vbs
' --------------------
' Launches run_sweep.bat with no visible console window. Used by Windows
' Task Scheduler so the 5-minute sweep doesn't flash a black box on screen.
'
' Why this exists: Task Scheduler running a .bat directly always shows a
' brief cmd.exe window (even with @echo off) on a "Run only when user is
' logged on" task. The hidden-window option in Task Scheduler doesn't fully
' suppress it. A VBScript launched with WindowStyle=0 actually disappears.
'
' Setup in Task Scheduler:
'   Action: Start a program
'   Program/script: wscript.exe
'   Add arguments: "C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\_automation\run_sweep_silent.vbs"
'   Start in:     C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\_automation
'
' All sweep output still goes to sweep.log; this script only changes how
' the launcher is displayed (or rather, not displayed).

Set WshShell = CreateObject("WScript.Shell")

' Resolve this script's folder so the .bat path stays portable across users.
strScriptPath = WScript.ScriptFullName
strFolder = Left(strScriptPath, InStrRev(strScriptPath, "\") - 1)
strBat = strFolder & "\run_sweep.bat"

' Run flags:
'   0     = hide window
'   True  = wait for the .bat to finish (so Task Scheduler reports the
'           correct exit code)
WshShell.Run """" & strBat & """", 0, True
