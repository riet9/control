@echo off
if exist "%~dp0ScreenTimeTracker.exe" (
    start "" "%~dp0ScreenTimeTracker.exe" %*
) else (
    start "" wscript.exe "%~dp0start-tracker.vbs" %*
)
