@echo off
REM One-time setup: makes print_bridge_run.bat start automatically every time this PC's
REM cashier account logs in -- so a reboot, power cut, or someone closing the window no
REM longer means printing stays off until a human notices and re-runs the script by hand.
REM
REM HOW TO RUN THIS (once, on the cafe's PC):
REM   1. Put this file and print_bridge_run.bat and print_bridge.py in the same folder.
REM   2. Right-click this file -> "Run as administrator".
REM   3. Press any key when it's done, then check Task Scheduler (search "Task Scheduler"
REM      in the Start menu) -> Task Scheduler Library -> you should see "WaqtiPrintBridge".
REM   4. To confirm it actually works without rebooting: run this, then type
REM        schtasks /run /tn "WaqtiPrintBridge"
REM      and check print_bridge.log appears/updates in this same folder.
REM
REM To remove it later: schtasks /delete /tn "WaqtiPrintBridge" /f

setlocal
set TASK_NAME=WaqtiPrintBridge
set SCRIPT_DIR=%~dp0

schtasks /create /tn "%TASK_NAME%" /tr "\"%SCRIPT_DIR%print_bridge_run.bat\"" /sc onlogon /rl highest /f

echo.
echo Installed task "%TASK_NAME%" -- it will now start automatically at every logon on this PC.
echo Test it right now without rebooting:   schtasks /run /tn "%TASK_NAME%"
echo Remove it later if ever needed:        schtasks /delete /tn "%TASK_NAME%" /f
pause
