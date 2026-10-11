@echo off
REM Manual daily fallback until the Task Scheduler auto-restart (install_print_bridge_task.bat
REM + print_bridge_run.bat) is confirmed working unattended on this PC -- double-tap this file
REM to start printing, no typing needed. Keeps the window open and VISIBLE (unlike
REM print_bridge_run.bat, this never redirects output to a log file) so staff can see at a
REM glance whether it's running or has stopped.
title Waqti Print Bridge - Squeeze
cd /d "%~dp0"
echo بيشتغل... سيب الشباك ده مفتوح طول اليوم
echo.
C:\Users\Pc\AppData\Local\Programs\Python\Python312\python.exe print_bridge.py
echo.
echo وقف! اقفل الشباك ده وافتح start_printing_squeeze تاني
pause
