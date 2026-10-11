@echo off
REM Wrapper for print_bridge.py -- keeps it running forever.
REM print_bridge.py's own polling loop already survives network/printer errors (it has a
REM try/except around the whole thing), so this wrapper only needs to handle the case where
REM the Python process itself dies outright (closed window, Python crash, PC woke from
REM sleep into a bad state, etc.) -- if that happens, restart it after a short pause instead
REM of leaving the cafe with no printing until someone notices and reruns this by hand.
REM Output goes to print_bridge.log next to this file, not a console window, so there's a
REM record to check later instead of errors scrolling past in a window nobody's watching.

REM Calling python.exe by its FULL PATH, not just "python" -- the Microsoft Store
REM app-execution-alias stub (...\AppData\Local\Microsoft\WindowsApps\python.exe) works fine
REM double-clicked or run from an interactive terminal, but CONFIRMED BROKEN on the Squeeze
REM cafe PC (2026-10-12) the moment its stdout/stderr get redirected to a file -- which is
REM exactly what the line below does, so this wrapper silently produced zero output forever
REM (order printing + تقفيل reports both just stopped, no error anywhere) even though running
REM `python print_bridge.py` directly in a terminal worked perfectly. Fix: install the real
REM Python from python.org (NOT the Store one) and point PYTHON_EXE at that install instead --
REM its default per-user install path, used below.
set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe

cd /d "%~dp0"

:loop
echo %date% %time% - starting print_bridge.py >> print_bridge.log
"%PYTHON_EXE%" print_bridge.py >> print_bridge.log 2>&1
echo %date% %time% - print_bridge.py exited, restarting in 5s... >> print_bridge.log
timeout /t 5 /nobreak >nul
goto loop
