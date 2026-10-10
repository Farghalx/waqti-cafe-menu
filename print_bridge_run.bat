@echo off
REM Wrapper for print_bridge.py -- keeps it running forever.
REM print_bridge.py's own polling loop already survives network/printer errors (it has a
REM try/except around the whole thing), so this wrapper only needs to handle the case where
REM the Python process itself dies outright (closed window, Python crash, PC woke from
REM sleep into a bad state, etc.) -- if that happens, restart it after a short pause instead
REM of leaving the cafe with no printing until someone notices and reruns this by hand.
REM Output goes to print_bridge.log next to this file, not a console window, so there's a
REM record to check later instead of errors scrolling past in a window nobody's watching.

cd /d "%~dp0"

:loop
echo %date% %time% - starting print_bridge.py >> print_bridge.log
python print_bridge.py >> print_bridge.log 2>&1
echo %date% %time% - print_bridge.py exited, restarting in 5s... >> print_bridge.log
timeout /t 5 /nobreak >nul
goto loop
