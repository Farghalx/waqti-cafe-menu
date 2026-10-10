@echo off
REM Wrapper for print_bridge.py -- keeps it running forever.
REM print_bridge.py's own polling loop already survives network/printer errors (it has a
REM try/except around the whole thing), so this wrapper only needs to handle the case where
REM the Python process itself dies outright (closed window, Python crash, PC woke from
REM sleep into a bad state, etc.) -- if that happens, restart it after a short pause instead
REM of leaving the cafe with no printing until someone notices and reruns this by hand.
REM Output goes to print_bridge.log next to this file, not a console window, so there's a
REM record to check later instead of errors scrolling past in a window nobody's watching.

REM Calling python.exe by its FULL PATH, not just "python" -- on this cafe PC, "python" on
REM PATH resolves to the Microsoft Store app-execution-alias stub
REM (...\AppData\Local\Microsoft\WindowsApps\python.exe), which works fine double-clicked
REM or run from an interactive terminal but is known to fail (0xc0000142) when launched by
REM Task Scheduler or any other non-interactive/background launcher -- it needs the normal
REM interactive shell activation that alias relies on. Calling it by full path sidesteps
REM nothing about THAT specific flakiness (still the same stub), so if this still crashes
REM under the scheduled task, the real fix is installing actual Python from python.org and
REM pointing this at THAT python.exe instead -- see the handoff notes.
set PYTHON_EXE=%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe

cd /d "%~dp0"

:loop
echo %date% %time% - starting print_bridge.py >> print_bridge.log
"%PYTHON_EXE%" print_bridge.py >> print_bridge.log 2>&1
echo %date% %time% - print_bridge.py exited, restarting in 5s... >> print_bridge.log
timeout /t 5 /nobreak >nul
goto loop
