@echo off
REM One-tap morning launcher for cafe staff -- opens the admin dashboard AND starts the
REM print bridge from a single icon, so nobody needs to touch a folder/terminal day to day.
REM Meant to be dragged onto the Desktop (see Ahmed's setup notes). Each piece runs in its
REM own window so closing the browser tab doesn't kill printing and vice versa.
title Waqti - Squeeze - ابدأ اليوم

REM Print bridge gets its own separate console window (same script as the manual fallback),
REM kept open and visible so staff can glance at it to confirm printing is alive.
start "Waqti Print Bridge" "%~dp0start_printing_squeeze.bat"

REM Admin dashboard in the default browser -- PIN-gated, so safe to launch straight to it.
start "" "https://farghalx.github.io/waqti-cafe-menu/admin.html?c=squeeze"
