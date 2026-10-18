@echo off
rem ------------------------------------------------------------------
rem  YearView launcher - double-click or point a shortcut at this file.
rem  Starts the Outlook year view bridge minimized and opens the browser.
rem  Keep this file in the same folder as New-OutlookYearView.ps1.
rem ------------------------------------------------------------------
start "YV-helper" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0New-OutlookYearView.ps1" -Serve -App
