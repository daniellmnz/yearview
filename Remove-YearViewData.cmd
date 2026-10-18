@echo off
rem ------------------------------------------------------------------
rem  Removes YearView's runtime data files so a genuine uninstall leaves
rem  no trace in %LOCALAPPDATA%\YearView. Invoked by the MSI ONLY on a
rem  real uninstall (condition: REMOVE=ALL AND NOT UPGRADINGPRODUCTCODE)
rem  and never during an upgrade, so version updates preserve the page
rem  cache, calendar selection and saved filter. These files are created
rem  by the app at runtime, are not owned by the installer, and would
rem  otherwise be left behind. Safe to run by hand; always exits 0.
rem ------------------------------------------------------------------
set "YVDIR=%~dp0"
rem v3.4.9: the Edge app window runs from its OWN profile under this folder
rem (since v3.3.2). Close any window still using it - it holds the profile's
rem files open - then remove the profile, which also holds that window's
rem browsing history. Before this an uninstall left the whole profile behind.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p = Join-Path $env:YVDIR 'edge-profile'; $n = 0; Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0 } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $n++ }; if ($n) { Start-Sleep -Milliseconds 1500 }"
if exist "%YVDIR%edge-profile\" rd /s /q "%YVDIR%edge-profile" 2>nul
del /q "%YVDIR%yv-page-*.html"       2>nul
del /q "%YVDIR%yv-page-*.html.tmp"   2>nul
del /q "%YVDIR%yv-page-*.meta"       2>nul
del /q "%YVDIR%yv-page-*.meta.tmp"   2>nul
del /q "%YVDIR%yv-calendar.cfg"      2>nul
del /q "%YVDIR%yv-filter.cfg"        2>nul
del /q "%YVDIR%yv-window.cfg"        2>nul
del /q "%YVDIR%yv-settings.cfg"      2>nul
rem v2.1.29: the config writer stages through .tmp; a crash mid-write can
rem leave one behind, and "leave no trace" has to mean no trace.
del /q "%YVDIR%yv-calendar.cfg.tmp"  2>nul
del /q "%YVDIR%yv-filter.cfg.tmp"    2>nul
del /q "%YVDIR%yv-window.cfg.tmp"    2>nul
del /q "%YVDIR%yv-settings.cfg.tmp"  2>nul
del /q "%YVDIR%yv-fastscan.ok"       2>nul
del /q "%YVDIR%yv-error.log"         2>nul
del /q "%YVDIR%YearView-Signing.cer" 2>nul
exit /b 0
