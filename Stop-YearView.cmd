@echo off
rem ------------------------------------------------------------------
rem  Stops a running YearView bridge (the hidden "YV-helper" PowerShell
rem  process) launched from THIS folder with -Serve. Used by the MSI
rem  before install and before uninstall, so a stale bridge cannot keep
rem  serving the old view on port 8377. (S-9: path- and switch-pinned
rem  via IndexOf, so wildcard characters in paths cannot mismatch.)
rem  Safe to run by hand; exits 0 even when nothing is running.
rem ------------------------------------------------------------------
set "YVDIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$t = Join-Path $env:YVDIR 'New-OutlookYearView.ps1'; Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($t, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and $_.CommandLine.IndexOf('-Serve', [StringComparison]::OrdinalIgnoreCase) -ge 0 } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
exit /b 0
