@echo off
rem ------------------------------------------------------------------
rem  Retargets any legacy YearView shortcut that still points straight
rem  at powershell.exe (which flashes a console box before the hidden
rem  style applies) to the windowless wscript launcher. Covers Desktop,
rem  Start Menu, Startup, and PINNED TASKBAR shortcuts - pinning makes
rem  a private copy of the .lnk that installers never otherwise touch.
rem  Preserves -NoOpen on startup shortcuts. Safe to run by hand.
rem
rem  v3.4.3 also repoints every YearView shortcut's ICON at the
rem  installed YearView.ico. Windows Installer extracts an Icon table
rem  entry to a folder named after the PRODUCT CODE, which changes
rem  every release - so after an upgrade a pinned copy points into a
rem  directory that no longer exists and shows a blank page. The
rem  installed .ico sits at a path that never changes.
rem ------------------------------------------------------------------
set "YVDIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; $vbs = Join-Path $env:YVDIR 'YearView-launch.vbs'; $dirs = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'), [Environment]::GetFolderPath('Startup'), (Join-Path $env:APPDATA 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar')); foreach ($d in $dirs) { if (-not (Test-Path $d)) { continue }; Get-ChildItem $d -Filter *.lnk -ErrorAction SilentlyContinue | ForEach-Object { try { $sc = $ws.CreateShortcut($_.FullName); if ($sc.TargetPath -like '*powershell.exe' -and $sc.Arguments -like '*New-OutlookYearView.ps1*') { $no = ''; if ($sc.Arguments -like '*-NoOpen*') { $no = ' -Serve -App -NoOpen' }; $sc.TargetPath = (Join-Path $env:SystemRoot 'System32\wscript.exe'); $sc.Arguments = ('\"{0}\"' -f $vbs) + $no; $sc.Save() } } catch {} } }"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; $ico = Join-Path $env:YVDIR 'YearView.ico'; if (Test-Path $ico) { $dirs = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'), [Environment]::GetFolderPath('Startup'), (Join-Path $env:APPDATA 'Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar')); foreach ($d in $dirs) { if (-not (Test-Path $d)) { continue }; Get-ChildItem $d -Filter *.lnk -ErrorAction SilentlyContinue | ForEach-Object { try { $sc = $ws.CreateShortcut($_.FullName); if ($sc.Arguments -like '*YearView-launch.vbs*' -or $sc.Arguments -like '*New-OutlookYearView.ps1*') { $want = $ico + ',0'; if ($sc.IconLocation -ne $want) { $sc.IconLocation = $want; $sc.Save() } } } catch {} } } }"
ie4uinit.exe -show >nul 2>&1
exit /b 0
