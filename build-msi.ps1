<#
.SYNOPSIS
  Builds YearView-<version>.msi from YearView.wxs using the WiX Toolset v3.

.DESCRIPTION
  The shipped MSI was produced with msitools/wixl. This script rebuilds the same
  package with WiX v3 (candle + light) on Windows, which additionally honours
  Shortcut/@Show and runs ICE validation.

  WiX v3.11+:  https://github.com/wixtoolset/wix3/releases  (install, or set -WixBin)

.PARAMETER WixBin
  Folder containing candle.exe and light.exe. Defaults to the usual install path.

.PARAMETER Sign
  Optional path to a code-signing certificate (.pfx). Signing is strongly
  recommended: an unsigned MSI shows "Unknown publisher" in SmartScreen, and the
  package installs a script that runs at every launch (and at sign-in with the
  autostart feature).

.EXAMPLE
  .\build-msi.ps1
  .\build-msi.ps1 -Sign C:\path\to\certificate.pfx
#>
[CmdletBinding()]
param(
    [string]$WixBin = "${env:ProgramFiles(x86)}\WiX Toolset v3.14\bin",
    [string]$Version = '3.4.9',
    [string]$Sign,
    [string]$TimestampUrl = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$candle = Join-Path $WixBin 'candle.exe'
$light  = Join-Path $WixBin 'light.exe'
foreach ($t in $candle, $light) {
    if (-not (Test-Path $t)) { throw "Not found: $t  (use -WixBin to point at the WiX v3 bin folder)" }
}

foreach ($f in 'src\New-OutlookYearView.ps1', 'src\YearView.cmd', 'src\Stop-YearView.cmd', 'src\YearView.ico') {
    if (-not (Test-Path $f)) { throw "Missing payload file: $f" }
}

$msi = "YearView-$Version.msi"
Write-Host "Building $msi ..." -ForegroundColor Cyan

& $candle -nologo -arch x86 -out obj\YearView.wixobj YearView.wxs
if ($LASTEXITCODE) { throw "candle failed ($LASTEXITCODE)" }

# ICE suppressions, each deliberate:
#   ICE38 / ICE64 - components install to the user profile with a file KeyPath.
#                   Correct here: this is a pure per-user package (ALLUSERS unset),
#                   so the profile IS the install target. The shortcut components do
#                   use HKCU KeyPaths, as ICE43 requires.
#   ICE91         - warns that LocalAppDataFolder is a per-user location. Intended.
& $light -nologo -out $msi -sice:ICE38 -sice:ICE64 -sice:ICE91 obj\YearView.wixobj
if ($LASTEXITCODE) { throw "light failed ($LASTEXITCODE)" }

if ($Sign) {
    $signtool = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin" -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '\\x64\\' } | Select-Object -Last 1
    if (-not $signtool) { throw 'signtool.exe not found (install the Windows SDK).' }
    & $signtool.FullName sign /fd SHA256 /f $Sign /tr $TimestampUrl /td SHA256 $msi
    if ($LASTEXITCODE) { throw "signtool failed ($LASTEXITCODE)" }
    Write-Host 'Signed.' -ForegroundColor Green
}

Write-Host ("Built {0} ({1:n0} bytes)" -f $msi, (Get-Item $msi).Length) -ForegroundColor Green
Write-Host ''
Write-Host 'Install (per-user, no elevation):' -ForegroundColor Cyan
Write-Host "  msiexec /i $msi"
Write-Host "  msiexec /i $msi /qn                                        # silent"
Write-Host "  msiexec /i $msi ADDLOCAL=FeatureMain,FeatureAutostart      # + start at sign-in"
Write-Host "  msiexec /x $msi                                            # uninstall"
Write-Host "  msiexec /i $msi /l*v install.log                           # verbose log"
