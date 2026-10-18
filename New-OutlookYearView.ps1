<#
.SYNOPSIS
  Year-at-a-glance HTML view of your Outlook calendar, with click-to-create.

.DESCRIPTION
  Reads all categorised appointments for the given year from classic Outlook
  via COM (recurrences expanded, category colours applied) and renders the
  familiar months-as-rows year view.

  Two modes:
    Default   - writes a self-contained HTML snapshot to your Desktop.
    -Serve    - additionally runs a local bridge on 127.0.0.1 so that clicking
                a category in the legend and then a day in the grid opens a
                new Outlook appointment window with that category pre-set.
                Keep the console window open while using the view; Ctrl+C to
                stop.

.EXAMPLE
  .\New-OutlookYearView.ps1                    # snapshot of current year
  .\New-OutlookYearView.ps1 -Serve             # interactive mode
  .\New-OutlookYearView.ps1 -Year 2027 -Serve
  .\New-OutlookYearView.ps1 -IncludeUncategorised

.NOTES
  Requires classic Outlook (new Outlook has no COM interface).
  Desktop shortcut for one-click interactive use:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Tools\New-OutlookYearView.ps1" -Serve
#>
# Classic Windows PowerShell 5.1 only: Outlook COM interop and
# Marshal.GetActiveObject need the .NET Framework, not .NET (pwsh).
#Requires -PSEdition Desktop
[CmdletBinding()]
param(
    [ValidateRange(1900, 9999)]
    [int]$Year = (Get-Date).Year,
    [string]$OutFile,
    [switch]$NoOpen,
    [switch]$IncludeUncategorised,   # '(No category)' filter starts VISIBLE instead of hidden (first run only; saved filter settings win thereafter)
    [switch]$ClassicScan,            # force the original single-path calendar reader (skips the fast Table reader and its validation entirely)
    [switch]$Serve,                  # run the local bridge for click-to-create
    [switch]$App,                    # open in a chromeless Edge app window instead of a browser tab
    [switch]$Visible,                # keep the console visible (for debugging)
    [switch]$NoOutlookPrompt,        # do not show the "please start Outlook" dialog at launch
    [switch]$HiddenChild,            # internal: this instance is the hidden respawn
    [ValidateRange(0, 10000)]
    [int]$WindowWidth = 0,           # explicit app-window size (default: ~92% of the screen work area)
    [ValidateRange(0, 10000)]
    [int]$WindowHeight = 0,
    [ValidateRange(1, 65535)]
    [int]$Port = 8377
)

$ErrorActionPreference = 'Stop'
try {

try { $Host.UI.RawUI.WindowTitle = 'YV-helper' } catch {}

# ==================== script-wide state - the ONE init block ====================
# Everything any function below may read is initialised HERE, before any of them
# can run. Two releases were bitten by state defined further down the file
# (v1.1.22: titlePort in the main-loop init, read by the pre-paint shrink before
# that init executed; the same latent risk existed for $script:pages, read by
# Ensure-Outlook, and $quitRequested, never initialised at all). Rule from here
# on: code below ASSIGNS these; nothing below INITIALISES a new script-scope
# variable.

# constants
$script:invCulture   = [Globalization.CultureInfo]::InvariantCulture  # F-1: fixed wire-format culture
# v3.4.9: page placeholders, filled in ONE pass (see Build-YearViewHtml)
$script:placeholderRx = [regex]'__([A-Z0-9]+)__'
$script:jsonEscRx    = [regex]'["\\\u0000-\u001F]'                     # chars needing escape in JSON strings
$script:fastJsonChecked = $false          # first build validates the fast serialiser once
$script:fastJsonOk      = $true           # flips off (for the session) if validation ever fails
$script:lastBaselineMs  = 0               # /build's re-baseline sort duration, shown in the perf stamp
# Hybrid year reader (v1.2.6). Field evidence: batched property reads changed
# nothing (laptop 10 s -> 11 s / 753 items), isolating the cost as the
# recurrence-expanded enumerator MATERIALISING each item - single-threaded
# Outlook CPU (~14 ms/item; the i9 desktop is 2.3x faster, tracking core speed),
# not property access. The fast reader therefore pushes fewer items through that
# path: non-recurring items come from the Table API (columnar, no per-item COM
# object), and only recurring masters go through Sort/IncludeRecurrences/
# Restrict. The FIRST build each session runs BOTH readers and compares results
# item-for-item on the live calendar; any error or mismatch falls back to the
# classic reader permanently, with a warning - so output can never differ.
$script:fastScanChecked = $false
$script:fastScanOk      = $true
# A waitable handle whose ONLY job is pumping: on an STA thread the CLR services
# COM messages during managed blocking waits (WaitOne), which Start-Sleep never
# does - and Outlook's ItemAdd/ItemChange/ItemRemove connection-point events can
# only be delivered while the thread pumps. Without this, item EDITS (which the
# per-poll count check cannot see) were only ever caught by the throttled
# LastModificationTime fingerprint - the reported 10-20 s update lag.
$script:comPump = New-Object Threading.ManualResetEvent($false)
$script:modCheckInterval = -1             # adaptive fingerprint cadence (seconds); -1 = measure on first run
$script:yvVersion = '3.4.9'              # ALSO bump the stamp string when changing this
$script:contentHashes = @{}               # per-year order-independent content hash (echo suppression)
# Calendar selection (v2.0.0): empty = the default calendar. Restricted by
# design to calendars in the DEFAULT STORE (the picker only offers those);
# shared/delegate calendars are explicitly out of scope in this version.
$script:selCalId      = ''                # EntryID of the selected calendar folder
$script:selCalStoreId = ''                # its StoreID
$script:selCalName    = ''                # display name (for toasts/logs)
$script:calList       = @()               # last /calendars enumeration - lets the client switch by tiny index (v2.0.5)
$script:hiddenCsv     = $null             # persisted filter csv (v2.0.8): $null = never set; '' = explicit "show all"
$script:listener = $null                  # TcpListener, shared with Invoke-QuickServe (set at serve start)
$script:lastBuildHash = ''                # hash of the most recent build, set by Build-YearViewHtml
if ($ClassicScan) {
    # user-forced classic: no fast reader, no validation pass
    $script:fastScanOk = $false
    $script:fastScanChecked = $true
}
$script:titlePort    = ('*127.0.0.1:{0}*' -f $Port)                   # R-14: port-pinned URL-title match
$script:sessionToken = [guid]::NewGuid().ToString('N')                # S-2: per-session auth token

# Outlook COM state (assigned by Connect-Outlook / Ensure-Outlook / Build)
$script:outlook        = $null           # Outlook.Application
$script:ns             = $null           # MAPI namespace
$script:outlookStarted = $false          # F-3: did this run launch Outlook itself?
$script:watchItems     = $null           # calendar Items kept referenced for change events
$script:calStoreId     = $null           # default store id for GetItemFromID

# change detection / page cache (bridge)
$script:pages           = @{}            # per-year rendered page cache (UTF-8 bytes)
# v3.4.1: per-year PAYLOAD cache - the same JSON a page carries, kept on its
# own so /data can answer without a calendar read. A view that starts after
# January needs the following year on EVERY page load, and every background
# update navigates, so without this the spill year was re-read from Outlook
# several times a session. Each entry carries the tick it was made at and is
# refused once anything in the calendar has changed since.
$script:payloads        = @{}            # year -> @{ Json; Hash; Ticks }
# v3.4.5: the content hash of the SPILL payload the last build embedded, and
# the same per displayed year. Change detection compared only the year on
# screen, so editing an appointment in the spill months moved no hash the
# bridge was watching: it reported 'nothing changed', the page never
# navigated, and the edit stayed invisible until something else forced a
# rebuild. A view that shows two years has to watch two years.
$script:lastSpillHash   = ''             # spill hash of the most recent build
$script:spillHashes     = @{}            # year -> spill hash embedded in that year's page
$script:pageHideLits    = @{}            # v2.1.22: the exact HIDE_SEED literal baked into each cached page (for re-stamping)
$script:confirmDelete   = $true          # v2.1.29: ask before deleting (default ON - deletion is irreversible in the view)
$script:scrollYears     = $false         # v3.1.6: continuous scrolling across years is now OPT-IN
# v3.4.7: categories whose bars are drawn WITHOUT their titles, as the client's
# csv ('__none__' = no category). Empty = every category shows its titles.
$script:noTitles        = ''
# v3.4.0: which month the year view starts at. -1 follows the current month,
# 0-11 fix it; 0 is the ordinary January-to-December year.
$script:startMonth      = 0
$script:printPages      = 1              # v3.2.0: sheets to scale a print onto; 0 = as many as it takes
$script:printLegend     = $true          # v3.4.8: print the category legend row at the foot of the sheet
$script:lastCount       = -1             # baseline item count
$script:lastModTicks    = 0L             # newest-LastModificationTime fingerprint
$script:lastChangeTicks = 0L             # stamp of last detected change (drives /dirty)
$script:nextModCheck    = Get-Date       # R-12: fingerprint-sort throttle

# window management (bridge, -App mode)
$script:resizeTarget  = $null            # two-stage sizing state (set by Open-View)
$script:appHwnd       = [IntPtr]::Zero   # watchdog anchor (captured by resize logic)
$script:windowSeen    = $false           # watchdog: app window sighted at least once
$script:quitRequested = $false           # set by /quit and by the watchdog
$script:outlookGoneHandled = $false      # v2.1.2: the "Outlook has gone" shutdown runs once
$script:noExplorerTicks = 0              # v2.1.5: consecutive 2s checks with no Outlook window open
$script:closingAt = $null                # v2.1.6: announce the closure to the page, THEN close

# v2.1.11: crisp dialogs on high-DPI displays. Without process DPI awareness,
# Windows bitmap-stretches our WinForms message boxes to the monitor scale and
# the text goes slightly fuzzy. Declare per-monitor-V2 awareness (Win10 1703+),
# falling back through per-monitor and then system awareness on older builds.
# Must run before ANY window or dialog is created; the Edge view is a separate
# process and unaffected either way.
try {
    Add-Type -Namespace YVDpi -Name Native -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
[DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value);
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
'@ -ErrorAction Stop
    $pmv2 = [IntPtr]::op_Explicit([int64]-4)   # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    if (-not [YVDpi.Native]::SetProcessDpiAwarenessContext($pmv2)) {
        if ([YVDpi.Native]::SetProcessDpiAwareness(2) -ne 0) {   # PROCESS_PER_MONITOR_DPI_AWARE
            [void][YVDpi.Native]::SetProcessDPIAware()
        }
    }
} catch {
    try { [void][YVDpi.Native]::SetProcessDPIAware() } catch {}
}
$script:respondedTo   = $null            # v2.0.24: the stream the current reply went out on
$script:savedWinRect  = $null            # user window geometry (v2.0.16): @(w,h,x,y) from yv-window.cfg, $null = none
$script:savedWinMax   = $false           # v2.0.17: window was maximised when last seen
$script:lastNormalRect = $null           # last known NON-maximised rect @(w,h,x,y) - what restore-down returns to
$script:lastWinRect   = ''               # last geometry written, as "w,h,x,y,max" - avoids rewriting an unchanged file
# =================================================================================

function ConvertTo-CliArg([string]$s) {
    # R-18: full Windows argv quoting (MSVCRT rules). Bare "..." wrapping breaks
    # on an embedded quote or a trailing backslash before the closing quote:
    # backslashes preceding a quote are doubled, the quote gets one escaping
    # backslash, and trailing backslashes are doubled before the final quote.
    if ($s.Length -and $s -notmatch '[\s"]') { return $s }
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('"')
    $bs = 0
    foreach ($ch in $s.ToCharArray()) {
        if ($ch -eq '\') { $bs++; continue }
        if ($ch -eq '"') {
            [void]$sb.Append('\' * ($bs * 2 + 1)); [void]$sb.Append('"'); $bs = 0; continue
        }
        if ($bs) { [void]$sb.Append('\' * $bs); $bs = 0 }
        [void]$sb.Append($ch)
    }
    if ($bs) { [void]$sb.Append('\' * ($bs * 2)) }
    [void]$sb.Append('"')
    $sb.ToString()
}

function ConvertTo-JsonStr([string]$s) {
    # JSON string literal with full escaping. The common no-specials case is a
    # single native regex scan with zero callback invocations, so this is far
    # cheaper than ConvertTo-Json while escaping exactly the same characters.
    '"' + $script:jsonEscRx.Replace($s, {
        param($m)
        switch ($m.Value) {
            '"'  { '\"' }
            '\'  { '\\' }
            "`t" { '\t' }
            "`n" { '\n' }
            "`r" { '\r' }
            default { '\u{0:x4}' -f [int][char]$m.Value }
        }
    }) + '"'
}

function ConvertTo-SeedStr([string]$s) {
    # v3.4.7: a JSON string literal that is safe both to SEED into the page's
    # script block and to RE-STAMP later. '<' is escaped so no '</script>' in a
    # category name can end the block, and ';' so the literal can never hold
    # the ';' that Update-SettingLiterals reads as its end. Both are ordinary
    # JSON escapes and decode to exactly the same text.
    (ConvertTo-JsonStr $s).Replace('<', '\u003c').Replace(';', '\u003b')
}
function Test-TitleCsv([string]$s) {
    # v3.4.7: the no-titles category list, checked at load AND at the route.
    # Outlook refuses commas in category names, so a comma can only ever be a
    # separator. Control characters are refused - a newline would split the
    # settings file - and so is anything implausibly long.
    if ($null -eq $s -or $s.Length -gt 16384) { return $false }
    if ($s -match '[\x00-\x1F]') { return $false }
    foreach ($n in $s.Split(',')) { if ($n.Length -gt 255) { return $false } }
    return $true
}

# DPAPI (v1.2.29): the page cache holds verbatim calendar content at rest under
# %LOCALAPPDATA%; ProtectedData needs this assembly in Windows PowerShell 5.1.
try { Add-Type -AssemblyName System.Security -ErrorAction Stop } catch {}

# R-16: keep the crash log bounded - it was append-only and never rotated.
try {
    $yvLog = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
    if ((Test-Path $yvLog) -and ((Get-Item $yvLog).Length -gt 64KB)) {
        (Get-Content $yvLog -Tail 200) | Set-Content $yvLog
    }
} catch {}

# In serve mode, respawn as a truly console-less process (CreateNoWindow) so no
# taskbar entry appears - Windows Terminal ignores ordinary hidden/minimized
# launch styles. Quit / closing the window shuts the bridge down; -Visible
# keeps the console for debugging.
if ($Serve -and -not $HiddenChild -and -not $Visible) {
    $argsList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Serve', '-HiddenChild')
    if ($App) { $argsList += '-App' }
    if ($NoOpen) { $argsList += '-NoOpen' }
    if ($IncludeUncategorised) { $argsList += '-IncludeUncategorised' }
    if ($NoOutlookPrompt) { $argsList += '-NoOutlookPrompt' }
    if ($PSBoundParameters.ContainsKey('Year')) { $argsList += @('-Year', [string]$Year) }
    if ($PSBoundParameters.ContainsKey('Port')) { $argsList += @('-Port', [string]$Port) }
    if ($WindowWidth -gt 0)  { $argsList += @('-WindowWidth',  [string]$WindowWidth) }
    if ($WindowHeight -gt 0) { $argsList += @('-WindowHeight', [string]$WindowHeight) }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = ($argsList | ForEach-Object { ConvertTo-CliArg $_ }) -join ' '   # R-18
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    [void][System.Diagnostics.Process]::Start($psi)
    return
}

if (-not $OutFile) {
    $OutFile = Join-Path ([Environment]::GetFolderPath('Desktop')) ("YearView_{0}.html" -f $Year)
}

# Outlook category colour enum (OlCategoryColor 0-25) -> hex
$catColors = @{
    0='#BFBFBF'; 1='#E7A1A2'; 2='#F9BA89'; 3='#F7DD8F'; 4='#FCFA90'; 5='#78D168'
    6='#9FDCC9'; 7='#C6D2B0'; 8='#9DB7E8'; 9='#B5A1E2'; 10='#DAAEC2'; 11='#DAD9DC'
    12='#6B7994'; 13='#BFBFBF'; 14='#6F6F6F'; 15='#4F4F4F'; 16='#C11A25'; 17='#E2620D'
    18='#C79930'; 19='#B9B300'; 20='#368F2B'; 21='#329B7A'; 22='#778B45'; 23='#2858A5'
    24='#5C3FA3'; 25='#93446B'
}

function Connect-Outlook {
    Write-Host 'Connecting to Outlook...' -ForegroundColor Cyan
    try {
        $script:outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $script:outlookStarted = $false
    } catch {
        if ($Serve) {
            # v2.0.30: YearView NEVER starts Outlook. Six releases of
            # auto-start heuristics (launch, storm guard, forced minimise,
            # age and window readiness gates) all fought the same fact: an
            # Outlook launched by Start-Process can sit part-initialised in
            # ways no outside probe can judge, while an Outlook the USER
            # starts is fully initialised, ROT-registered and windowed. So
            # the user is asked instead - a dialog at startup, a persistent
            # request in the page mid-session - and this branch only REPORTS
            # which state holds: absent (please start it) or running but not
            # ready yet (starting). The window-readiness test keeps an attach
            # from wedging the loop while a user-started Outlook is still
            # booting; the age fallback covers windowless-but-functional
            # states such as a tray-hidden explorer.
            if (Get-Process -Name outlook -ErrorAction SilentlyContinue) {
                # v2.1.1: NEVER bind an Outlook that cannot answer. A
                # half-shut-down OUTLOOK.EXE has a process but no window and
                # has already revoked its ROT registration - which is why
                # GetActiveObject failed above - and New-Object against it
                # BLOCKS this single-threaded loop indefinitely. That wedge is
                # the field-reported permanent "Refreshing - Reading Outlook"
                # toast: no reply is ever written, so the page never learns
                # anything, not even that Outlook needs starting.
                #
                # So a window is now REQUIRED before attaching. A visible
                # explorer frame, or a main window handle on the process
                # (covering hidden/tray states), means the message loop is
                # alive. Without one: a young process is still booting, an old
                # one is stuck, and each gets its own honest verdict. We never
                # call New-Object in either case, so the bridge cannot wedge.
                $olReady = Test-OutlookWindowed   # v2.1.10: shared with the startup gate
                if (-not $olReady) {
                    $procAge = 999
                    try {
                        $op2 = @(Get-Process -Name outlook -ErrorAction SilentlyContinue) | Select-Object -First 1
                        if ($op2) { $procAge = ((Get-Date) - $op2.StartTime).TotalSeconds }
                    } catch {}
                    if ($procAge -lt 20) { throw 'YV_OUTLOOK_STARTING' }
                    throw 'YV_OUTLOOK_STUCK'
                }
                # Outlook is a singleton COM server: this binds the RUNNING
                # instance. Explicit namespace + return - no fallthrough to
                # the snapshot path below.
                $script:outlook = New-Object -ComObject Outlook.Application
                $script:outlookStarted = $false
                $script:ns = $script:outlook.GetNamespace('MAPI')
                # v2.1.0: hear Outlook's own goodbye - the Quit event fires as
                # Outlook begins closing, and the 2s link check releases our
                # references so the shutdown completes instead of zombifying.
                try { Register-ObjectEvent -InputObject $script:outlook -EventName Quit -SourceIdentifier 'YV_AppQuit' | Out-Null } catch {}
                return
            }
            throw 'YV_OUTLOOK_ABSENT'
        }
        $script:outlook = New-Object -ComObject Outlook.Application
        # F-3: this run launched Outlook headless; remember, so a snapshot can
        # close it again instead of leaving a windowless OUTLOOK.EXE behind.
        $script:outlookStarted = $true
    }
    $script:ns = $script:outlook.GetNamespace('MAPI')
    try { Register-ObjectEvent -InputObject $script:outlook -EventName Quit -SourceIdentifier 'YV_AppQuit' | Out-Null } catch {}   # v2.1.0
}
function Release-OutlookLink {
    # v2.1.0: let go of Outlook COMPLETELY, the moment it no longer serves us.
    # Closing Outlook while an external process holds COM references - above
    # all the ItemAdd/ItemChange/ItemRemove EVENT CONNECTION POINTS - pins
    # OUTLOOK.EXE in a windowless half-shutdown state that only End Task
    # clears (the field report behind this release). Unregister the events,
    # drop every RCW's native references deterministically, and collect twice
    # so the finalizer queue cannot keep the pin alive.
    #
    # $script:watchItems deliberately KEEPS its (now dead) value: Ensure-
    # Outlook keys firstInit vs reconnect off its NULLNESS, and its reconnect
    # branch reassigns it after Connect-Outlook before any use - so releasing
    # the COM side must not disturb the state machine's memory that a link
    # once existed.
    if ($null -eq $script:outlook -and $null -eq $script:ns) { return }
    Get-EventSubscriber -ErrorAction SilentlyContinue |
        Where-Object { $_.SourceIdentifier -like 'YV_*' } |
        Unregister-Event -ErrorAction SilentlyContinue
    Get-Event -ErrorAction SilentlyContinue |
        Where-Object { $_.SourceIdentifier -like 'YV_*' } |
        Remove-Event -ErrorAction SilentlyContinue
    foreach ($ref in @($script:watchItems, $script:ns, $script:outlook)) {
        if ($null -ne $ref) {
            try {
                if ([Runtime.InteropServices.Marshal]::IsComObject($ref)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ref)
                }
            } catch {}
        }
    }
    $script:outlook = $null
    $script:ns = $null
    # v2.1.4: NO WaitForPendingFinalizers here. An RCW finalizer releasing a
    # hung or half-shut-down Outlook can block forever, and this routine runs
    # on paths that must stay responsive. (The build-enumeration collect
    # elsewhere keeps its wait - that one only runs against a healthy Outlook
    # mid-build.) The Outlook-has-gone path no longer relies on this at all:
    # it ends the process, which drops every reference through the OS.
    [GC]::Collect()
    Write-Host 'Outlook link released.' -ForegroundColor DarkYellow
}
function Close-YearViewOutlookGone {
    # v2.1.2: Outlook has gone (or become unusable) DURING a session, so
    # YearView stops rather than trying to carry on. Reconnecting mid-session
    # has been the source of every recent fault; a fresh YearView against a
    # freshly started Outlook is deterministic.
    #
    # v2.1.4: do NOT try to release politely on this path. Every release calls
    # back INTO Outlook (FinalReleaseComObject unadvises the event sinks over
    # RPC), and against a half-shut-down instance those calls can block for a
    # long time - WaitForPendingFinalizers can block forever. That is the loop
    # wedging inside the very code meant to clean up: the page keeps its last
    # toast because no reply is ever written, and OUTLOOK.EXE stays pinned
    # because our release never completes, which is exactly what blocked the
    # user's next manual start. ENDING THE PROCESS drops every reference at
    # once through the OS - no RPC, no waiting, nothing to hang on - and that
    # is what lets a half-dead OUTLOOK.EXE finish exiting.
    #
    # Order is deliberate (user request): acknowledge FIRST, then take the
    # view away, so the message is never orphaned by a window that has
    # already vanished.
    if ($script:outlookGoneHandled) { return }
    $script:outlookGoneHandled = $true
    if ($NoOpen -or -not $App) {
        # Headless bridge: nobody to read a dialog. Release what we can and
        # keep serving; a later reconnect may still succeed.
        Write-Host 'Outlook is no longer available - link released (headless bridge continues).' -ForegroundColor Yellow
        Release-OutlookLink
        $script:outlookGoneHandled = $false
        return
    }
    Write-Host 'Outlook is no longer running - closing YearView.' -ForegroundColor Yellow
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.ShowInTaskbar = $false
        [void][System.Windows.Forms.MessageBox]::Show($owner,
            ("YearView will close because Outlook (classic) is not running." + [Environment]::NewLine + [Environment]::NewLine +
             "Please start Outlook and then re-start YearView."),
            'YearView',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
        $owner.Dispose()
    } catch {
        Write-Host ('Could not show the Outlook message: ' + $_.Exception.Message) -ForegroundColor Yellow
    }
    try {
        $win = Get-Process msedge -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle -like '* YearView' } |
            Select-Object -First 1
        if ($win) { [void]$win.CloseMainWindow() }
    } catch {}
    try { if ($listener) { $listener.Stop() } } catch {}
    [Environment]::Exit(0)
}
function Ensure-Outlook {
    # Serve mode connects to Outlook LAZILY (v1.2.17) so a restored page can be
    # on screen before Outlook is even touched. This function is the single
    # gate: the first call performs the full initialisation (connect, watch
    # collection, count baseline, event subscriptions); later calls return
    # instantly while the link is alive, or reconnect after Outlook restarts.
    $alive = $false
    try { if ($script:outlook) { [void]$script:outlook.Version; $alive = $true } } catch {}
    if ($alive -and $null -ne $script:watchItems) { return }
    $firstInit = ($null -eq $script:watchItems)
    if (-not $alive) {
        if ($firstInit) {
            Write-Host 'Connecting to Outlook (deferred - the cached page shows first)...' -ForegroundColor Cyan
        } else {
            Write-Host 'Outlook connection lost - reconnecting...' -ForegroundColor Yellow
        }
        Get-EventSubscriber -ErrorAction SilentlyContinue |
            Where-Object { $_.SourceIdentifier -like 'YV_*' } |
            Unregister-Event -ErrorAction SilentlyContinue
        Connect-Outlook
    }
    $script:watchItems = (Get-SelectedCalendar).Items
    if ($firstInit) {
        # Persisted fast-scan verdict (v1.2.20): the dual-reader validation
        # proves properties of THIS reader code against THIS Outlook build -
        # neither changes between sessions - yet it ran every session, and at
        # startup it rode the revalidation build, blocking the single-threaded
        # loop ~15 s and queueing any year-switch behind it. If a verdict for
        # the current (script version, Outlook version) pair is on disk, skip
        # the per-session dual read; any runtime fast-scan failure deletes the
        # verdict, forcing full re-validation next session.
        try {
            $vf = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-fastscan.ok'
            if ((Test-Path $vf) -and ((Get-Date) - (Get-Item $vf).LastWriteTime).TotalDays -lt 7) {
                # 7-day expiry: silent environmental divergence can only be
                # DETECTED by the dual read, so re-earn the verdict weekly.
                $want = ('{0}|{1}' -f $script:yvVersion, [string]$script:outlook.Version)
                if (((Get-Content $vf -ErrorAction Stop) -join '') -eq $want) {
                    $script:fastScanChecked = $true
                    Write-Host 'Fast-scan verdict restored from disk - skipping this session''s dual-reader validation.' -ForegroundColor DarkCyan
                }
            }
        } catch {}
        # Baseline the count now; deliberately do NOT touch lastChangeTicks -
        # Restore-PageCache may have stamped it to force the one silent
        # revalidation of restored pages, and the init block zeroed it anyway.
        $script:lastCount = -1
        try { $script:lastCount = [int]$script:watchItems.Count } catch {}
        $script:lastModTicks = 0L
    } else {
        # Reconnect after an Outlook restart: the world may have changed while
        # the link was down. lastCount = -1 forces a structural refresh of the
        # viewed year; other cached years stay and self-revalidate on visit
        # against this stamp (v1.2.25).
        $script:lastCount = -1
        $script:lastModTicks = 0L
        $script:lastChangeTicks = (Get-Date).Ticks
    }
    foreach ($evName in 'ItemAdd', 'ItemChange', 'ItemRemove') {
        try {
            Register-ObjectEvent -InputObject $script:watchItems -EventName $evName `
                -SourceIdentifier ("YV_" + $evName) | Out-Null
        } catch {}
    }
}
function Test-OutlookWindowed {
    # v2.1.10: TRUE only when classic Outlook is running WITH a window - the
    # state YearView requires. A visible explorer frame counts, and so does a
    # main window handle (covering hidden and tray states); a windowless
    # process - still booting, or lingering to finish a send/receive - does
    # not. Used by the startup gate and by Connect-Outlook's readiness test,
    # so the two can never drift apart.
    if (-not (Get-Process -Name outlook -ErrorAction SilentlyContinue)) { return $false }
    try {
        # v2.1.11: count HIDDEN frames too. A tray-minimised Outlook's window
        # exists but is not visible; a window the user CLOSED is destroyed.
        # This is pure Win32 - it cannot block, no matter how busy Outlook is.
        foreach ($oh in @([YVWin]::GetAllTopWindowsByClass('rctrl_renwnd32'))) {
            if (([YVWin]::GetTitle($oh)) -like '* - Outlook') { return $true }
        }
    } catch { return $false }
    return $false
}
function Show-OutlookNeededPrompt {
    # v2.1.10: the dialog now GATES startup instead of merely informing.
    # Field sequence behind this: OK was pressed without starting Outlook,
    # YearView started anyway, and the first calendar action met a
    # windowless Outlook - sometimes closing the session, sometimes leaving
    # a hidden OUTLOOK.EXE that then blocked a proper Outlook start. The
    # rule is now the one the user stated: YearView does not start unless
    # Outlook is running in window mode.
    #
    # OK re-tests and re-presents the dialog if Outlook still has no window;
    # Cancel exits YearView. After OK, a started-but-still-booting Outlook
    # (process present, window not yet) is given up to 60 seconds to show
    # its window before the dialog returns - pressing OK just after starting
    # Outlook is the normal path and must not be nagged.
    #
    # Returns $true to continue startup, $false to exit. Skipped (returns
    # $true) for -NoOutlookPrompt and -NoOpen: automation and power users
    # keep today's behaviour.
    if ($NoOutlookPrompt -or $NoOpen) { return $true }
    while ($true) {
        if (Test-OutlookWindowed) { return $true }
        $answer = $null
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            $owner = New-Object System.Windows.Forms.Form
            $owner.TopMost = $true
            $owner.ShowInTaskbar = $false
            $answer = [System.Windows.Forms.MessageBox]::Show($owner,
                ("YearView needs Outlook running to access your calendar." + [Environment]::NewLine + [Environment]::NewLine +
                 "Please start Outlook (classic), then press OK." + [Environment]::NewLine + [Environment]::NewLine +
                 "Press Cancel to close YearView."),
                'YearView',
                [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            $owner.Dispose()
        } catch {
            Write-Host ('Could not show the Outlook prompt: ' + $_.Exception.Message) -ForegroundColor Yellow
            return $true   # no dialog available: do not brick startup
        }
        if ($answer -ne [System.Windows.Forms.DialogResult]::OK) { return $false }
        # OK pressed: give a started-but-booting Outlook time to show a window
        $deadline = (Get-Date).AddSeconds(60)
        while ((Get-Date) -lt $deadline) {
            if (Test-OutlookWindowed) { return $true }
            if (-not (Get-Process -Name outlook -ErrorAction SilentlyContinue)) { break }
            Start-Sleep -Seconds 1
        }
        # still windowless (or never started): the dialog comes back
    }
}

if (-not $Serve) { Connect-Outlook }   # serve mode connects lazily (Ensure-Outlook)

$template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>__YEAR__ YearView</title>
<!-- Edge --app windows take their window-corner and taskbar icons from the page
     favicon (not the launching shortcut); without one they show a generic globe.
     Data URI: no extra bridge route to gate, and static snapshots get it too. -->
<link rel="icon" type="image/x-icon" href="data:image/x-icon;base64,AAABAAQAEBAAAAAAIAAsAgAARgAAABgYAAAAACAAvAIAAHICAAAgIAAAAAAgAFoDAAAuBQAAMDAAAAAAIACPCAAAiAgAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAAfNJREFUeJyNkrFrU1EUh79z730vTWKSxmKDcbC2XcQqtmDp0EmxiwhOjoJTBwcnd3HQ/0HwD9DFuVARRURxqgVBwVC0UNRWzEti2rx373FoKQ7t09/+/c7HOUcAAfRp7eRwo1I2CfkpG6v14GVq/cNP9mCeNydvlsU86KuiYASVwyvEF8WYbfWPfZTekWeNiTNx5N7HYkyGkkPuR4GiGNpZetlZwrj53TeJTwO5k/9yQDQYi4J1duHqysSNa+1quV6zYhXJlxCETD2D0Jfv336o27h7uz03Pj74pB/Z6K0T2wKqejAsQhpSjsWjnI3maG9t4forr5xpjsny+hNeb7ygElXx6g8ssGLpZl2mhy8wdfoeO0mCa5QamolnoXmF8/VpIhvnGmQh5Wg0iliLswUcVAl4ar0Z6Mxgbf4CQ4ByCaiABVwCOCO8ayWsft6kWHAcIoAIbA+UyeYQ1xsNPOCqgA/KuVMVTozERNZwCI8AWYBaafdQFnCQENQxduQtY3wBE0NehaYQHwfmgQy35xa0+5Xu5irGDuUWaNjBlbcZqs0HES/OWitoVpX6RS1V5vQff4SiWFdUwKjG3rVa9WR2dvCoWB25ZaOR/3plQDq93nKn8+vNPrC2tnZJIrE+213OQfFAwTlN0zTcX1p6+XBxMf0DX33AxmY0R+wAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAGAAAABgIBgAAAOB3PfgAAAKDSURBVHictZW9a5RBEIef2d37SO5yyXkHGmMRNTFySakIUSSFkF5Io4Xg32BvZSeW9oIgGNBaCzUIKWysDGpIIRH8AKM575Lz/dix2JzJ2d17urDFLru/Z2ZnZgeAm2D4T0MUjIB/cWRqqWjs9QhvFCSLmIIWEYm8rlz8vH5LQQTg5cTUUlncQ4+iA1qsQFkMX31y5+LH9zcEYGV8+lnJmoWWT2NB7KAMZ0TiOGm1vnw77gDESIwxYsASZnZ1ESVJxR4+HE3euysOIGm3xRtD4v2AxoOKwG4Hc2JSKouLOIDTy/cpDw8Rqwf13dBkUFdAsAi/nGPjzWYAjF5aYHhg23tHBDSXnwRA0myipRIbrXest9bImwKq/eWTiBD5X0yXG5wszxB3OlCpBIAXEGtZb7/l0acHjLhRvKZ9AYxYfibbXD56lanRRqguCICDh3KSx0keT/+AnOQxf2W5A7AmbJ6rn2emOoPNmKkpKWO2vgfs8SAs8kmdkU4dI/Rd0QJ4hXwRyO1rOoA0DXJrm01W15oM5y2+T4IR2IlS5hsVzpyq0C2pnhiogveC9/QNQMLdv5PPBXpw51i9yIU5h7PZnihJYaIWbO7W6t5fFBbjNcN4ze0dz4JQuq2lB4APKRl9f02ytYrYofBl9KVv0HQXd2iefPXsn/s9MRDfgXgLKGUCELeDxoERnkjEA0iuhh05jdhiRg86SK7W3dF9D8QMAamrNNRVGv+iqXnEFKy1IbxxFD8lNJpuhAeZBrBJ4led29kWVTUioj9+NG8XioUroAZFFO27KQjixYhEUfLq4+aHa7Ozs996RB4/fz42VCgYmhkeBiiVUrXVqszPzW0BqKr8BrwV5/mNiVTUAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAADIUlEQVR4nL2Xz2sdVRTHP+fc+ZGXJrFBzas/0iJdFKKoRYSQhWThQrAiFSouXLv2H5BQcCFaXIsUdCUSF+JOkO5FlCpoqK5KxEBa25eEPGfezJ17XLz3yoP3bILc1y9cZrhzOd/PvffMmTswkG1sKA9Im5cuueG9ABiIgAHy0/K5F6rAXIkniWTogYSENIRm26fX37y9dTj0FNvYUC5ftp8/+Hil/vTzq1YWq6qK2YAugmxwFaCB7bC+9u7aF599bYaKgX61spK0C/1+0Yfzd33VIBLLe4ylJeqaLC3rplp5afuPmyoQlg6K52bK3vm7vvIi4gR0Ss39Y00926tmpKovClgCkJrMq6phYeqJqCA4Z2TJPNDPMxFtEBUxTET66TEteY/kubQuvtrw0VYfoCkKcaGHtwB1jTVhChAGCJIkhADB+gYJwKmrn9BeepQGo7v1O3Wng7gEzO4b8tgSwAxJEuaef5Ysy9ifbcGV9/sAM+vrzJ+cA+Dk2moc0yNUd7vAYAXY28PmWwhGP02mlIsWCBYQEUJZjgAkCeL61XGn+JNv/vqSVNOo3nWoef2Jt3i8tdzvUB0BGFHXH/LbwXVybWHEyQFB6IWCl9sXxp6NAagomeZkmkUFMAI6YWsnAsy6E2Q6ExXAiTsewJkTZ3nv6Q+J9ykaykh15mgAMcWF2cjmg9jC2LzGAHY7Fd/+2CF1Eq8OCdSN8cqLi5xazO4PYAa92jCLWAgFKm8T400E8N6G1TMagD8uQJ4JZ9opSeQt8I2Rp+OJPQbw8ELKhdVH4jgfQ+NbUO/jD34FcZPG/39ZQ7LwDJI+dBTAHv7va6A5RCpEIBB6uNaTRwMgDknmQLPIAOnEVZ1w9A/QlGAhMkDVj/0fADJ0k7xNtvw2oHEBCEjeBsBGXq9k0NEMR4nmSOt0JONxhRBQVWMwOzUzuXVr55eiKHcBB1QQ/HQataoqIAfd7nfD5XAAOzu7b1RVvW9TVhOC3b5z5x2Azc1NJwMIERG7cePmU0tLi68lWbpQlqVorKNhAFRR8J3D4trZ04/9MPS8N8bMHtjv+ajXv50I0iaYHOK7AAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAIVklEQVR4nM2aTYxkVRXHf+fe+159dE/3DPQMQ+MAMzbNhwiLTjQhCIiABFEUQy/cuJPExARMXJk4gyuXbIzRvZtBEyWCATaAEDFhQgIBw4D2zDgfOsMw1V3d9fHeu/e4eNU1VdWvuqvaSM8/qfSrV/fdd/7n3PN1bwtDoGAAeW7YgM8ITwDPAYsQAN3ygcNgFOT/Ldh2cBTs4L0+QQ+DeSZnyrH9N92q1uxfJUOwO0JIES0DRnT1hdPHjz0DQcFIR0boIaAgAvrK7MH5CdyzIA86EafsnDnW3+1RUN5rqv/J/ef+8VIvCYFc80dA3zwwd8gG87dYzNX14KGH6U5CQSoiEmFYIXviq2c+/t1RsIvgpTPAAPrq7E0vTYt9cEWzRJB4h+UegGYOYwN8UlFuWzh3/CKAOdwxx5uzN8/HyP119eHKEx5AXErwk8bsXYNvC+ir3GvNfbn2UeGOkojVUULVDkHyjwI3AuxiVcwuFgTAh/BFJ4KgVyyBDkSFdP2L6V6I+J2RZ3yIajcwur5fjEHyBJwHVblC8plq/kHBGHoXSZdASFMXWi3SkCEImnk0y3aehCpiDRJFKEowjtBLoH74UeWZY1Sfevqdufu+RAAJQWmfOUPzxClMFAF6OauM9fK+P13IhothENRnRHv2MHHrzSDgEE79+wJ895vw6x+AqlqAd0+dv1dzeL1ykaqqLv3n4hGAt99+O+ouIb1wwYXZqzAiLJ2rsdpIADh47TTVUjS61mVdsYIPgY9PXyLzobsSVcFZw9zn9mCN2WSirmSdZSwIigJhdbX762UC1qrYvNirt1I+XU0QgQMIJnKDs27+ys4aVeBSIyXN+glEzqDGgB2FQA/E5MrpIV4o2Z6JmNiAGMGnKaeXLtGj2k0Fd86yZ+80QZUQlL1TMV61+6gCVoR2OyE1ozmVqmKtpVzaWCAUEqi3PBdX2kxP7Ob5vz7PG+0/MFmZIoRNUoUKxgn1T9f48cLPuH3+C2Q+UGssF1rg0IEKI8q/KQoJCAFLQDVhurKbmbVZyskEqpsUpwrihQoJzkS02m2yLCDqMfRbQFRpNlvYUdsMBWMMcbzRFwsJKALGstZu8tCXv87j5hsEHSElKIgFjSDLMnCSO11Pi6ediGysKXZiyd11g0RDXl5MQCFoPk0WMrS6oZPbhIMiGKLIEQE2KhEISEdWDWAjQ6VgPW8HhQScUSITEBHSJKHZzEltbQFFjGG1vkyz2USB2oUa3veaAKwVTkZrA0tICMEzNbWbuFRG9bLWx3bixEMrE0SgVC5TrYyureXaJX75278gpRlQjw/FfvP6Oxf7FCIC7USZnW7xw+8/TEH/XohCArvKDtE8cmRpQqMZtraAdpIMhjvm99FOFDERIRRX50b6w7IAaeaZ3XsNjUYbRbZvgXYWWGt7jIC1jnIpLiRgCpxwolrh8Ue+0v3+wYlPCsPobTfObKKN0VHsxGLBxCjg4hJmSMDOvC/o37Sr9cwH6mudcNpbSjhDo9nGDWRiJVeSkUHF5P5g7cZlVUigduE0Z8+exTkH9QnKce8wQdVTKle54YY5xJihCVpRIisIMlAL5VWNFjR/quDxA/OAMVsQyDojM4XzH/6JAxOnICvjz3nWtHcyIXKBf54Xrp75OXtndvdFjF5YH8iCkHr6CCBCKY43WGA7cH0XAk7g4J3f4fz5czjnqJYjXO9aF9AQmL+xSqUSs9ZoDrVAUGXPpCME2xdGjRFarVbuyCNAYTwnjqeuZZfdB6LcOXc18TYVFRQunayRphud+OBsaSwLDNtpKCQQ0hZZu4mIUKtFTFaiDR2ZIMjQaiy/H4JS3VUlHewHnMG60bP75RlHJKBi8qJGBOcccVQcRocj11fLJ7z1+z/SulTHOJvf7fTlf5+qDK1v+gQX8JmnetUUD3/vMSIX9e13blkLWWtxzo5FYN2ps2aDj15+DXOxAZHt2004OeLukxhBkwxmJrnn8YeYnoygJ7tvWQslSZuG0TEtAKjiTMS3fvFTGj0Oq4AzhkPX7R6aX/oIACEEJqpVpiengP4EunUtVBqvFupFGgJRPMmEndjgxNfv37etOQcxVi1UiWOM176lsBlsUPbFlixoXz8QOUPSzkh9uunz61BVnLWUy6XRCCReaSaKCERRTKUSIQgvvPEav3rrz0xVJ4dWmYNIs/5x1hhWGqv86O5HePTu+/DBY814EWlLAmnmaacZIpBmGVlmEITMZ2SZJ/Me70choCRp/9mcMZYk7cyTeXzw6JbpYMxayAo40zlzCoHg8yX02D1f47G77h95CRGUD5Y+yfPAuigKcWSZn5sh8xkA3m++rzxSLdTZ+crDX0/94qKIUqecVhSNpKBnLUbqA3U8SRjIxEEJKKVttpXaUwV2CVhr3Xqom646nMSICD5LaTTCtrZGgypXTXSamoFaqN1ukybj10KqijHS3Z5wnd9ZXl4722i20mql7JabGbV67gOz15SoVsbYWuwjALVTK4UNzecPVLezLyQiQpb4JYCFhQV1IhJUVZ588jfHbzi4/+PqdftviQxZyakTgaTdpmVGaCkL4L0SWcUw2NSPvy8EhHK5RG253jp+fOmF9Vfkp5SqVkT8u+9/+MD83KFXSrEjhNzD8mi5/VOnYdl2WK9cDME56wA+/OjE07fMH3x2XebuENU8mL37/keLteWVM5/djvloWKmvriyd/NdTvbLm1HqgqkZEwosvvr731ttvejRy9vpGq4HI//KvBsPyxda9gKrXSqki7SQ9eeLcqZcfuOuuM6oqIjLcfEc7Bx5XIo4ePbpBtg2JbFHEa34KaI8dO3ZFnPItLCzokSNHwuLi4oaM91/n+BV7KaS68AAAAABJRU5ErkJggg==">
<style>
  :root{
    --ink:#1a1d21; --muted:#7a8088; --grid:#d7dade; --gridlight:#e9ebee;
    --weekend:#c6d3e4; --daynum:#8a9099; --paper:#ffffff; --band:#f6f7f9;
    --today:#c11a25;
  }
  *{box-sizing:border-box}
  html,body{margin:0;padding:0;background:var(--paper);color:var(--ink);
    font:12px/1.25 "Segoe UI", system-ui, -apple-system, sans-serif;}
  .wrap{padding:14px 18px 0}
  header.yv{display:flex;align-items:baseline;gap:14px;margin:2px 0 10px}
  header.yv h1{font-size:26px;font-weight:600;letter-spacing:.5px;margin:0;display:flex;
    align-items:center;gap:8px}
  header.yv .yrnav{font:inherit;font-size:16px;line-height:1;padding:2px 9px 4px;border:1px solid var(--grid);
    background:#fff;border-radius:4px;cursor:pointer;color:var(--ink)}
  header.yv .yrnav:hover{background:var(--band)}
  header.yv .sub{color:var(--muted);font-size:11px}
  header.yv .actions{margin-left:auto;display:flex;gap:8px}
  header.yv button{font:inherit;font-size:11px;padding:4px 10px;border:1px solid var(--grid);
    background:#fff;border-radius:4px;cursor:pointer}
  header.yv button:hover{background:var(--band)}

  /* weekday header */
  .dowhead{display:grid;grid-template-columns:44px repeat(37,minmax(0,1fr));
    position:sticky;top:var(--headh,44px);background:var(--paper);z-index:5;
    border-bottom:1px solid var(--grid)}
  .dowhead .cell{font-size:9.5px;color:var(--muted);text-align:center;padding:3px 0 2px;
    border-left:1px solid var(--gridlight);overflow:hidden}
  .dowhead .cell.we{background:var(--weekend)}
  .dowhead .corner{border:none}

  /* v3.1.0: the timeline is a sequence of year blocks, each 12 month rows.
     The band names the year the rows below it belong to - without it the
     January after a December is indistinguishable from any other January. */
  .ysep{padding:6px 0 5px 48px;background:var(--band);color:var(--ink);
    font-weight:600;font-size:13px;letter-spacing:.5px;
    border-top:1px solid var(--grid);border-bottom:1px solid var(--grid)}
  .yearblock:first-child .ysep{border-top:none}
  /* v3.1.6: a timeline runs onto further sheets rather than being squeezed
     onto one - so a month row must not be split down the middle by a page
     break, and a year band must not be left stranded at the foot of a page. */
  @media print{
    .month{break-inside:avoid; page-break-inside:avoid}
    .ysep{break-after:avoid; page-break-after:avoid}
  }

  /* month rows */
  .month{display:grid;grid-template-columns:44px repeat(37,minmax(0,1fr));
    border-bottom:1px solid var(--grid);position:relative}
  .mlabel{grid-column:1;display:flex;align-items:flex-start;justify-content:flex-start;
    padding:4px 6px;font-weight:600;font-size:12px;color:var(--ink);
    border-right:1px solid var(--grid);background:var(--band)}
  .mgrid{grid-column:2 / -1;display:grid;grid-template-columns:repeat(37,minmax(0,1fr));
    grid-auto-rows:min-content;position:relative}
  .day{grid-row:1 / var(--rows);border-left:1px solid var(--gridlight);min-height:100%;}
  .day.we{background:var(--weekend)}
  .day.out{background:#eceef0}
  .day.today{}
  /* Today marker is a separate click-transparent overlay. Putting z-index on the
     day cell itself lifted that full-height background cell above the event bars,
     which silently swallowed clicks on any bar inside the red box. */
  .todayov{grid-row:1 / var(--rows);outline:2px solid var(--today);outline-offset:-2px;
    z-index:3;pointer-events:none;min-height:100%}
  .dn{grid-row:1;font-size:9px;color:var(--daynum);padding:1px 0 1px 3px;pointer-events:none;
    min-height:13px}
  .bar{height:14px;margin:1px 1px 0 1px;border-radius:2px;font-size:9.5px;line-height:12px;
    padding:0 3px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;cursor:pointer;
    border:1px solid rgba(0,0,0,.18)}
  .bar.dark{color:#fff}
  /* uncategorised: hollow - no fill, just an outline and dark text */
  .bar.nocat{background:#fff;border:1px solid var(--muted);color:var(--ink)}
  .bar.light{color:#20242a}
  /* v3.4.7: a category whose titles are switched off in the Filter panel - the
     bar keeps its colour, size, outline and tooltip; only the text goes.
     Applies on paper too, as on screen. */
  .bar.notitle{color:transparent}
  /* v3.4.8: set only for the duration of a print when "Print category legend"
     is off. A class rather than a print-media rule, because the fit-to-pages
     scale is measured at beforeprint, before print rules are in force. */
  html.nolegend .legend{display:none !important}
  .day.rangesel{background:#ffe3a3 !important;outline:1px solid #d99e2b;outline-offset:-1px}
  body.armed{user-select:none;-webkit-user-select:none}
  .spacer{height:15px}
  .padrow{height:3px;grid-column:1 / -1}

  /* legend */
  .legend{display:flex;flex-wrap:wrap;gap:6px 14px;margin:12px 0 0;align-items:center}
  .legend .item{display:flex;gap:5px;align-items:center;font-size:10.5px;color:var(--ink);
    cursor:pointer;padding:2px 7px;border-radius:10px;border:1px solid transparent;user-select:none}
  .legend .item:hover{background:var(--band);border-color:var(--grid)}
  .legend .item.sel{background:#fff;border-color:var(--today);box-shadow:0 0 0 2px var(--today);font-weight:600}
  .legend .sw{width:12px;height:12px;border-radius:2px;border:1px solid rgba(0,0,0,.2)}
  /* v2.0.20: "(No category)" swatch echoes the hollow uncategorised bar */
  .legend .sw.none{background:#fff;border:1px solid var(--muted)}
  .legend .item.newcat{border:1px dashed var(--muted);color:var(--muted)}
  .legend .item.newcat:hover{color:var(--ink);border-color:var(--ink);background:#fff}
  .legend .item.clearcat{display:none;border:2px solid var(--today);color:var(--today);
    background:#fff;font-weight:600}
  .legend .item.clearcat:hover{background:#fbe0e0;border-color:var(--today)}
  body.armed .legend .item.clearcat{display:flex}
  .legend .hint{font-size:10px;color:var(--muted);margin-left:4px}
  body.armed .day.in{cursor:copy}
  .toast{position:fixed;right:16px;bottom:16px;background:#20242a;color:#fff;font-size:11.5px;
    padding:8px 14px;border-radius:6px;opacity:0;transition:opacity .25s;pointer-events:none;z-index:20}
  .toast.show{opacity:.95}
  .ctx{position:fixed;background:#fff;border:1px solid var(--grid);border-radius:6px;
    box-shadow:0 4px 14px rgba(0,0,0,.18);z-index:30;min-width:170px;padding:4px;display:none}
  .ctx div{padding:6px 10px;font-size:11.5px;border-radius:4px;cursor:pointer;white-space:nowrap}
  .ctx div:hover{background:var(--band)}
  .ctx div.danger:hover{background:#fbe0e0;color:#a11616}
  .legend .item.off{opacity:.35}
  .legend .item.filtercat{border:1px solid var(--grid);background:#fff}
  /* filter in use: same red ring as a selected category */
.legend .item.filtercat.active{border-color:var(--today);box-shadow:0 0 0 2px var(--today);font-weight:600}
  .fpanel{padding:8px;max-height:60vh;overflow:auto}
  .fpanel .frow{display:flex;align-items:center;gap:6px;font-size:11.5px;padding:4px 6px;
    border-radius:4px;cursor:pointer;user-select:none}
  .fpanel .frow:hover{background:var(--band)}
  .fpanel .fbtns{display:flex;gap:10px;padding:2px 6px 8px;border-bottom:1px solid var(--gridlight);
    margin-bottom:5px}
  .fpanel .fbtns span{font-size:11px;color:#2858A5;cursor:pointer}
  .fpanel .fbtns span:hover{text-decoration:underline}
  .fpanel .sw{width:12px;height:12px;border-radius:2px;border:1px solid rgba(0,0,0,.2)}
  /* v3.4.7: the filter panel's second column - titles, per category */
  .fpanel .fhead{display:flex;justify-content:space-between;font-size:10.5px;color:var(--muted);
    padding:0 6px 3px}
  .fpanel .fhead .fth{width:44px;text-align:center}
  .fpanel .fbtns .tfirst{margin-left:auto}
  .fpanel .fcat{padding:0;gap:0;cursor:default}
  .fpanel .fcat .fshow{display:flex;align-items:center;gap:6px;flex:1;padding:4px 6px;cursor:pointer}
  .fpanel .fcat .ftitle{display:flex;align-items:center;justify-content:center;width:44px;
    align-self:stretch;cursor:pointer;border-left:1px solid var(--gridlight)}
  .helpov{display:none;position:fixed;inset:0;background:rgba(20,22,26,.45);z-index:40;
    align-items:center;justify-content:center}
  .helpov.show{display:flex}
  .helpbox{background:#fff;border-radius:8px;box-shadow:0 10px 34px rgba(0,0,0,.3);
    max-width:560px;width:92%;max-height:82vh;overflow:auto;padding:14px 18px 16px;font-size:12px}
  .helphead{font-size:14px;font-weight:600;display:flex;align-items:center;margin-bottom:8px}
  .helpx{margin-left:auto;cursor:pointer;color:var(--muted);padding:2px 6px;border-radius:4px}
  .helpx:hover{background:var(--band);color:var(--ink)}
  .helpbox dl{margin:0}
  .helpbox dt{font-weight:600;margin-top:8px}
  .helpbox dd{margin:2px 0 0 0;color:#3a3f45;line-height:1.45}
  /* v2.1.32: in-page confirmation, so the heading is ours rather than the
     browser's "127.0.0.1:8377 says". Reuses the help overlay's box. */
  .askbox{max-width:380px}
  .askmsg{color:#3a3f45;line-height:1.5;margin:2px 0 14px 0;white-space:pre-line}
  .askbtns{display:flex;justify-content:flex-end;gap:8px}
  .askbtns button{font:inherit;padding:5px 14px;border-radius:5px;cursor:pointer;
    border:1px solid var(--gridlight);background:#fff;color:var(--ink)}
  .askbtns button:hover{background:var(--band)}
  .askbtns button.danger{border-color:#c11a25;background:#c11a25;color:#fff}
  .askbtns button.danger:hover{background:#a5141d}
  .credit{text-align:right;color:var(--muted);font-size:10px;margin-top:10px}

  /* v3.1.3: Chromium adjusts the scroll position ITSELF when content changes
     above the reading position (scroll anchoring). Since v3.1.2 keeps the
     surrounding blocks alive rather than rebuilding them, it now finds an
     anchor and compensates - on top of our own compensation, which is what
     made a year's arrival jump by exactly one year. Turn the heuristic off so
     that ONE mechanism is in charge, and let that one measure its own result. */
  html,body,.wrap,#months{overflow-anchor:none}

  /* v3.1.1: the controls stay put while the timeline moves. The document is
     still the scroller - sticky, not a fixed shell - so the year-loading,
     scroll-anchoring and print-scaling code all keep working unchanged. */
  header.yv{position:sticky;top:0;z-index:6;background:var(--paper);
    margin:2px 0 0;padding-bottom:8px}
  .legend{position:sticky;bottom:0;z-index:6;background:var(--paper);
    margin:0;padding:9px 0 8px;border-top:1px solid var(--grid)}
  /* keep the toast clear of the pinned legend */
  .toast{bottom:calc(var(--footh,34px) + 14px)}

  /* v3.3.2: NO size declaration. This browser ignores one - named sizes and
     explicit lengths alike were both tried, and in each case the dialog went
     on offering a LIVE Layout control, honoured its own Portrait setting and
     printed on a portrait sheet. Continuing to declare a size only meant the
     page scaled itself for a sheet the printer was never going to use, which
     is worse than not declaring one: a year that fits a landscape sheet ran
     onto two portrait ones. So the dialog owns paper and orientation, and
     the setting tells YearView only what to SCALE for. @page stays for the
     margin, at the top level where it belongs. */
  @page{margin:8mm}

  @media print{
    /* nothing is pinned on paper - every row prints in flow */
    header.yv{position:static;padding-bottom:0}
    .legend{position:static;border-top:none}
    header.yv .actions{display:none}
    .yrnav,.toast,.ctx,.fpanel,.helpov,.credit{display:none !important}
    .legend .newcat,.legend .clearcat,.legend .filtercat,.legend .hint{display:none !important}
    .dowhead{position:static}
    body{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  }
</style>
</head>
<body>
<div class="wrap">
  <header class="yv">
    <h1><button class="yrnav" id="prevYr" title="Previous year">&#8249;</button><span id="yrword">Year </span><span id="yrlabel">__YEAR__</span><button class="yrnav" id="nextYr" title="Next year">&#8250;</button></h1>
    <span class="sub" id="stamp"></span>
    <div class="actions">
      <button id="calBtn" style="display:none" title="Choose which calendar to display">Calendar</button>
      <button id="refreshBtn" style="display:none" title="Re-read the calendar from Outlook">Refresh</button>
      <button onclick="window.print()">Print</button>
      <button id="setBtn" title="Settings">&#9881;</button>
      <button id="helpBtn" title="How to use YearView">?</button>
      <button id="quitBtn" style="display:none" title="Stop and close YearView">Quit</button>
    </div>
  </header>
  <div class="dowhead" id="dowhead"></div>
  <div id="months"></div>
  <div class="legend" id="legend"></div>
  <div class="toast" id="toast"></div>
  <div class="ctx" id="ctx"></div>
  <div class="helpov" id="helpov">
    <div class="helpbox">
      <div class="helphead">YearView &#8212; quick help<span class="helpx" id="helpx">&#10005;</span></div>
      <dl>
        <dt>Create an item</dt><dd>Click a category in the legend &#8212; or <b>(No category)</b> for an item with none &#8212; then click a day, or drag across days for a multi-day event. The Outlook item opens with the category set. The selection stays armed for repeat adds; press Esc or click <b>&#10005; Deselect</b> when done.</dd>
        <dt>Open an item</dt><dd>Click its bar, or right-click it &#8594; <b>Open in Outlook</b>.</dd>
        <dt>Delete an item</dt><dd>Right-click its bar &#8594; <b>Delete</b>. A recurring item offers <b>Delete this occurrence</b> or <b>Delete entire series</b> instead. Deleted items go to Outlook's Deleted Items. YearView asks you to confirm first; you can turn that off in <b>&#9881; Settings</b>.</dd>
        <dt>Right-click a day</dt><dd>Right-click a blank part of any day of the month &#8594; <b>Open Outlook calendar here</b>, which switches Outlook to that date in the calendar YearView is showing.</dd>
        <dt>Filter categories</dt><dd>Click <b>&#9703; Filter</b> to choose which categories are shown. Hidden categories dim in the legend. The <b>Titles</b> column chooses which categories show item titles: untick one and its items become plain coloured bars, with the title still shown when you hover.</dd>
        <dt>New category</dt><dd>Click <b>+ New category&#8230;</b>, pick an existing Outlook category or type a name to create one, then click or drag days.</dd>
        <dt>Change year</dt><dd>Use the &#8249; &#8250; buttons beside the title. Previously opened years come back instantly; a year never built before shows a short progress message while it is read.</dd>
        <dt>Calendar</dt><dd><b>Calendar</b> lists your calendars (default store) and switches the whole view &#8212; including creation, deletion and live updates &#8212; to the one you pick. The choice persists across restarts.</dd>
        <dt>Settings</dt><dd>Click <b>&#9881;</b> to change how YearView behaves: <b>Confirm before deleting</b>, on by default; <b>Print &#8212; fit to</b>, one page by default; <b>Print category legend</b>, on by default &#8212; turn it off to leave the row of categories off the printout; and <b>Year starts at</b>, January by default &#8212; set it to <b>Current month</b> or a named month and the view shows twelve months from there, running on into the following year. Other years always show January to December. Choices are remembered across restarts.</dd>
        <dt>Refresh</dt><dd>The view updates itself a few seconds after calendar changes; <b>Refresh</b> forces it.</dd>
        <dt>Zoom</dt><dd>Ctrl + and Ctrl - (or Ctrl + mouse wheel). F11 toggles full screen.</dd>
        <dt>Print</dt><dd>Scales what is on the page onto the number of sheets set by <b>Print &#8212; fit to</b> in <b>&#9881; Settings</b> &#8212; one page by default. Choose <b>As many as needed</b> to print at full size and let it run onto further sheets, which is usually what you want with continuous scrolling on. Turn off <b>Print category legend</b> in <b>&#9881; Settings</b> to print without the row of categories at the foot. Paper and orientation are chosen in the print dialog as usual. The page count assumes A4 landscape, since a page cannot ask the dialog what was chosen; on other paper, use the dialog's own <b>Scale</b> setting instead, or print to PDF and enlarge.</dd>
        <dt>Quit</dt><dd>Closes this window and stops YearView completely, including the part that keeps running when the window is closed. Restart from the YearView shortcut. Closing the window does the same thing a few seconds later.</dd>
      </dl>
    </div>
  </div>
</div>

<script id="yvdata" type="application/json">
__DATA__
</script>

<script id="yvdata2" type="application/json">
__DATA2__
</script>

<script>
"use strict";
/* ---------- pure layout core (testable) ---------- */
const YV = (() => {
  const DAY = 86400000;
  const dOnly = s => { const d = new Date(s); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); };

  // Clip an event to [mStart, mEndExcl) and return {d0,d1} day indices within month, or null.
  function clipToMonth(ev, mStart, mEndExcl){
    let s = dOnly(ev.s);
    let eDate = new Date(ev.e);
    let e;
    if (ev.ad){
      e = dOnly(ev.e);                    // all-day: end is exclusive midnight
      if (e <= dOnly(ev.s)) e = new Date(dOnly(ev.s).getTime() + DAY);
    } else {
      e = dOnly(ev.e);
      // timed event ending exactly at midnight shouldn't bleed into next day
      if (eDate.getTime() === e.getTime() && e > dOnly(ev.s)) { /* ends 00:00 */ }
      else e = new Date(e.getTime() + DAY); // make exclusive
      if (e <= s) e = new Date(s.getTime() + DAY);
    }
    // F-1: a timestamp the Date parser rejects (e.g. culture-mangled input)
    // must drop the event, not auto-flow a NaN-positioned ghost bar.
    if (isNaN(s.getTime()) || isNaN(e.getTime())) return null;
    const cs = Math.max(s.getTime(), mStart.getTime());
    const ce = Math.min(e.getTime(), mEndExcl.getTime());
    if (ce <= cs) return null;
    return { d0: Math.round((cs - mStart.getTime())/DAY), d1: Math.round((ce - mStart.getTime())/DAY) - 1 };
  }

  // Greedy lane assignment; events: [{d0,d1,...}] sorted by (d0, -(d1-d0))
  function assignLanes(evs){
    const laneEnds = [];
    for (const ev of evs){
      let lane = laneEnds.findIndex(end => end < ev.d0);
      if (lane === -1){ lane = laneEnds.length; laneEnds.push(ev.d1); }
      else laneEnds[lane] = ev.d1;
      ev.lane = lane;
    }
    return laneEnds.length;
  }

  function monthLayout(year, month /*0-11*/, events){
    const mStart = new Date(year, month, 1);
    const mEndExcl = new Date(year, month + 1, 1);
    const nDays = Math.round((mEndExcl - mStart)/DAY);
    const offset = mStart.getDay();            // 0 = Sunday, matches column 0
    const clipped = [];
    for (const ev of events){
      const c = clipToMonth(ev, mStart, mEndExcl);
      if (c) clipped.push(Object.assign({}, ev, c));
    }
    // Bar order within a day (v1.2.12): all-day events at the top, then timed
    // events earliest to latest; longer spans take the upper lanes within each
    // group so multi-day bars read as continuous ribbons; subject is only the
    // final stable tie-break (it was previously the SOLE same-day key, which
    // made bar order look random).
    clipped.sort((a,b) =>
      a.d0 - b.d0 ||
      ((b.ad ? 1 : 0) - (a.ad ? 1 : 0)) ||
      ((b.d1 - b.d0) - (a.d1 - a.d0)) ||
      (a.s < b.s ? -1 : a.s > b.s ? 1 : 0) ||
      String(a.t).localeCompare(String(b.t)));
    const lanes = assignLanes(clipped);
    return { nDays, offset, events: clipped, lanes };
  }

  return { clipToMonth, assignLanes, monthLayout, dOnly };
})();
if (typeof module !== "undefined") module.exports = YV;   // node test hook
/* ---------- rendering ---------- */
(function render(){
  if (typeof document === "undefined") return;
  const data = JSON.parse(document.getElementById("yvdata").textContent);
  data.events = [].concat(data.events || []);
  data.categories = [].concat(data.categories || []);
  const YEAR = data.year;                   // the year this PAGE was built for
  // v3.1.0: the page holds a MAP of loaded years, not one year. The home year
  // is seeded from the embedded payload; neighbours arrive from /data. Note
  // that YV.monthLayout clips events to a month by ABSOLUTE date, so one
  // combined array can be handed to every month of every year - no
  // per-year partitioning is needed at render time.
  const years = new Map([[YEAR, data]]);
  let allEvents = data.events;
  let focusYear = YEAR;                     // whichever year dominates the viewport
  const yearsAsc = () => [...years.keys()].sort((a, b) => a - b);
  const BUILD_TS = "__TICKS__";   // server build time (ticks, as string); "0" in static snapshots
  const TOK = "__TOKEN__";        // per-session token, sent as the X-YV-Token header on every bridge call (empty in static snapshots)
  const HIDE_SEED = __HIDE__;
  const CONFIRM_DELETE = __CONFIRMDEL__;
  const SCROLL_YEARS = __SCROLLYEARS__;     // v2.0.8: server substitutes a COMPLETE JSON string literal (safe for any category content)
  const PRINT_PAGES = __PRINTPAGES__;       // v3.2.0: sheets to scale a print onto; 0 = as many as it takes
  const PRINT_LEGEND = __PRINTLEGEND__;     // v3.4.8: false = leave the category legend off the printout
  let printLegend = PRINT_LEGEND;           // v3.4.8: state directly beneath its literal
  const NO_TITLES = __NOTITLES__;           // v3.4.7: csv of categories drawn without titles
  // v3.4.7: state directly beneath its literal and above the first render.
  // Category KEYS, with "" for (No category), exactly as the filter's hidden
  // set holds them - and parsed inline, because the shared csv helper is
  // declared further down and would still be in its dead zone here.
  const noTitles = new Set(String(NO_TITLES).split(",").filter(Boolean)
    .map(s => s === "__none__" ? "" : s));
  const START_MONTH = __STARTMONTH__;       // v3.4.0: -1 = follow the current month, 0-11 = fixed
  // v3.4.0: declared HERE, immediately below the literal it reads, and not up
  // with the years map where it would sit ABOVE that declaration and die in
  // its dead zone - which is exactly what the first attempt did.
  //
  // The offset applies ONLY to the year this page was built for, and only
  // while that is the current calendar year. Any other year shows whole,
  // January to December - so stepping forward or back always lands on a plain
  // year, which is what was asked for and also what keeps the arithmetic
  // simple: at most two years are ever in view.
  let startMonthPref = START_MONTH;
  function resolveStartMonth(){
    const now = new Date();
    if (YEAR !== now.getFullYear()) return 0;
    if (startMonthPref === -1) return now.getMonth();
    return (startMonthPref > 0 && startMonthPref <= 11) ? startMonthPref : 0;
  }
  let startMonth = resolveStartMonth();
  // v3.4.2: the spill year, if the page was built carrying it. Seeded BEFORE
  // the first render so those months are part of the first paint rather than
  // appearing a moment later. When it is absent - the setting was changed
  // since this page was built, or the pre-read failed - ensureSpillYear()
  // still fetches it, so this is an optimisation and not a dependency.
  const DATA2 = (function(){
    const el = document.getElementById("yvdata2");
    if (!el) return null;
    try { return JSON.parse(el.textContent); } catch(_){ return null; }
  })();
  if (DATA2 && DATA2.events && startMonth > 0){
    DATA2.events = [].concat(DATA2.events || []);
    DATA2.categories = [].concat(DATA2.categories || []);
    years.set(YEAR + 1, DATA2);
    recombine();
  }
  // v3.1.5: declared HERE, not with the other settings much further down,
  // because the first render reads it and runs long before that point - the
  // settings block's declaration would still have been in its dead zone. It
  // derives from the ONE seeded literal above: Restore-PageCache rewrites
  // that declaration in cached pages from yv-settings.cfg, and a second
  // literal of its own would not be rewritten (v2.1.24 - one authority,
  // overwrite the copy). The normaliser matches on the declaration text by
  // FIRST occurrence, so that phrase must not be repeated here either.
  // v3.3.6: pinned OFF while the setting is out of the panel. The seeded
  // value is deliberately ignored rather than deleted: an existing
  // yv-settings.cfg may still say the feature was on, and honouring that
  // would leave a reader in a continuous timeline with no control anywhere
  // to leave it. The stored value is left untouched for when the setting
  // comes back.
  let continuous = false;
  const HIDE_SET = __HIDESET__;   // false only when yv-filter.cfg has never been written
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const DOW = ["Su","Mo","Tu","We","Th","Fr","Sa"];
  const COLS = 37;
  const today = new Date(); let todayKey = [today.getFullYear(), today.getMonth(), today.getDate()].join("-");
  // v2.1.16: keep the red box on TODAY. The box is stamped at render time,
  // and the background update deliberately re-renders only when calendar
  // DATA changes - so a page left open past midnight kept yesterday's box
  // until Refresh forced a re-render (the field report). Move it live
  // instead: once a minute, if the date has rolled over, restyle in place -
  // no rebuild, no server call. The overlay's stacking is CSS-controlled
  // (z-index, pointer-events:none), so a re-added overlay renders
  // identically wherever it sits in the DOM. If the new today falls outside
  // the displayed year, the old box is simply removed.
  setInterval(function(){
    const t = new Date();
    const k = [t.getFullYear(), t.getMonth(), t.getDate()].join("-");
    if (k === todayKey) return;
    todayKey = k;
    document.querySelectorAll(".day.today").forEach(el => el.classList.remove("today"));
    document.querySelectorAll(".todayov").forEach(el => el.remove());
    const iso = t.getFullYear() + "-" + String(t.getMonth()+1).padStart(2,"0") + "-" +
                String(t.getDate()).padStart(2,"0");
    const cell = document.querySelector('.day[data-date="' + iso + '"]');
    if (cell){
      cell.classList.add("today");
      const ov = document.createElement("div");
      ov.className = "todayov";
      ov.style.gridColumn = cell.style.gridColumn;
      cell.parentElement.appendChild(ov);
    }
  }, 60000);

  const stampEl = document.getElementById("stamp");
  const perf = data.perf || null;
  const secs = ms => (ms / 1000).toFixed(ms >= 9500 ? 0 : 1) + " s";
  stampEl.textContent =
    "Generated " + new Date(data.generated || Date.now()).toLocaleString() +
    (data.source ? " \u00B7 " + data.source : "") + " \u00B7 " + data.events.length + " items" +
    (perf ? " \u00B7 Outlook read " + secs(perf.catMs + perf.setupMs + perf.enumMs + perf.baselineMs) : "") +
    (data.truncated ? " \u00B7 \u26A0 truncated at 8000 items" : "") +
    " \u00B7 YearView version 3.4.9";
  if (perf) stampEl.title =
    "reader: " + (perf.reader || "classic") +
    " \u00B7 categories " + secs(perf.catMs) + " \u00B7 restrict " + secs(perf.setupMs) +
    " \u00B7 enumerate " + secs(perf.enumMs) + " (" + data.events.length + " items)" +
    (perf.omN || perf.tableN
      ? " \u00B7 of which table " + perf.tableN + " in " + secs(perf.tableMs) +
        " (" + (perf.bulkN
        ? "bulk " + perf.bulkN + (perf.batchN ? " in " + perf.batchN + " call" + (perf.batchN > 1 ? "s" : "") : " row by row")
        : "per-column") + ")" +
        ", object model " + perf.omN + " in " + secs(perf.omMs)
      : "") +
    " \u00B7 baseline sort " + secs(perf.baselineMs) +
    (perf.validateMs ? " \u00B7 first-build validation +" + secs(perf.validateMs) + " (once per session)" : "");

  // weekday header
  const head = document.getElementById("dowhead");
  head.insertAdjacentHTML("beforeend", '<div class="cell corner"></div>');
  for (let c = 0; c < COLS; c++){
    const d = c % 7;
    head.insertAdjacentHTML("beforeend",
      `<div class="cell${(d===0||d===6)?" we":""}">${DOW[d]}</div>`);
  }

  function luma(hex){
    const h = hex.replace("#",""); const r=parseInt(h.substr(0,2),16),
      g=parseInt(h.substr(2,2),16), b=parseInt(h.substr(4,2),16);
    return .299*r + .587*g + .114*b;
  }
  // S2 (v2.1.26): the apostrophe is escaped as well. Every current use is
  // either text content or a DOUBLE-quoted attribute, so this changes nothing
  // today - but it means a future single-quoted attribute cannot become an
  // injection point, which is not a property worth leaving to chance.
  const esc = s => String(s).replace(/[&<>"']/g, m => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
  // S1 (v2.1.26): colours are interpolated into a style attribute, so they are
  // the one value that would be a CSS-injection point if it ever carried user
  // text. Today it cannot: every colour comes from Outlook's fixed
  // OlCategoryColor table or the grey default. That is an INVARIANT nothing
  // enforced, and invariants that live only in a comment do not survive
  // refactoring - so anything that is not a plain six-digit hex colour is
  // replaced with the default here, at the point of use.
  const safeCol = c => (/^#[0-9A-Fa-f]{6}$/.test(String(c || "")) ? String(c) : "#bfbfbf");

  const monthsEl = document.getElementById("months");
  const catKey = ev => (ev.cat || "").split(/[;,]/)[0].trim();
  // Filter persistence (v2.0.8): the hidden-category set is persisted by the
  // BRIDGE in yv-filter.cfg, so it survives sessions, travels with the app's
  // own data, and a genuine uninstall removes it (the old localStorage copy
  // lived in Edge's profile, beyond the uninstaller's reach). Seeding order:
  // a ?hide= URL param (state carried across navigations) wins when present;
  // otherwise the server-injected persisted value; a never-written filter
  // falls back to the first-run default ("(No category)" hidden unless the
  // bridge runs with -IncludeUncategorised). Every change is POSTed to
  // /setfilter at once, so "Show all" persists as an explicit empty filter.
  const csvToSet = src => new Set(String(src).split(",").filter(Boolean)
    .map(s => s === "__none__" ? "" : s));
  function seedHidden(){
    const q = new URLSearchParams(location.search);
    if (q.has("hide")) return csvToSet(q.get("hide") || "");
    if (!HIDE_SET) return new Set(data.uncatDefaultHidden ? [""] : []);
    return csvToSet(HIDE_SEED);
  }
  const hidden = seedHidden();
  // v2.1.31: persist a small setting, WAITING OUT a build.
  //
  // While a calendar read is running, the cooperative in-build server answers
  // anything outside its read-only whitelist with {"busy":true} and expects
  // the caller to retry. Nothing here ever did. saveHidden fired and forgot -
  // it caught network errors, but a busy answer RESOLVES successfully and was
  // discarded in silence. So a filter change made during a build was applied
  // in the page and never written to disk, and the page keeps its own copy, so
  // everything looked right until the next start. Since v2.1.12 a build runs
  // four seconds after every start and every ten minutes, which is exactly how
  // often this had the chance to happen - and why the filter kept coming back.
  //
  // Retries while the answer is busy or absent; stops at once on a real
  // refusal, since retrying that cannot help. A newer value for the same
  // setting supersedes an older in-flight save, so a burst of toggles can
  // never persist out of order.
  const cfgSeq = {};
  // v3.1.8: a config save is a PENDING ACTION. It can spend a minute waiting
  // out a build, and a background update finishing in that minute would
  // navigate the page away - taking the in-flight save with it, and setting
  // `leaving`, which makes the retry loop below stand down deliberately. The
  // save was then simply lost. Counting it as an action parks that navigation
  // until the value has landed, the same gate item creation has used since
  // the equivalent fault there.
  async function persistCfg(path, body, failMsg){
    pendingActions++;
    try {
      const mine = (cfgSeq[path] = (cfgSeq[path] || 0) + 1);
      for (let n = 0; n < 45; n++){
        if (leaving || mine !== cfgSeq[path]) return true;   // superseded: newer value wins
        let j = null;
        try { const r = await api(path, body); j = await r.json(); } catch(_){ j = null; }
        if (j && j.ok) return true;
        if (j && j.err) break;                               // a real refusal
        await new Promise(res => setTimeout(res, 1500));     // busy, or no answer
      }
      if (!leaving && mine === cfgSeq[path] && failMsg) toast(failMsg);
      return false;
    } finally { actionDone(); }
  }
  function saveHidden(){
    const csv = [...hidden].map(k => k === "" ? "__none__" : k).join(",");
    persistCfg("/setfilter", "hide=" + encodeURIComponent(csv),
               "Could not save the category filter");
  }

  // v3.1.2: ONE year's block, built on its own. A neighbour arriving used to
  // rebuild every month of every loaded year - up to three years of grids
  // destroyed and recreated mid-scroll - which is what made the arrival lurch
  // and stutter. Now only the new block is built, and it is spliced in.
  // v3.4.0: ONE month row for a given year and month, lifted out of
  // buildYearBlock unchanged. Twelve rows can now be assembled from ANY
  // sequence of year-and-month pairs, which is what a view starting later
  // than January needs - it necessarily crosses into the following year.
  function buildMonthRow(Y, m, visible){
      const L = YV.monthLayout(Y, m, visible);
    const rows = 1 + Math.max(L.lanes, 1);          // row 1 = day numbers, then lanes
    const mo = document.createElement("div");
    mo.className = "month";
    mo.innerHTML = `<div class="mlabel">${MONTHS[m]}</div>`;
    const g = document.createElement("div");
    g.className = "mgrid";
    g.dataset.month = m;
    g.dataset.year = Y;              // v3.1.0: dates resolve from the BLOCK, not a global
    g.dataset.offset = L.offset;
    g.dataset.ndays = L.nDays;
    g.style.setProperty("--rows", rows + 2);        // +1 pad row, +1 because line ref is exclusive

    // background day cells
    for (let c = 0; c < COLS; c++){
      const dayN = c - L.offset + 1;
      const inMonth = dayN >= 1 && dayN <= L.nDays;
      const dow = c % 7;
      const div = document.createElement("div");
      div.className = "day" + ((dow===0||dow===6) ? " we" : "") + (inMonth ? " in" : " out");
      if (inMonth){
        div.dataset.date = Y + "-" + String(m+1).padStart(2,"0") + "-" + String(dayN).padStart(2,"0");
        if ([Y,m,dayN].join("-") === todayKey){
          div.className += " today";
          const ov = document.createElement("div");
          ov.className = "todayov";
          ov.style.gridColumn = (c+1) + " / " + (c+2);
          g.appendChild(ov);
        }
      }
      div.style.gridColumn = (c+1) + " / " + (c+2);
      g.appendChild(div);
      if (inMonth){
        const dn = document.createElement("div");
        dn.className = "dn";
        dn.style.gridColumn = (c+1) + " / " + (c+2);
        dn.textContent = dayN;
        g.appendChild(dn);
      }
    }
    // event bars
    for (const ev of L.events){
      const bar = document.createElement("div");
      if (!catKey(ev)){
        // "(No category)": hollow bar, no colour fill (keyed on the category
        // string, not the colour - grey #bfbfbf is also a real palette colour)
        bar.className = "bar nocat";
      } else {
        const col = safeCol(ev.col);
        bar.className = "bar " + (luma(col) < 140 ? "dark" : "light");
        bar.style.background = col;
      }
      bar.style.gridColumn = (L.offset + ev.d0 + 1) + " / " + (L.offset + ev.d1 + 2);
      bar.style.gridRow = (ev.lane + 2);
      // v3.4.7: titles per category. The text stays IN the bar - opening and
      // deleting an item both read it - and only its ink is removed.
      if (noTitles.has(catKey(ev))) bar.classList.add("notitle");
      const when = ev.ad ? "All day" :
        new Date(ev.s).toLocaleTimeString([], {hour:"numeric", minute:"2-digit"});
      bar.title = `${ev.t}\n${when}${ev.cat ? "\n" + ev.cat : ""}\n(click to open in Outlook)`;
      if (ev.id) bar.dataset.id = ev.id;
      if (ev.r) bar.dataset.rec = "1";
      bar.dataset.start = ev.s;
      bar.innerHTML = esc(ev.t);
      g.appendChild(bar);
    }
    // ensure at least one empty lane row + padding for visual air
    const pad = document.createElement("div");
    pad.className = "padrow";
    pad.style.gridRow = rows + 1;
    g.appendChild(pad);
    mo.appendChild(g);
    return mo;
  }
  function buildYearBlock(Y, visible){
    const yb = document.createElement("div");
    yb.className = "yearblock";
    yb.dataset.year = Y;
    // v3.1.5: the band exists to say which year a run of twelve month rows
    // belongs to - a question that cannot arise when the page shows one year
    // and the heading already answers it. Continuous scrolling only.
    if (continuous){
      const ysep = document.createElement("div");
      ysep.className = "ysep";
      ysep.textContent = Y;
      yb.appendChild(ysep);
    }
    for (let m = 0; m < 12; m++) yb.appendChild(buildMonthRow(Y, m, visible));
    return yb;
  }
  // v3.4.0: twelve rows beginning at `start` in year Y and running on into
  // Y+1. Each row carries its OWN year, so every date, drag-create and
  // today-box check resolves from the row rather than from the page.
  function buildSpanBlock(Y, start, visible){
    const yb = document.createElement("div");
    yb.className = "yearblock";
    yb.dataset.year = Y;
    for (let i = 0; i < 12; i++){
      const n = start + i;
      yb.appendChild(buildMonthRow(Y + (n > 11 ? 1 : 0), n % 12, visible));
    }
    return yb;
  }
  function visibleEvents(){
    return allEvents.filter(ev => !hidden.has(catKey(ev)));
  }
  // Full rebuild - the filter changing, the setting being turned off, the
  // first paint. Loading a year does NOT come through here any more.
  function renderMonths(){
    monthsEl.innerHTML = "";
    const visible = visibleEvents();
    if (startMonth > 0){
      monthsEl.appendChild(buildSpanBlock(YEAR, startMonth, visible));
      return;
    }
    for (const Y of yearsAsc()) monthsEl.appendChild(buildYearBlock(Y, visible));
  }
  {
    // time the initial render and append it to the perf tooltip, so client-side
    // cost is visible alongside the server-side breakdown
    const rt0 = performance.now();
    renderMonths();
    const stampT = document.getElementById("stamp");
    if (stampT && stampT.title) stampT.title += " \u00B7 render " +
      ((performance.now() - rt0) / 1000).toFixed(2) + " s";
  }

  // legend (clickable: arm a category, then click a day to create an item)
  const legend = document.getElementById("legend");
  const used = new Map();
  for (const ev of data.events) if (ev.cat) used.set(ev.cat.split(/[;,]/)[0].trim(), ev.col);
  // v2.1.18: the legend row is alphabetical too, matching the picker and the
  // filter panel (v2.1.17). Sorting a COPY here also feeds the filter panel,
  // which derives its rows from this list - the payload order is untouched.
  const catsCmp = (a, b) => String(a.name).localeCompare(String(b.name), undefined,
                                                         { numeric: true, sensitivity: "base" });
  const cats = data.categories && data.categories.length ? data.categories.slice().sort(catsCmp)
             : [...used].map(([name, col]) => ({name, col}));
  for (const c of cats){
    if (!c.name) continue;
    const chip = document.createElement("span");
    chip.className = "item";
    chip.dataset.cat = c.name;
    chip.innerHTML = `<span class="sw" style="background:${safeCol(c.col)}"></span>${esc(c.name)}`;
    legend.appendChild(chip);
  }
  // v2.0.20: "(No category)" arms creation of items with no category at all.
  // Its dataset.cat is the EMPTY STRING - the same key the filter already uses
  // for uncategorised - so filter dimming, unhide-on-arm and the hidden set all
  // work through the existing paths without a parallel notion of "no category".
  const noneChip = document.createElement("span");
  noneChip.className = "item";
  noneChip.dataset.cat = "";
  noneChip.title = "Create items with no category";
  noneChip.innerHTML = '<span class="sw none"></span>(No category)';
  legend.appendChild(noneChip);
  const clearChip = document.createElement("span");
  clearChip.className = "item clearcat";
  clearChip.textContent = "\u2715 Deselect";
  clearChip.title = "Deselect the armed category (or press Esc)";
  legend.appendChild(clearChip);
  const filterChip = document.createElement("span");
  filterChip.className = "item filtercat";
  filterChip.title = "Choose which categories are shown";
  legend.appendChild(filterChip);
  const newCatChip = document.createElement("span");
  newCatChip.className = "item newcat";
  newCatChip.textContent = "+ New category\u2026";
  newCatChip.title = "Select an Outlook category, then click or drag days to use it";
  legend.appendChild(newCatChip);
  legend.insertAdjacentHTML("beforeend",
    '<span class="hint">Click a category, then click or drag across days to create an item \u00B7 click a bar to open it \u00B7 right-click a bar to delete</span>');

  /* ---------- new-item creation ---------- */
  const SERVED = location.protocol === "http:" || location.protocol === "https:";
  // S-2: every bridge call is a POST carrying the token in a header - the token
  // never appears in a URL, so never in history or the address bar.
  // One silent retry: after sleep/wake or a long idle the first request can fail
  // transiently (socket teardown, the browser waking the page) even though the
  // bridge is fine. A genuine bridge outage still fails the retry and surfaces.
  let reloaded403 = false;
  async function api(path, body, signal){
    const opts = {method:"POST", headers:{"X-YV-Token": TOK}};
    if (signal) opts.signal = signal;
    if (body !== undefined){
      opts.headers["Content-Type"] = "application/x-www-form-urlencoded";
      opts.body = body;
    }
    try {
      const r = await fetch(path, opts);
      // v2.0.29: a 403 here almost always means THIS PAGE outlived the bridge
      // it was served by - the bridge restarted (sleep, crash, relaunch) and
      // this page's token is dead, so every action fails and its text/plain
      // body used to surface as "Bridge not reachable". The page routes are
      // tokenless: one reload fetches a fresh page and token from the running
      // bridge and the session continues. Once only, so a genuine 403 storm
      // cannot loop the page.
      if (r.status === 403 && !leaving && !reloaded403){
        reloaded403 = true;
        toast("Reconnecting \u2014 reloading\u2026", 0);
        setTimeout(() => location.reload(), 300);
      }
      return r;
    }
    catch(err){
      // F-4: retry only routes that are safe to repeat. Retrying /new or /delete
      // could double-create or double-delete if the first request landed but its
      // response was lost mid-flight.
      if (leaving || !/^\/(dirty|build|quit)\b/.test(path)) throw err;
      await new Promise(res => setTimeout(res, 1500));
      return fetch(path, opts);
    }
  }
  // v2.0.29: one silent retry for paths that are safe to reissue (the build
  // is deduplicated server-side; categories and calendars are reads). Heals a
  // momentary connection-level transient instead of toasting for it.
  function apiR(path, signal){
    return api(path, undefined, signal).catch(() =>
      new Promise(res => setTimeout(res, 800)).then(() => api(path, undefined, signal)));
  }
  // When this page navigates away, the browser aborts every in-flight fetch; those
  // rejections are not bridge failures and must not toast an error.
  let leaving = false;
  window.addEventListener("beforeunload", () => { leaving = true; });
  window.addEventListener("pagehide",     () => { leaving = true; });
  // Rebuild the current year on the bridge, then navigate to the cached result.
  // Single-flight: the delete flow and the dirty poll can both request a rebuild in
  // the same second, and the first navigation would abort the second's fetch.
  let rebuilding = false;
  let bgUpdating = false;   // background (quiet-path) rebuild in flight - does
                            // NOT block year navigation, unlike `rebuilding`
  // v2.0.12: navigation destroys all running JavaScript - including a create/
  // open/delete waiting out a build (apiWaitBusy). "Waiting ... will create
  // the item" then silently never created it whenever the finished build
  // navigated to a changed page. ALL programmatic navigations now pass this
  // gate: while an item action pends, the navigation is parked; when the
  // last action resolves, the parked navigation runs (a short beat later, so
  // the action's toast/inspector lands first).
  let pendingActions = 0, deferredNav = null;
  function navTo(url){
    if (pendingActions > 0){ deferredNav = url; return; }
    leaving = true; location.href = url;
  }
  function actionDone(){
    if (--pendingActions <= 0){
      pendingActions = 0;
      if (deferredNav){
        const u = deferredNav; deferredNav = null;
        setTimeout(() => { leaving = true; location.href = u; }, 900);
      }
    }
  }
  // Build-then-navigate for ANY year: a plain location.href into a dead bridge
  // renders Edge's raw ERR_CONNECTION_REFUSED page; going through /build first
  // keeps failures on this page as a toast, and pre-warms the cache so year
  // changes are as seamless as refreshes (no splash flash).
  function goYear(y, msg, force){
    if (rebuilding || leaving) return;
    if (y < 1990 || y > 2100){                       // F-2: match the server clamp
      toast("Year " + y + " is outside the supported range (1990\u20132100)");
      return;
    }
    rebuilding = true;
    let buildWaits = 0;                        // v2.0.30
    const doBuild = () => {
      // v2.1.3: FIRST attempt only. This ran on every retry tick, so a server
      // message such as "Outlook is not running - please start Outlook" was
      // shown, then overwritten two seconds later by the progress text - which
      // is why the field report saw a persistent "Refreshing - Reading
      // Outlook" and never the instruction it was replacing.
      if (msg && buildWaits === 0) toast(msg, 0);
      // v2.1.1: bound the request. If the bridge is ever wedged (it should no
      // longer be able to wedge on Outlook, but this is the last line of
      // defence) the fetch would hang and the progress toast would sit there
      // forever - exactly the reported symptom. Abort and say something true.
      const ac = (typeof AbortController !== "undefined") ? new AbortController() : null;
      const killer = ac ? setTimeout(() => ac.abort(), 150000) : null;
      apiR("/build?year=" + y, ac ? ac.signal : undefined)
        .then(r => {
          if (killer) clearTimeout(killer);
          if (r.status === 503){   // bridge busy building: navigate; splash retries
            navTo("/?year=" + y + navSuffix()); return;
          }
          if (!r.ok) throw 0;
          return r.json().then(b => {
            if (b && b.closing){
              // v2.1.9: the session is ending. Navigating here rendered the
              // JSON notice as the page; stand down instead - the dialog
              // explains everything.
              leaving = true;
              hideToast();
              return;
            }
            // v2.0.30: a 200 can be the retryable "please start Outlook"
            // state. The Refresh path used to leave its progress toast up
            // forever here; now it shows the request, re-tries every 2s,
            // and gives up honestly. Outlook is never started by YearView.
            if (b && b.retry){
              if (++buildWaits > 60){
                rebuilding = false;
                toast("Outlook still isn\u2019t running \u2014 press Refresh again once it is");
                return;
              }
              toast((b.err || "Waiting for Outlook") + "\u2026", 0);
              setTimeout(doBuild, 2000);
              return;
            }
            navTo("/?year=" + y + navSuffix());
          });
        })
        .catch(() => {
          if (!leaving && bgUpdating){
            // A fetch REJECTION (not a 503) while a background update runs is a
            // transient of the busy bridge, not an outage - navigating is the
            // most resilient move: GET navigations are retried by the browser
            // and answered by the in-build quick-serve or, failing that, the
            // main loop once it frees. The splash covers uncached years.
            navTo("/?year=" + y + navSuffix()); return;
          }
          rebuilding = false;
          if (!leaving) toast("YearView is not responding \u2014 restart YearView from the shortcut");
        });
    };
    if (force){ doBuild(); return; }               // Refresh/delete/dirty: always rebuild
    // Year navigation: switch to a cached year INSTANTLY. Since v1.2.25 a
    // cached year may be STALE after a calendar change - it serves at once and
    // then self-revalidates (quiet rebuild + hash compare), refreshing with
    // "YearView updated" only if that year was actually affected. Uncached
    // years pay the build; a cache vanishing between check and load falls to
    // the splash flow gracefully.
    api("/cached?year=" + y)
      .then(r => r.json())
      .then(c => {
        if (c.cached || c.busy){
          // busy: the bridge is mid-build - navigate directly; the in-build
          // quick-serve answers page GETs from cache, or with a splash that
          // retries the build once the bridge frees (v1.2.23)
          navTo("/?year=" + y + navSuffix());
        }
        else doBuild();
      })
      .catch(() => {
        // Probe failed at the network level (transient of a busy bridge, a
        // speculative connection being closed, etc). Do NOT escalate to
        // another fragile POST - navigate directly; GET navigation is the
        // most resilient operation and every outcome of it is handled.
        navTo("/?year=" + y + navSuffix());
      });
  }
  function rebuild(){ goYear(YEAR, null, true); }
  const prevYr = document.getElementById("prevYr"), nextYr = document.getElementById("nextYr");
  if (SERVED){
    // v2.0.0: calendar picker - lists the default store's calendars and
    // switches the whole view (server re-points change detection, clears the
    // page cache, and the forced rebuild lands on the new calendar).
    const calB = document.getElementById("calBtn");
    calB.style.display = "";
    const calPanel = document.createElement("div");
    calPanel.className = "ctx";
    calPanel.style.display = "none";
    document.body.appendChild(calPanel);
    const hideCalPanel = () => { calPanel.style.display = "none"; };
    document.addEventListener("click", e => {
      if (calPanel.style.display !== "none" && !calPanel.contains(e.target) && e.target !== calB) hideCalPanel();
    });
    calB.onclick = () => {
      if (calPanel.style.display !== "none"){ hideCalPanel(); return; }
      apiR("/calendars")
        .then(r => r.json())
        .then(j => {
          if (j.busy){ toast("YearView is busy \u2014 try again in a moment"); return; }
          if (!j.ok && j.err){ toast(j.err); return; }   // v2.0.24
          if (!j.ok || !j.calendars || !j.calendars.length) throw 0;
          calPanel.textContent = "";
          j.calendars.forEach(c => {
            const row = document.createElement("div");
            row.textContent = (c.sel ? "\u25CF " : "\u25CB ") + c.name + (c.def ? " (default)" : "");
            row.style.padding = "5px 12px";
            row.style.cursor = "pointer";
            row.style.whiteSpace = "nowrap";
            row.onmouseenter = () => { row.style.background = "rgba(0,0,0,0.06)"; };
            row.onmouseleave = () => { row.style.background = ""; };
            row.onclick = () => {
              hideCalPanel();
              if (c.sel) return;
              toast("Switching to \u201C" + c.name + "\u201D\u2026", 0);
              // v2.0.3: EVERY failure class must name itself - the running
              // page version (settles which build is answering), then: busy,
              // the server's staged reason, the HTTP status for a non-JSON
              // reply (403s send plain text), or the network error message.
              const ver = ((document.getElementById("stamp") || {textContent:""}).textContent
                            .match(/version ([0-9.]+)/) || [,"?"])[1];
              // v2.0.5: switch by tiny INDEX into the server-cached list -
              // no folder identifiers on the wire, no request body, and the
              // exact route shape that provably works for listing.
              api("/calendars?set=" + c.i)
                .then(async r => {
                  let body = null;
                  try { body = await r.json(); } catch (_) {}
                  if (body && body.busy){
                    toast("YearView is busy updating \u2014 try again in a moment"); return;
                  }
                  if (body && body.ok){ goYear(YEAR, null, true); return; }
                  const detail = body ? (body.err || "no error detail returned")
                                      : ("HTTP " + r.status + " with non-JSON reply");
                  toast("Could not switch calendar [page v" + ver + "] \u2014 " + detail);
                })
                .catch(err => toast("Could not switch calendar [page v" + ver + "] \u2014 request failed: " +
                  (err && err.message ? err.message : String(err))));
            };
            calPanel.appendChild(row);
          });
          const br = calB.getBoundingClientRect();
          calPanel.style.left = Math.max(8, br.left - 30) + "px";
          calPanel.style.top = (br.bottom + 6) + "px";
          calPanel.style.display = "block";
        })
        .catch(() => toast("Could not list calendars"));
    };
    const rb = document.getElementById("refreshBtn");
    rb.style.display = "";
    rb.onclick = () => { rb.textContent = "Refreshing\u2026"; toast("Refreshing \u2014 reading Outlook\u2026", 0); rebuild(); };
    const qb = document.getElementById("quitBtn");
    qb.style.display = "";
    qb.onclick = async () => {
      try { await api("/quit"); } catch(e){}
      try { window.close(); } catch(e){}
      document.title = "YearView (stopped)";
      toast("YearView stopped \u2014 this window will close. Use the YearView shortcut to start again.");
    };
    // v3.1.0: with the timeline on, a loaded neighbour is already on the page -
    // scroll to it rather than rebuilding and navigating. An unloaded year
    // still goes the long way round.
    const stepTo = d => {
      const y = focusYear + d;
      if (continuous && years.has(y)){
        const yb = [...monthsEl.querySelectorAll(".yearblock")].find(x => +x.dataset.year === y);
        if (yb){ yb.scrollIntoView({ behavior: "smooth", block: "start" }); return; }
      }
      goYear(y, "Loading " + y + "\u2026");
    };
    prevYr.onclick = () => stepTo(-1);
    nextYr.onclick = () => stepTo(1);
  } else {
    prevYr.style.display = "none";
    nextYr.style.display = "none";
  }
  let armed = null;
  const toastEl = document.getElementById("toast");
  let toastTimer = null;
  // v2.1.7: once the session is ending (a navigation, or Outlook gone) nothing
  // may raise a toast: the page is about to disappear behind a dialog that
  // explains itself, and a late arrival from an in-flight request would sit on
  // top of it saying something else.
  function hideToast(){
    toastEl.classList.remove("show");
    toastEl.textContent = "";
    clearTimeout(toastTimer);
  }
  function toast(msg, ms){
    if (leaving) { hideToast(); return; }
    // ms === 0 -> sticky: stays until replaced or the page navigates away.
    // Used for progress messages, which previously died at the fixed 3.5 s
    // while a long Outlook read was still only half done.
    toastEl.textContent = msg;
    toastEl.classList.add("show");
    clearTimeout(toastTimer);
    if (ms !== 0) toastTimer = setTimeout(() => toastEl.classList.remove("show"), ms || 3500);
  }
  function armChip(chip){
    legend.querySelectorAll(".item.sel").forEach(x => x.classList.remove("sel"));
    document.body.classList.remove("armed");
    armed = null;
    if (chip){
      chip.classList.add("sel");
      document.body.classList.add("armed");
      armed = chip.dataset.cat;
      if (hidden.has(armed)){        // arming a hidden category unhides it
        hidden.delete(armed);
        saveHidden();
        syncFilterUI();
        renderMonths();
      }
    }
  }

  /* ---- settings ---- */
  let confirmDelete = CONFIRM_DELETE;
  let scrollYears   = SCROLL_YEARS;
  let printPages    = PRINT_PAGES;
  // v3.0.0: one table, so adding a setting is an entry rather than another
  // copy of the open/change/revert logic.
  // v3.4.4: the ORDER here is the order the panel shows. "Year starts at"
  // leads because it is the one that changes what you are looking at; the
  // other two change what happens later.
  const SETTINGS = [
    // v3.4.0: "Current month" follows the calendar and so moves on its own;
    // the named months are fixed. Changing it takes effect at once - no
    // reload - which is why the value is resolved through a function rather
    // than captured once.
    { id: "setStartMonth",  key: "startmonth",    label: "Year starts at", type: "select",
      options: [[0, "January (whole year)"], [-1, "Current month"],
                [1, "February"], [2, "March"],     [3, "April"],    [4, "May"],
                [5, "June"],     [6, "July"],      [7, "August"],   [8, "September"],
                [9, "October"],  [10, "November"], [11, "December"]],
      get: () => startMonthPref, set: v => { startMonthPref = v; applyStartMonth(); } },
    { id: "setConfirmDel",  key: "confirmdelete", label: "Confirm before deleting",
      get: () => confirmDelete, set: v => { confirmDelete = v; } },
    // v3.3.6: the continuous-scrolling entry is REMOVED FROM THE PANEL, not
    // from the application. Every mechanism it drove is still here and still
    // works - the year blocks, the neighbour loading, the scroll anchoring,
    // setContinuous() - it simply has no control any more, and `continuous`
    // is pinned off below. Restoring it means putting this entry back and
    // nothing else.
    // v3.2.0: the first setting that is not a yes/no, so the table gained a
    // `type`. Everything else about it - open, change, persist, revert - goes
    // through the same one path the checkboxes use.
    { id: "setPrintPages",  key: "printpages",    label: "Print \u2014 fit to", type: "select",
      options: [[1, "1 page"], [2, "2 pages"], [3, "3 pages"], [4, "4 pages"], [0, "As many as needed"]],
      get: () => printPages,    set: v => { printPages = v; } },
    // v3.4.8: printing only, so it sits with the other print setting.
    { id: "setPrintLegend", key: "printlegend",   label: "Print category legend",
      get: () => printLegend,   set: v => { printLegend = v; } }
  ];
  const spanel = document.createElement("div");
  spanel.className = "ctx fpanel";
  document.body.appendChild(spanel);
  spanel.innerHTML = SETTINGS.map(o =>
    o.type === "select"
      ? '<label class="frow"><span>' + esc(o.label) + '</span>' +
        '<select id="' + o.id + '">' +
        o.options.map(p => '<option value="' + p[0] + '">' + esc(p[1]) + '</option>').join("") +
        '</select></label>'
      : '<label class="frow"><input type="checkbox" id="' + o.id + '">' +
        '<span>' + esc(o.label) + '</span></label>').join("");
  const setBtn = document.getElementById("setBtn");
  spanel.addEventListener("click", e => e.stopPropagation());
  setBtn.addEventListener("click", e => {
    e.stopPropagation();
    hideCtx();
    if (spanel.style.display === "block"){ spanel.style.display = "none"; return; }
    SETTINGS.forEach(o => {
      const el = spanel.querySelector("#" + o.id);
      if (!el) return;
      if (o.type === "select") el.value = String(o.get()); else el.checked = o.get();
    });
    spanel.style.display = "block";
    const r = setBtn.getBoundingClientRect();
    spanel.style.left = Math.max(8, Math.min(r.left, innerWidth - spanel.offsetWidth - 8)) + "px";
    spanel.style.top  = (r.bottom + 4) + "px";
  });
  spanel.addEventListener("change", e => {
    const opt = SETTINGS.find(o => o.id === e.target.id);
    if (!opt) return;
    // v2.1.30: if the save fails, put the tick BACK. A control that shows a
    // state the server did not accept is worse than the failure itself.
    const sel  = opt.type === "select";
    const prev = opt.get();
    // v3.3.0: a select's value is a string; only coerce to a number when the
    // option list is numeric, or "a3l" would arrive as NaN.
    const numeric = sel && opt.options.every(p => typeof p[0] === "number");
    const want = sel ? (numeric ? +e.target.value : e.target.value) : e.target.checked;
    opt.set(want);
    const revert = () => {
      opt.set(prev);
      const el = spanel.querySelector("#" + opt.id);
      if (el){ if (sel) el.value = String(prev); else el.checked = prev; }
      if (!leaving) toast("Could not save that setting");
    };
    // Server is the authority and reseeds every page from it on restart.
    // v2.1.31: waits out a build rather than failing against it.
    persistCfg("/setsetting", opt.key + "=" + (sel ? String(want) : (want ? "1" : "0")), null)
      .then(ok => { if (!ok) revert(); });
  });
  document.addEventListener("click", () => { spanel.style.display = "none"; });

  /* ---- v3.1.7: the wheel and the swipe no longer change year ----
     v3.0.0 made a wheel past the edge of the page, or a vertical swipe, step
     to the next year. Continuous scrolling replaced what that was reaching
     for, and leaving it as the fallback meant a single-year page could still
     jump a year under a stray scroll - and a touch swipe had to suppress the
     synthetic mouse events it produced, which is machinery worth being rid
     of. Off means off: the wheel and touch now only scroll the page, and the
     year changes by the < > buttons alone. */

  /* ---- category filter ---- */
  const fpanel = document.createElement("div");
  fpanel.className = "ctx fpanel";
  document.body.appendChild(fpanel);
  function navSuffix(){
    let s = "";
    if (armed !== null) s += "&arm=" + encodeURIComponent(armed === "" ? "__none__" : armed);
    if (hidden.size) s += "&hide=" + encodeURIComponent(
      [...hidden].map(k => k === "" ? "__none__" : k).join(","));
    return s;
  }
  function syncFilterUI(){
    filterChip.textContent = "\u25E7 Filter" + (hidden.size ? " (" + hidden.size + " hidden)" : "");
    filterChip.classList.toggle("active", hidden.size > 0);
    legend.querySelectorAll(".item").forEach(c => {
      if (c.dataset.cat !== undefined) c.classList.toggle("off", hidden.has(c.dataset.cat));
    });
    fpanel.querySelectorAll("input[type=checkbox]").forEach(cb => {
      const k = cb.dataset.key === "__none__" ? "" : cb.dataset.key;
      cb.checked = cb.dataset.col === "title" ? !noTitles.has(k) : !hidden.has(k);
    });
  }
  function buildFilterPanel(){
    // v2.1.17: same alphabetical ordering as the category picker, so the two
    // lists read consistently.
    const rows = cats.filter(c => c.name)
      .slice()
      .sort((a, b) => String(a.name).localeCompare(String(b.name), undefined,
                                                   { numeric: true, sensitivity: "base" }))
      .map(c => filterRow(esc(c.name), safeCol(c.col), esc(c.name)));
    // "(No category)" is always offered (v1.20), whether or not any such
    // events currently exist - the filter state should be settable either way.
    rows.push(filterRow("__none__", "#fff", "(No category)"));
    fpanel.innerHTML =
      '<div class="fbtns"><span data-f="all">Show all</span><span data-f="none">Hide all</span>' +
      '<span class="tfirst" data-f="tall">All titles</span><span data-f="tnone">No titles</span></div>' +
      '<div class="fhead"><span>Show</span><span class="fth">Titles</span></div>' +
      rows.join("");
  }
  // v3.4.7: TWO checkboxes per category - shown at all, and shown with titles -
  // so each sits in its OWN label. One label around both would hand every
  // click on the name to whichever input came first.
  function filterRow(key, col, name){
    return '<div class="frow fcat">' +
      '<label class="fshow"><input type="checkbox" data-col="show" data-key="' + key + '">' +
      '<span class="sw" style="background:' + col + '"></span>' + name + '</label>' +
      '<label class="ftitle" title="Show item titles for this category">' +
      '<input type="checkbox" data-col="title" data-key="' + key + '"></label></div>';
  }
  buildFilterPanel();
  filterChip.addEventListener("click", e => {
    e.stopPropagation();
    hideCtx();
    const r = filterChip.getBoundingClientRect();
    fpanel.style.display = "block";
    fpanel.style.left = Math.min(r.left, innerWidth - fpanel.offsetWidth - 8) + "px";
    fpanel.style.top = Math.max(8, r.top - fpanel.offsetHeight - 6) + "px";
    syncFilterUI();
  });
  // v3.4.7: the checkboxes act on CHANGE, which fires once per real toggle
  // whether the box or its label was clicked; clicks are left to the buttons
  // (and to keeping the panel open).
  fpanel.addEventListener("click", e => {
    e.stopPropagation();
    const f = e.target.dataset ? e.target.dataset.f : null;
    if (!f) return;
    if (f === "all"){ hidden.clear(); saveHidden(); }
    else if (f === "none"){
      cats.forEach(c => { if (c.name) hidden.add(c.name); });
      hidden.add("");
      saveHidden();
    }
    else if (f === "tall" || f === "tnone"){
      const prev = new Set(noTitles);
      noTitles.clear();
      if (f === "tnone"){
        cats.forEach(c => { if (c.name) noTitles.add(c.name); });
        noTitles.add("");
      }
      saveNoTitles(prev);
    }
    else return;
    syncFilterUI();
    renderMonths();
  });
  fpanel.addEventListener("change", e => {
    const input = e.target;
    if (!input || !input.dataset || input.dataset.key === undefined) return;
    const key = input.dataset.key === "__none__" ? "" : input.dataset.key;
    if (input.dataset.col === "title"){
      const prev = new Set(noTitles);
      if (input.checked) noTitles.delete(key); else noTitles.add(key);
      saveNoTitles(prev);
    } else {
      if (input.checked) hidden.delete(key); else hidden.add(key);
      saveHidden();
    }
    syncFilterUI();
    renderMonths();
  });
  // v3.4.7: saved through /setsetting like any other setting, so the server
  // re-stamps built pages and a background update cannot bring titles back.
  // A refused or failed save puts the previous choice back - a panel must not
  // show a state the server did not accept.
  function saveNoTitles(prev){
    const csv = [...noTitles].map(k => k === "" ? "__none__" : k).join(",");
    persistCfg("/setsetting", "notitles=" + encodeURIComponent(csv), null).then(ok => {
      if (ok || leaving) return;
      noTitles.clear();
      prev.forEach(k => noTitles.add(k));
      syncFilterUI();
      renderMonths();
      toast("Could not save the title choices");
    });
  }
  syncFilterUI();
  legend.addEventListener("click", e => {
    const chip = e.target.closest(".item");
    // v2.0.20: an EMPTY dataset.cat is the "(No category)" chip and must arm;
    // only chips without the attribute at all (Deselect/Filter/+New) are skipped.
    if (!chip || chip.dataset.cat === undefined) return;
    const was = chip.classList.contains("sel");
    armChip(was ? null : chip);
    if (!was) toast((armed === "" ? "No category armed" : "Category \u201C" + armed + "\u201D armed") +
      " \u2014 click a day, or drag across days (Esc to cancel)");
  });
  clearChip.addEventListener("click", e => {
    e.stopPropagation();
    armChip(null);
    endDrag(false);
  });
  // v2.0.14: "+ New category" opens a picker of the FULL Outlook master
  // category list (the legend shows only categories in use this year), so an
  // existing category can be selected instead of retyped from memory. A text
  // field at the top still creates a genuinely new one via /newcat. Selecting
  // an existing entry needs no server round-trip - the list already carries
  // its colour - and arms it exactly as creation does.
  function armCategory(name, col){
    let chip = [...legend.querySelectorAll(".item")].find(c => c.dataset.cat === name);
    if (!chip){
      chip = document.createElement("span");
      chip.className = "item";
      chip.dataset.cat = name;
      chip.innerHTML = `<span class="sw" style="background:${safeCol(col)}"></span>` + esc(name);
      legend.insertBefore(chip, newCatChip);
    }
    armChip(chip);
    toast("Category \u201C" + name + "\u201D ready \u2014 click or drag days to create an item");
  }
  const catPanel = document.createElement("div");
  catPanel.className = "ctx";
  catPanel.style.display = "none";
  document.body.appendChild(catPanel);
  const hideCatPanel = () => { catPanel.style.display = "none"; };
  document.addEventListener("click", e => {
    if (catPanel.style.display !== "none" && !catPanel.contains(e.target) && e.target !== newCatChip) hideCatPanel();
  });
  // v2.0.21: category-name rules mirrored from the server so the picker can
  // reject a bad name immediately instead of round-tripping. The SERVER check
  // is the authoritative one - this is a courtesy, not the guard.
  function catNameError(name){
    if (name.length > 255) return "Category names cannot be longer than 255 characters.";
    if (/[,;]/.test(name)) return "Category names cannot contain a comma or semicolon.";
    if (/[\u0000-\u001F\u007F]/.test(name)) return "Category names cannot contain control characters.";
    return null;
  }
  async function createCategory(name){
    try{
      const r = await api("/newcat?name=" + encodeURIComponent(name));
      const j = await r.json();
      if (j && j.retry){ toast((j.err || "Waiting for Outlook") + " \u2014 try again in a moment"); return; }
      if (!j.ok){ toast("Outlook error: " + (j.err || "unknown")); return; }
      armCategory(j.name, j.col);
    } catch(err){
      if (!leaving) toast("YearView is not responding \u2014 restart YearView");
    }
  }
  newCatChip.addEventListener("click", async e => {
    e.stopPropagation();
    if (catPanel.style.display !== "none"){ hideCatPanel(); return; }
    if (!SERVED){ toast("Creating categories needs the YearView app \u2014 run the script with -Serve"); return; }
    try{
      const r = await apiR("/categories");
      const j = await r.json();
      if (j.busy){ toast("YearView is busy \u2014 try again in a moment"); return; }
      // v2.0.24: a JSON error means the bridge answered - say what it said
      // rather than falling through to "Bridge not reachable".
      if (!j.ok){ toast(j.err || "Could not read categories from Outlook"); return; }
      catPanel.textContent = "";
      const inRow = document.createElement("div");
      inRow.style.cssText = "padding:6px 10px;display:flex;gap:6px;border-top:1px solid rgba(0,0,0,0.12)";
      const inp = document.createElement("input");
      inp.type = "text"; inp.placeholder = "New category name\u2026"; inp.maxLength = 255;
      inp.style.cssText = "flex:1;min-width:150px;font:inherit;padding:3px 6px";
      const btn = document.createElement("button");
      btn.textContent = "Create";
      btn.style.cssText = "font:inherit;padding:3px 10px;cursor:pointer";
      const doCreate = () => {
        const name = (inp.value || "").trim();
        if (!name) return;
        const nErr = catNameError(name);          // v2.0.21
        if (nErr){ toast(nErr); return; }
        hideCatPanel();
        createCategory(name);
      };
      btn.onclick = ev => { ev.stopPropagation(); doCreate(); };
      inp.onkeydown = ev => { if (ev.key === "Enter") doCreate(); };
      inp.onclick = ev => ev.stopPropagation();
      inRow.appendChild(inp); inRow.appendChild(btn);
      // v2.1.17: Outlook returns its master list in ITS OWN order (roughly
      // creation order), which reads as no order at all once there are more
      // than a handful. Sort a COPY alphabetically for display - the server
      // still sends master-list order, and nothing else depends on the
      // sequence. localeCompare with sensitivity "base" ignores case and
      // accents; numeric:true keeps "Project 2" before "Project 10".
      const catsSorted = (j.cats || []).slice().sort((a, b) =>
        String(a.name).localeCompare(String(b.name), undefined,
                                     { numeric: true, sensitivity: "base" }));
      catsSorted.forEach(c => {
        const row = document.createElement("div");
        row.style.cssText = "padding:5px 12px;cursor:pointer;white-space:nowrap;display:flex;align-items:center;gap:7px";
        const sw = document.createElement("span");
        sw.style.cssText = "width:11px;height:11px;border-radius:3px;display:inline-block;background:" + safeCol(c.col);
        row.appendChild(sw);
        row.appendChild(document.createTextNode(c.name));
        row.onmouseenter = () => { row.style.background = "rgba(0,0,0,0.06)"; };
        row.onmouseleave = () => { row.style.background = ""; };
        row.onclick = () => { hideCatPanel(); armCategory(c.name, c.col); };
        catPanel.appendChild(row);
      });
      catPanel.appendChild(inRow);   // v2.0.15: create field sits at the BOTTOM of the list
      const br = newCatChip.getBoundingClientRect();
      catPanel.style.left = Math.max(8, br.left - 30) + "px";
      // v2.0.15: the legend lives near the window's bottom edge, so the panel
      // opens UPWARD - its bottom edge anchored just above the chip (position
      // is fixed, so this is pure viewport arithmetic); maxHeight keeps the
      // top on-screen and the list scrolls.
      catPanel.style.top = "auto";
      catPanel.style.bottom = (window.innerHeight - br.top + 6) + "px";
      catPanel.style.maxHeight = "60vh";
      catPanel.style.overflowY = "auto";
      catPanel.style.display = "block";
      setTimeout(() => inp.focus({ preventScroll: true }), 0);
    } catch(err){
      if (!leaving) toast("YearView is not responding \u2014 restart YearView");
    }
  });
  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && armed !== null){
      armChip(null);
      endDrag(false);
    }
  });
  // restore armed category after an auto-refresh
  const armParam = new URLSearchParams(location.search).get("arm");
  if (armParam){
    const armKey = (armParam === "__none__") ? "" : armParam;   // v2.0.20
    const rechip = [...legend.querySelectorAll(".item")].find(c => c.dataset.cat === armKey);
    if (rechip) armChip(rechip);
  }

  /* ---- pointer geometry: which date is under x within a month grid ---- */
  const pad2 = n => String(n).padStart(2,"0");
  const isoOf = (y, m, d) => y + "-" + pad2(m+1) + "-" + pad2(d);
  function dateAt(el, clientX){
    const g = el && el.closest ? el.closest(".mgrid") : null;
    if (!g) return null;
    const r = g.getBoundingClientRect();
    const col = Math.floor((clientX - r.left) / (r.width / 37));
    const off = +g.dataset.offset, nd = +g.dataset.ndays;
    const dayN = col - off + 1;
    if (col < 0 || col > 36 || dayN < 1 || dayN > nd) return null;
    return isoOf(+g.dataset.year, +g.dataset.month, dayN);
  }

  /* ---- drag-to-create (a plain click is a one-day drag) ---- */
  const monthsEl2 = document.getElementById("months");
  let dragA = null, dragB = null, dragging = false;
  function paintRange(){
    monthsEl2.querySelectorAll(".day.rangesel").forEach(x => x.classList.remove("rangesel"));
    if (!dragging || !dragA || !dragB) return;
    const lo = dragA < dragB ? dragA : dragB, hi = dragA < dragB ? dragB : dragA;
    monthsEl2.querySelectorAll(".day.in").forEach(d => {
      if (d.dataset.date >= lo && d.dataset.date <= hi) d.classList.add("rangesel");
    });
  }
  function endDrag(commit){
    if (dragging && commit && dragA && dragB){
      const lo = dragA < dragB ? dragA : dragB, hi = dragA < dragB ? dragB : dragA;
      createOn(lo, hi, armed);
    }
    dragging = false; dragA = dragB = null;
    paintRange();
  }
  monthsEl2.addEventListener("mousedown", e => {
    if (armed === null || e.button !== 0) return;
    // v2.0.9: existing bars are never creation targets - pressing one while a
    // category is armed must OPEN the item (the click handler below), not
    // start a drag on the day beneath it. Creation on a busy day still works
    // from any empty cell space; a drag simply cannot START on a bar.
    if (e.target.closest(".bar")) return;
    const d = dateAt(e.target, e.clientX);
    if (!d) return;
    e.preventDefault();
    dragging = true; dragA = dragB = d;
    paintRange();
  });
  monthsEl2.addEventListener("mousemove", e => {
    if (!dragging) return;
    const d = dateAt(e.target, e.clientX);
    if (d && d !== dragB){ dragB = d; paintRange(); }
  });
  document.addEventListener("mouseup", e => {
    if (dragging) endDrag(true);
  });

  /* ---- click a bar: open the item in Outlook. Works while a category is
     armed too (v2.0.9) - the tooltip promises "click to open", and armed mode
     governs only empty day space. No double-action risk: the drag flow above
     ignores bars, so a bar click never also creates. ---- */
  monthsEl2.addEventListener("click", e => {
    const bar = e.target.closest(".bar");
    if (!bar) return;
    openItem(bar.dataset.id, bar.dataset.start, bar.textContent);
  });
  // v2.0.10: item actions issued while the bridge is mid-build (the
  // "Outlook calendar changed" window) received the quick-serve's
  // {"ok":false,"busy":true} and surfaced as "Outlook error: unknown". The
  // quick-serve rejects BEFORE touching Outlook, so busy requests are safe to
  // retry; this helper waits the build out (bounded), with an honest toast,
  // and returns the first non-busy JSON for the caller to handle as before.
  async function apiWaitBusy(path, waitMsg){
    pendingActions++;
    try {
      let told = false, waits = 0;
      for (let n = 0; n < 40; n++){          // hard ceiling across all wait kinds
        let r = null;
        try { r = await api(path); }
        catch(_){
          // v2.0.12: a connection-level rejection during the wait is the
          // build-finish transition (or a busy socket), not an outage - the
          // request never reached Outlook, so retrying is safe. Give up
          // quietly only if this page is actually unloading.
          if (leaving) return null;
          if (!told && waitMsg){ toast(waitMsg, 0); told = true; }
          await new Promise(res => setTimeout(res, 1500));
          continue;
        }
        let j = null;
        try { j = await r.json(); } catch(_){ throw 0; }
        if (j && j.retry){
          // v2.0.25: retry:true is a WAIT-OUT state, not a one-shot. The
          // bridge answered instantly because Outlook is starting (launched
          // visibly, never blocking the bridge) or a dead link was just
          // reset; either way the operation never reached Outlook and
          // repeating it is safe. A cold start takes tens of seconds, so
          // waiting beats reporting - budget ~45s, then say so honestly.
          if (++waits > 45){ toast("Outlook still isn\u2019t available \u2014 please try again once it is open"); return null; }
          toast((j.err || "Waiting for Outlook") + "\u2026", 0);
          await new Promise(res => setTimeout(res, 2000));
          continue;
        }
        if (!j || !j.busy) return j;
        if (++waits > 12){ break; }
        if (!told && waitMsg){ toast(waitMsg, 0); told = true; }
        await new Promise(res => setTimeout(res, 1500));
      }
      toast("YearView is still updating \u2014 please try again in a moment");
      return null;
    } finally { actionDone(); }
  }

  async function openItem(id, start, label){
    if (!SERVED){ toast("Opening items needs the YearView app \u2014 run the script with -Serve"); return; }
    if (!id){ toast("No Outlook ID stored for this item"); return; }
    try{
      const j = await apiWaitBusy("/open?id=" + encodeURIComponent(id) + "&start=" + encodeURIComponent(start||""),
        "Waiting for the update to finish \u2014 will open \u201C" + label + "\u201D\u2026");
      if (j === null) return;
      toast(j.ok ? "Opened \u201C" + label + "\u201D in Outlook" : ("Outlook error: " + (j.err || "unknown")));
    } catch(err){
      if (!leaving) toast("YearView is not responding \u2014 restart YearView");
    }
  }

  /* ---- right-click a bar: context menu with open/delete ---- */
  const ctx = document.getElementById("ctx");
  let ctxBar = null;
  let ctxDay = null;                              // v2.1.12: day under a blank-space right-click
  function hideCtx(){ ctx.style.display = "none"; ctxBar = null; ctxDay = null; fpanel.style.display = "none"; }
  document.addEventListener("click", hideCtx);
  window.addEventListener("blur", hideCtx);
  window.addEventListener("scroll", hideCtx, true);
  document.addEventListener("keydown", e => { if (e.key === "Escape") hideCtx(); });
  monthsEl2.addEventListener("contextmenu", e => {
    const bar = e.target.closest(".bar");
    if (!bar){
      // v2.1.12: right-click on a blank part of an in-month day offers to
      // open the Outlook calendar there. Out-of-month cells and the rest of
      // the page keep the normal browser menu.
      const day = e.target.closest(".day");
      if (!day || !day.dataset.date) return;
      e.preventDefault();
      ctxBar = null; ctxDay = day.dataset.date;
      ctx.innerHTML = '<div data-act="goto">Open Outlook calendar here</div>';
      ctx.style.display = "block";
      ctx.style.left = Math.min(e.clientX, innerWidth - ctx.offsetWidth - 8) + "px";
      ctx.style.top  = Math.min(e.clientY, innerHeight - ctx.offsetHeight - 8) + "px";
      return;
    }
    e.preventDefault();
    ctxBar = bar;
    const rec = bar.dataset.rec === "1";
    ctx.innerHTML = '<div data-act="open">Open in Outlook</div>' + (rec
      ? '<div class="danger" data-act="delocc">Delete this occurrence</div>' +
        '<div class="danger" data-act="delseries">Delete entire series</div>'
      : '<div class="danger" data-act="del">Delete</div>');
    ctx.style.display = "block";
    ctx.style.left = Math.min(e.clientX, innerWidth - ctx.offsetWidth - 8) + "px";
    ctx.style.top  = Math.min(e.clientY, innerHeight - ctx.offsetHeight - 8) + "px";
  });
  ctx.addEventListener("click", e => {
    e.stopPropagation();
    const act = e.target.dataset.act;
    const bar = ctxBar;
    const dayD = ctxDay;
    hideCtx();
    if (act === "goto" && dayD){ gotoDate(dayD); return; }
    if (!act || !bar) return;
    const label = bar.textContent;
    if (act === "open"){ openItem(bar.dataset.id, bar.dataset.start, label); return; }
    if (!SERVED){ toast("Deleting needs the YearView app \u2014 run the script with -Serve"); return; }
    const mode = act === "delseries" ? "series" : (act === "delocc" ? "occ" : "item");
    deleteItem(bar.dataset.id, bar.dataset.start, mode, label);
  });
  async function gotoDate(d){
    // v2.1.12: navigate Outlook's own calendar to the clicked date.
    if (!SERVED){ toast("Needs the YearView app \u2014 run the script with -Serve"); return; }
    try{
      const j = await apiWaitBusy("/goto?date=" + encodeURIComponent(d),
        "Opening the Outlook calendar\u2026");
      if (j === null) return;
      if (j.ok) toast("Outlook calendar opened at " + d);
      else toast("Outlook error: " + (j.err || "unknown"));
    } catch(err){
      if (!leaving) toast("YearView is not responding \u2014 restart YearView from the shortcut");
    }
  }
  // v2.1.32: our own confirmation, because a browser's confirm() is headed
  // with the page's ORIGIN - "127.0.0.1:8377 says" - which script cannot
  // change and which tells the user nothing about what is asking. This one is
  // headed "YearView". Cancel is focused by default and Escape cancels, so
  // the safe answer is the one a stray keypress gives.
  const askov = document.createElement("div");
  askov.className = "helpov";
  askov.innerHTML =
    '<div class="helpbox askbox">' +
      '<div class="helphead">YearView</div>' +
      '<div class="askmsg" id="askmsg"></div>' +
      '<div class="askbtns">' +
        '<button id="askNo">Cancel</button>' +
        '<button id="askYes" class="danger">Delete</button>' +
      '</div>' +
    '</div>';
  document.body.appendChild(askov);
  function yvConfirm(msg){
    return new Promise(resolve => {
      askov.querySelector("#askmsg").textContent = msg;
      askov.classList.add("show");
      const yes = askov.querySelector("#askYes");
      const no  = askov.querySelector("#askNo");
      const done = answer => {
        askov.classList.remove("show");
        yes.onclick = no.onclick = null;
        document.removeEventListener("keydown", key, true);
        resolve(answer);
      };
      // Capture phase: the page's other Escape handlers must not also fire
      // (they disarm the category and close menus) while this is open.
      const key = e => {
        if (e.key === "Escape"){ e.stopPropagation(); done(false); }
        else if (e.key === "Enter"){ e.stopPropagation(); done(document.activeElement === yes); }
      };
      document.addEventListener("keydown", key, true);
      yes.onclick = () => done(true);
      no.onclick  = () => done(false);
      askov.onclick = e => { if (e.target === askov) done(false); };
      no.focus();
    });
  }
  async function deleteItem(id, start, mode, label){
    // v2.1.29: confirmation is ON by default and can be turned off in
    // Settings. Deleted items do go to Outlook's Deleted Items and are
    // recoverable, which is why deleting straight away was defensible - but
    // the click target is a small bar in a dense grid, and the default should
    // be the one that cannot lose work by a slip.
    if (confirmDelete){
      const what = mode === "occ" ? "this occurrence of \u201C" + label + "\u201D"
                 : mode === "series" ? "the entire series \u201C" + label + "\u201D"
                 : "\u201C" + label + "\u201D";
      if (!await yvConfirm("Delete " + what + "?\n\nIt will go to Outlook's Deleted Items.")) return;
    }
    try{
      const j = await apiWaitBusy("/delete?id=" + encodeURIComponent(id) +
        "&start=" + encodeURIComponent(start||"") + "&mode=" + mode,
        "Waiting for the update to finish \u2014 will delete \u201C" + label + "\u201D\u2026");
      if (j === null) return;
      if (j.ok){
        toast("Deleted \u201C" + label + "\u201D \u2014 updating\u2026", 0);
        setTimeout(rebuild, 700);
      } else {
        toast("Outlook error: " + (j.err || "unknown"));
      }
    } catch(err){
      if (!leaving) toast("YearView is not responding \u2014 restart YearView");
    }
  }

  async function createOn(dateStr, endStr, cat){
    const range = endStr && endStr !== dateStr;
    if (SERVED){
      try{
        const j = await apiWaitBusy("/new?date=" + dateStr + (range ? "&end=" + endStr : "") +
          "&cat=" + encodeURIComponent(cat),
          "Waiting for the update to finish \u2014 will create the item\u2026");
        if (j === null) return;
        if (j.ok) toast((cat === "" ? "New uncategorised item opened in Outlook"
                                    : "New \u201C" + cat + "\u201D item opened in Outlook") + " \u2014 " +
          dateStr + (range ? " \u2192 " + endStr : ""));
        else toast("Outlook error: " + (j.err || "unknown"));
      } catch(err){
        if (!leaving) toast("YearView is not responding \u2014 restart YearView");
      }
    } else {
      // static file mode: hand Outlook an ICS instead
      const [y,mo,d] = (endStr || dateStr).split("-").map(Number);
      const nx = new Date(y, mo-1, d+1);
      const fmt = dt => (typeof dt === "string") ? dt.replace(/-/g,"")
        : dt.getFullYear() + String(dt.getMonth()+1).padStart(2,"0") + String(dt.getDate()).padStart(2,"0");
      const stamp = new Date().toISOString().replace(/[-:]/g,"").replace(/\..+/,"") + "Z";
      const ics = ["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//YearView//EN","BEGIN:VEVENT",
        "UID:yv-" + Date.now() + "@local", "DTSTAMP:" + stamp,
        "DTSTART;VALUE=DATE:" + fmt(dateStr), "DTEND;VALUE=DATE:" + fmt(nx),
        "SUMMARY:New appointment",
        "CATEGORIES:" + cat.replace(/([,;\\])/g, "\\$1"),
        "END:VEVENT","END:VCALENDAR"].join("\r\n");
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([ics], {type:"text/calendar"}));
      a.download = "new-" + dateStr + ".ics";
      a.click();
      toast("ICS downloaded \u2014 open it to add the item (or run with -Serve for one-click)");
    }
  }
  /* ---- help overlay ---- */
  const helpov = document.getElementById("helpov");
  document.getElementById("helpBtn").onclick = e => { e.stopPropagation(); hideCtx(); helpov.classList.add("show"); };
  document.getElementById("helpx").onclick = () => helpov.classList.remove("show");
  helpov.addEventListener("click", e => { if (e.target === helpov) helpov.classList.remove("show"); });
  document.addEventListener("keydown", e => { if (e.key === "Escape") helpov.classList.remove("show"); });

  /* ---- v3.1.1: pinned header and legend ---- */
  // The weekday strip has to sit directly beneath the header, and the toast has
  // to clear the legend - but both heights depend on the font, the display
  // scaling and how the stamp and the category chips wrap, so they are MEASURED
  // rather than assumed, and re-measured whenever that can change.
  const hdrEl = document.querySelector("header.yv");
  const legEl = document.getElementById("legend");
  function syncPinned(){
    const rs = document.documentElement.style;
    if (hdrEl) rs.setProperty("--headh", hdrEl.offsetHeight + "px");
    if (legEl) rs.setProperty("--footh", legEl.offsetHeight + "px");
  }
  syncPinned();
  setTimeout(syncPinned, 0);            // once the stamp text has landed
  addEventListener("resize", syncPinned);

  /* ---- v3.1.0: continuous scrolling across years ---- */
  // Years are loaded on demand as either end of the timeline comes into
  // reach, appended or prepended to the rendered blocks, and trimmed once
  // more than KEEP_YEARS are in memory. Every year costs a FULL Outlook read
  // (tens of seconds on a slow machine), so loading is always in the
  // background with a progress toast and never blocks what is on screen.
  const yrlabel = document.getElementById("yrlabel");
  // v3.1.4: four years kept, and the reach shortened. With three years and a
  // 900px reach, a tall window could be within reach of BOTH ends at once -
  // the scrollable range is only a couple of thousand pixels - so releasing a
  // year at one end immediately brought the other end into reach. That is the
  // geometry behind the loading loop; the guards below close the rest of it.
  const KEEP_YEARS = 4;
  const NEAR_PX = 400;          // how close to an end before the next year is fetched
  const TRAVEL_PX = 200;        // how far the reader must move before another fetch
  const loadingYears = new Set();
  let lastReqAt = 0;
  let topAtLastLoad = null;     // scroll position when a year last arrived

  const yearBlocks = () => [...monthsEl.querySelectorAll(".yearblock")];
  // Splice a year in or out without touching anything already rendered, and
  // hold a month row still across the change by MEASURING where it ends up -
  // not by assuming how far it moved. Adding an assumed height was right only
  // while nothing else touched the scroll position, which stopped being true
  // the moment the surrounding blocks survived (see the overflow-anchor note
  // in the stylesheet). Measuring the outcome is right either way, and stays
  // right if an inserted block is not the height it was expected to be. The
  // read forces layout but no paint, so the correction lands in the same task
  // as the change and the intermediate state is never shown.
  function holdPlace(exclude, mutate){
    const el = document.scrollingElement || document.documentElement;
    let anchor = null, was = 0;
    for (const mo of monthsEl.querySelectorAll(".month")){
      if (exclude && exclude.contains(mo)) continue;   // it is about to be removed
      const r = mo.getBoundingClientRect();
      if (r.bottom > 0){ anchor = mo; was = r.top; break; }
    }
    mutate();
    if (!anchor || !anchor.isConnected){ lastTop = el.scrollTop; return; }
    const now = anchor.getBoundingClientRect().top;
    if (now !== was) el.scrollTop += now - was;
    // v3.1.4: this correction moves the scroll position OURSELVES, and the
    // scroll event it raises is indistinguishable from the reader moving. Left
    // unrecorded, a correction downwards read as "scrolling down" and asked for
    // the year ahead, whose trim then read as "scrolling up" and asked for the
    // year behind - the loading loop, ping-ponging between the two ends.
    lastTop = el.scrollTop;
  }
  function insertYearBlock(y){
    const before = yearBlocks().find(b => +b.dataset.year > y) || null;
    const yb = buildYearBlock(y, visibleEvents());
    holdPlace(null, () => monthsEl.insertBefore(yb, before));
  }
  // v3.1.4: releasing a block that is still close to the view is what re-arms
  // the opposite end. A block is only released once it is a full window clear
  // of the viewport; otherwise it is kept, and the caller stops trimming.
  function dropYearBlock(y){
    const yb = yearBlocks().find(b => +b.dataset.year === y);
    if (!yb){ years.delete(y); return true; }
    const r = yb.getBoundingClientRect();
    if (r.bottom > -innerHeight && r.top < innerHeight * 2) return false;
    years.delete(y);
    holdPlace(yb, () => yb.remove());
    return true;
  }

  // A neighbouring year can use a category the home year never mentions. Add
  // it to the legend (in the same alphabetical order) and rebuild the filter
  // panel, so it can be seen and hidden like any other. The armed chip is
  // untouched: chips are inserted, never replaced.
  function mergeYearCategories(payload){
    const have = new Set(cats.map(c => c.name));
    let added = false;
    for (const c of (payload.categories || [])){
      if (!c.name || have.has(c.name)) continue;
      cats.push({ name: c.name, col: c.col });
      have.add(c.name);
      added = true;
      const chip = document.createElement("span");
      chip.className = "item";
      chip.dataset.cat = c.name;
      chip.innerHTML = '<span class="sw" style="background:' + safeCol(c.col) + '"></span>' + esc(c.name);
      const after = [...legend.querySelectorAll(".item")].find(
        x => x.dataset.cat && catsCmp({ name: x.dataset.cat }, c) > 0);
      legend.insertBefore(chip, after || noneChip);
    }
    if (added){ buildFilterPanel(); syncFilterUI(); syncPinned(); }
  }

  // An appointment spanning New Year is in BOTH adjacent years' payloads -
  // each year's Outlook restriction is "starts before this year ends AND ends
  // after it begins", which that item satisfies twice over - so combining the
  // payloads naively drew it twice, in two lanes, in the months either side of
  // the boundary. Keyed on the id AND the span, because the occurrences of a
  // recurring series share one id.
  function recombine(){
    const seen = new Set();
    allEvents = [];
    for (const y of yearsAsc()){
      for (const ev of (years.get(y).events || [])){
        const k = ev.id + "\u001f" + ev.s + "\u001f" + ev.e;
        if (seen.has(k)) continue;
        seen.add(k);
        allEvents.push(ev);
      }
    }
  }
  // Drop from the END OPPOSITE the arrival. Dropping "whichever is furthest
  // from the year in view" looks right and is not: while the reader sits in a
  // year and the timeline grows forwards, the furthest year IS the one just
  // fetched, so it would be discarded the moment it arrived and fetched again
  // on the next scroll - tens of seconds of Outlook reading, repeatedly, for
  // nothing. The year in view is never dropped.
  function trimYears(added){
    while (years.size > KEEP_YEARS){
      const ks = yearsAsc();
      const drop = (added > ks[0]) ? ks[0] : ks[ks.length - 1];
      if (drop === focusYear || drop === added) break;
      if (!dropYearBlock(drop)) break;      // still too close to the view
    }
    recombine();
  }
  function loadYear(y, force){
    if (!SERVED || !continuous) return;
    if (years.has(y) || loadingYears.has(y) || loadingYears.size) return;
    if (y < 1990 || y > 2100) return;                 // F-2: match the server clamp
    // Never fetch a year further away than the window we keep. Without this,
    // a window tall enough to reach the far end while the reader stays put
    // would keep asking for one more year for ever.
    if (Math.abs(y - focusYear) > KEEP_YEARS - 1) return;
    if (rebuilding || bgUpdating || leaving || dragging) return;
    // v3.1.4: and the reader must have MOVED since the last year arrived. The
    // position is recorded AFTER our own corrections, so only genuine scrolling
    // counts - which is what makes this terminate: a timeline nobody is
    // scrolling cannot ask for another year, however the geometry falls.
    // The distance is capped by what there is to scroll, because a tall window
    // holding two short years may have less range than the threshold - and a
    // requirement that cannot be met would strand the timeline instead.
    //
    // v3.1.9: `force` skips it. At the very top there is nowhere to travel
    // BACKWARDS to, so a reader asking for the earlier year could never meet a
    // movement requirement - and the sustained upward gesture that is the only
    // caller of `force` is already the evidence of intent this guard exists to
    // establish. It cannot loop: the prepend leaves the reader a year clear of
    // the top, so asking again means scrolling back up there first.
    const sel = document.scrollingElement || document.documentElement;
    const room = Math.max(0, sel.scrollHeight - sel.clientHeight);
    const need = Math.min(TRAVEL_PX, room / 2);
    if (!force && topAtLastLoad !== null && Math.abs(sel.scrollTop - topAtLastLoad) < need) return;
    loadingYears.add(y);
    toast("Loading " + y + "\u2026", 0);
    api("/data?year=" + y)
      .then(r => r.json())
      .then(b => {
        loadingYears.delete(y);
        if (leaving) return;
        // v3.1.5: the setting can be turned off while a year is in flight;
        // without this the reply would splice that year back into a view
        // that has just collapsed to one.
        if (!continuous){ hideToast(); return; }
        if (b && b.closing){ hideToast(); return; }
        if (b && b.busy){ hideToast(); return; }      // a build is running; the next scroll asks again
        if (!b || !b.ok || !b.data){
          toast("Could not load " + y);
          return;
        }
        const payload = b.data;
        payload.events = [].concat(payload.events || []);
        payload.categories = [].concat(payload.categories || []);
        years.set(y, payload);
        recombine();
        mergeYearCategories(payload);
        insertYearBlock(y);
        trimYears(y);
        topAtLastLoad = (document.scrollingElement || document.documentElement).scrollTop;
        hideToast();
      })
      .catch(() => {
        loadingYears.delete(y);
        if (!leaving) toast("Could not load " + y);
      });
  }

  // The header year follows whichever year covers the middle of the window,
  // so it names what is being looked at rather than what the page was built
  // for. focusYear also decides which year prints and which is kept on trim.
  function updateFocusYear(){
    const mid = innerHeight / 2;
    let best = null;
    for (const yb of monthsEl.querySelectorAll(".yearblock")){
      const r = yb.getBoundingClientRect();
      if (r.top <= mid && r.bottom >= mid){ best = +yb.dataset.year; break; }
      if (best === null && r.bottom > 0) best = +yb.dataset.year;
    }
    if (best === null || best === focusYear) return;
    focusYear = best;
    if (yrlabel) yrlabel.textContent = focusYear;
    document.title = focusYear + " YearView";
  }
  // Load in the DIRECTION OF TRAVEL only. Each year is a full Outlook read, so
  // fetching the year behind the reader while they are heading forwards spends
  // tens of seconds on something they may never look at.
  //
  // v3.1.9: the direction comes from the READER'S INPUT first, and from the
  // change in scroll position only as a fallback. Taking it from the position
  // alone was wrong in exactly the situation that matters most: a reader
  // beginning to scroll forwards is BY DEFINITION near the top, where the
  // backward trigger is armed, and any single pixel of apparent upward
  // movement there - trackpad inertia, a scrollbar nudge, a sub-pixel
  // correction of our own - read as "going backwards" and fetched the year
  // BEHIND. A wheel notch says which way the reader is going and cannot be
  // misread; the position fallback carries touch and keyboard, with a small
  // dead zone so jitter decides nothing.
  let lastTop = 0, inputDir = 0, inputAt = 0;
  function travelDir(dy){
    if (inputDir && Date.now() - inputAt < 600) return inputDir;   // the reader just said
    if (Math.abs(dy) >= 4) return dy > 0 ? 1 : -1;
    return 0;                                                      // jitter: decide nothing
  }
  function onTimelineScroll(){
    if (!continuous) return;
    updateFocusYear();
    const el = document.scrollingElement || document.documentElement;
    const dir = travelDir(el.scrollTop - lastTop);
    lastTop = el.scrollTop;
    if (!dir) return;
    const now = Date.now();
    if (now - lastReqAt < 400) return;
    lastReqAt = now;
    const ks = yearsAsc();
    if (dir > 0 && el.scrollTop + el.clientHeight >= el.scrollHeight - NEAR_PX){
      loadYear(ks[ks.length - 1] + 1);
    } else if (dir < 0 && el.scrollTop <= NEAR_PX){
      loadYear(ks[0] - 1);
    }
  }
  // Scroll events arrive far faster than frames, and this handler measures
  // elements - so answering every one forced a layout per event and made the
  // scroll itself stutter. One response per animation frame is all that can
  // ever be seen.
  let scrollQueued = false;
  function onScrollFrame(){
    if (scrollQueued) return;
    scrollQueued = true;
    requestAnimationFrame(() => { scrollQueued = false; onTimelineScroll(); });
  }
  addEventListener("scroll", onScrollFrame, { passive: true });
  addEventListener("resize", onScrollFrame, { passive: true });
  // The wheel is the plainest statement of which way the reader is going, so
  // it is recorded here and trusted by travelDir above. It also covers the one
  // case no scroll event can: at the very top there is nothing left to scroll,
  // so the year BEFORE this one could never otherwise be asked for.
  let upAcc = 0, upAt = 0;
  addEventListener("wheel", e => {
    if (!continuous || e.ctrlKey || leaving || rebuilding || dragging) return;
    if (!e.deltaY) return;
    inputDir = e.deltaY > 0 ? 1 : -1;
    inputAt = Date.now();
    const el = document.scrollingElement || document.documentElement;
    // A DELIBERATE upward gesture, not any upward number. A mouse or a
    // trackpad commonly emits a small negative delta at the start of a
    // gesture, and a reader about to scroll forwards is sitting at the very
    // top - so a single stray flick was enough to fetch the year BEHIND,
    // ahead of the year they were heading towards. The upward movement now
    // has to accumulate, and any forward movement or a pause clears it.
    if (inputDir > 0 || inputAt - upAt > 600) upAcc = 0;
    upAt = inputAt;
    if (inputDir < 0) upAcc += -e.deltaY;
    if (upAcc >= 150 && el.scrollTop <= 1){
      upAcc = 0;
      loadYear(yearsAsc()[0] - 1, true);
    }
  }, { passive: true });

  // Turning the setting off collapses the timeline back to the year in view;
  // turning it on starts fetching a neighbour. Neither needs a reload.
  function setContinuous(on){
    continuous = on;
    if (!on){
      const keep = years.has(focusYear) ? focusYear : YEAR;
      for (const y of yearsAsc()) if (y !== keep) years.delete(y);
      recombine();
      renderMonths();
      focusYear = keep;
      if (yrlabel) yrlabel.textContent = keep;
      scrollTo(0, 0);
    } else {
      renderMonths();                     // v3.1.5: the year bands appear with it
      setTimeout(onTimelineScroll, 50);
    }
  }
  // A window tall enough to show the whole year has NO scrollbar, so no scroll
  // event can ever fire and nothing would load. Reach forward once, after the
  // first paint settles, so there is always somewhere to scroll to.
  if (continuous) setTimeout(() => {
    if (!continuous || leaving) return;
    updateFocusYear();
    loadYear(yearsAsc()[yearsAsc().length - 1] + 1);
  }, 1200);

  /* ---- v3.4.0: a year view that starts at a chosen month ---- */
  // The heading stops saying "Year 2026" and says the span instead, because
  // that is what is on screen - twelve months from one year into the next.
  function updateHeading(){
    const yw = document.getElementById("yrword");
    const yl = document.getElementById("yrlabel");
    if (startMonth > 0){
      if (yw) yw.style.display = "none";
      if (yl) yl.textContent = MONTHS[startMonth] + " " + YEAR + " \u2013 " +
                               MONTHS[(startMonth + 11) % 12] + " " + (YEAR + 1);
      document.title = MONTHS[startMonth] + " " + YEAR + " YearView";
    } else {
      if (yw) yw.style.display = "";
      if (yl) yl.textContent = YEAR;
      document.title = YEAR + " YearView";
    }
  }
  // The spill months belong to the FOLLOWING year, whose events this page was
  // never given. Fetching them is a full calendar read on the bridge, so it
  // happens in the background after the first paint: this year's months are on
  // screen immediately and the remainder fill in when they arrive. A build in
  // progress answers busy, so it waits that out rather than giving up.
  let spillTries = 0;
  function ensureSpillYear(){
    const y = YEAR + 1;
    if (!SERVED || startMonth <= 0 || years.has(y) || leaving) return;
    // v3.4.1: the message only appears if the answer is slow enough to need
    // one. A cached payload comes back in milliseconds, and a toast that
    // flashes up and vanishes reads as a glitch rather than as progress.
    const spillMsg = setTimeout(() => toast("Loading " + y + "\u2026", 0), 400);
    api("/data?year=" + y).then(r => r.json()).then(b => {
      clearTimeout(spillMsg);
      if (leaving || startMonth <= 0) { hideToast(); return; }
      if (b && (b.busy || b.closing)){
        if (++spillTries <= 30){ setTimeout(ensureSpillYear, 2000); return; }
        hideToast(); return;
      }
      if (!b || !b.ok || !b.data){ toast("Could not load " + y); return; }
      const p = b.data;
      p.events = [].concat(p.events || []);
      p.categories = [].concat(p.categories || []);
      years.set(y, p);
      recombine();               // dedupes the item that spans New Year
      mergeYearCategories(p);    // a category used only next year joins the legend
      renderMonths();
      hideToast();
    }).catch(() => {
      clearTimeout(spillMsg);          // the failure path must cancel it too
      if (!leaving) toast("Could not load " + y);
    });
  }
  function applyStartMonth(){
    startMonth = resolveStartMonth();
    updateHeading();
    renderMonths();
    if (startMonth > 0) ensureSpillYear();
    scrollTo(0, 0);
  }
  updateHeading();
  // The legend and the filter panel are built after the first render, so a
  // category belonging only to the spill year has to join them here rather
  // than at seed time - it would have been a dead-zone error up there.
  if (DATA2 && years.has(YEAR + 1)) mergeYearCategories(DATA2);
  // Only reaches the bridge when the page did not already carry the year.
  if (startMonth > 0) setTimeout(ensureSpillYear, 900);

  /* ---- print: scale to fit a single page ---- */
  window.addEventListener("beforeprint", () => {
    const wrap = document.querySelector(".wrap");
    wrap.style.zoom = "";
    // v3.4.8: drop the legend FIRST, so the height measured below is the
    // height that will actually print. Removed again at afterprint. It has to
    // come before the early return too: "as many as needed" prints no legend
    // either when the setting is off.
    document.documentElement.classList.toggle("nolegend", !printLegend);
    // v3.2.0: how many sheets to scale onto is now the reader's choice, and it
    // applies to whatever is on the page - one year or a timeline. Zero means
    // "as many as it takes": nothing is scaled and the page flows, which is
    // v3.1.6's behaviour kept as an option.
    //
    // Nothing is ever HIDDEN here. v3.1.0 restricted the sheet to the year in
    // view and scaled to fit, but measured the height while the print
    // stylesheet was not yet in force - so the years it had just hidden were
    // still counted and the survivor was shrunk against a height several times
    // its own, and a fraction of a year reached the paper.
    const pages = (printPages > 0 && printPages <= 4) ? printPages : 0;
    if (!pages) return;                     // flows onto as many sheets as it needs
    const MM = 96 / 25.4;
    // v3.3.4: an ASSUMED sheet - A4 LANDSCAPE. Paper and orientation belong to
    // the print dialog, which will not tell a page what was chosen and will
    // not accept a page saying otherwise, so fitting to a number of sheets has
    // to assume one; A4 is the sheet most printers hold, and a year is a wide
    // thing, so landscape. Printing to PDF and enlarging is an exact scale-up
    // from here. On any other paper the count will be out, and the dialog's
    // own Scale control is the remedy.
    //
    // Worth knowing when reading old numbers: A4 portrait and the A3 landscape
    // assumed until now are BOTH 297mm tall, so every release up to v3.3.3
    // was, by coincidence, scaling correctly for A4 portrait as well.
    const pageH = (210 - 16) * MM;          // A4 landscape printable height, less 8mm margins
    const h = wrap.scrollHeight;
    const z = Math.min(1, (pageH * pages / h) * 0.97);   // safety factor covers legend reflow at page width (when printed)
    if (z < 1) wrap.style.zoom = z;   // zoom reflows to full width while scaling height
  });
  window.addEventListener("afterprint", () => {
    document.querySelector(".wrap").style.zoom = "";
    document.documentElement.classList.remove("nolegend");
  });

  /* ---- auto-refresh: poll the bridge for calendar changes ---- */
  // arrival from a silent edit-driven rebuild: announce it, then scrub the
  // marker so reload/bookmark of this URL doesn't re-toast
  try {
    const uq = new URLSearchParams(location.search);
    if (uq.get("updated")){
      toast("YearView updated");
      uq.delete("updated");
      history.replaceState(null, "", location.pathname + "?" + uq.toString());
    }
  } catch(e){}

  if (SERVED && BUILD_TS !== "0"){
    // v1.2.30: the dirty baseline MUST advance after a clean revalidation.
    // It was the page's constant build stamp, so a quiet rebuild that changed
    // nothing left lastChangeTicks forever ahead of it - fp-dirt re-fired
    // every poll: an endless silent-rebuild loop (visible since the v1.2.21
    // progress toast as "Updating calendar…" repeating every few seconds, and
    // before that as an unexplained busy bridge).
    let sinceTs = BUILD_TS;
    // v2.1.12: ONE revalidation routine, three callers - the fingerprint
    // dirty signal (as before), a run shortly after page load, and a
    // periodic timer. The latter two exist because change detection is
    // EVENT-driven: anything added, edited or deleted while the bridge was
    // OFFLINE raised no event, so the cache stayed stale until something
    // forced a build. `cont` is the poll continuation: the dirty-poll caller
    // passes its own next step so the chain carries on exactly as before;
    // timer callers pass null - scheduling the poll from them would start a
    // SECOND poll chain.
    function quietRevalidate(cont){
      if (dragging || rebuilding || bgUpdating || leaving) return;
      bgUpdating = true;
      toast("Updating YearView\u2026", 0);
      api("/build?year=" + YEAR)
        .then(r => r.json())
        .then(b => {
          bgUpdating = false;
          if (leaving) return;                  // user navigation already committed
          if (b && b.closing){ hideToast(); return; }
          if (b && b.ok && b.changed){
            navTo("/?year=" + YEAR + "&updated=1" + navSuffix());
          } else {
            if (b && b.ts) sinceTs = b.ts;      // retire the revalidated stamp (v1.2.30)
            toastEl.classList.remove("show");   // nothing changed: dismiss quietly
            if (cont) setTimeout(cont, 2000);
          }
        })
        .catch(() => {
          bgUpdating = false;
          if (!leaving) toastEl.classList.remove("show");
          if (cont && !leaving) setTimeout(cont, 10000);
        });
    }
    // Startup freshness: any load that did not just arrive FROM a rebuild
    // (those carry ?updated=1) revalidates a few seconds after paint - this
    // is what catches changes made while YearView was offline. The interval
    // is belt-and-braces for anything the event stream ever misses.
    if (!(new URLSearchParams(location.search).get("updated"))){
      setTimeout(() => quietRevalidate(null), 4000);
    }
    setInterval(() => quietRevalidate(null), 600000);
    (function pollDirty(){
      api("/dirty?since=" + sinceTs)
        .then(r => r.json())
        .then(j => {
          if (j && j.closing){
            // v2.1.6: the bridge is closing this session because Outlook is
            // gone. Set the same flag a navigation sets, which suppresses the
            // "bridge not reachable" toasts everywhere, stop polling, and say
            // what is happening. The dialog explains the rest.
            // v2.1.7: no toast. The dialog that follows says everything, and a
            // toast alongside it reads as a second, contradictory answer.
            leaving = true;
            hideToast();
            return;
          }
          if (j.dirty && !dragging && !rebuilding && !bgUpdating){
            if (j.fp){
              // Fingerprint-only dirt: possibly just the store's sync echo
              // re-touching LastModificationTime after a save. Rebuild in the
              // BACKGROUND (v1.2.22: bgUpdating, not rebuilding - so the year
              // arrows stay usable while "Updating calendar…" shows). If the
              // content actually changed, refresh - unless the user has
              // already navigated away, in which case their navigation wins:
              // the year they are heading to revalidates itself against the
              // same change stamp on arrival (v1.2.25), so nothing is lost.
              quietRevalidate(pollDirty);           // v2.1.12: shared routine
              return;
            }
            toast("Outlook calendar changed \u2014 updating YearView\u2026", 0);
            setTimeout(rebuild, 600);
            return;   // navigation ends this page's polling
          }
          setTimeout(pollDirty, 2000);   // 2 s: with the per-poll fingerprint,
                                         // this bounds edit-detection latency
        })
        .catch(() => setTimeout(pollDirty, 10000));
    })();
  }
})();
</script>
</body>
</html>
'@

function Read-YearClassic([object]$cal, [string]$filter) {
    # The original, reference reader: full Items collection, Sort, expand
    # recurrences, Restrict to the year window, materialise every item.
    $items = $cal.Items
    $items.Sort('[Start]')           # ascending sort REQUIRED before IncludeRecurrences
    $items.IncludeRecurrences = $true
    $restricted = $items.Restrict($filter)
    $list = New-Object System.Collections.Generic.List[object]
    $n = 0; $trunc = $false
    foreach ($it in $restricted) {
        Invoke-QuickServe   # keep page GETs responsive during this long read
        # v2.1.20: deliberately NO early exit on Quit. v2.1.19 broke out of this
        # loop when Quit arrived, which left the item list SHORT - and the build
        # went on to serialise that partial list, render it and SAVE IT TO THE
        # CACHE, so the next start opened a half-built year. A read either
        # completes or it is not worth having. Quit has already been answered
        # and the window already closed by the time we get here; finishing the
        # read costs a few invisible seconds and leaves a whole, correct cache
        # behind. The main loop exits on the flag as soon as the build returns.
        if ($n -ge 8000) { Write-Warning 'Stopped at 8000 items (safety cap).'; $trunc = $true; break }
        try {
            # v3.2.1: one property access, not two - the classic reader runs in
            # full on the FIRST build of every session as the validation half.
            $catRaw = $it.Categories
            $catStr = ''
            if ($catRaw) { $catStr = [string]$catRaw }
            $catStr = ((($catStr -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join '; ')
            $list.Add([pscustomobject]@{
                s   = $it.Start.ToString('yyyy-MM-ddTHH:mm:ss', $script:invCulture)
                e   = $it.End.ToString('yyyy-MM-ddTHH:mm:ss', $script:invCulture)
                t   = [string]$it.Subject
                ad  = [bool]$it.AllDayEvent
                cat = $catStr
                id  = [string]$it.EntryID
                r   = [bool]$it.IsRecurring
            })
            $n++
        } catch { Write-Verbose ("Skipped one item: {0}" -f $_.Exception.Message) }
    }
    @{ Events = $list; Truncated = $trunc }
}

function Read-YearFast([object]$cal, [string]$filter, [datetime]$rangeStart, [datetime]$rangeEnd) {
    $inv = $script:invCulture
    $daslRec = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8223000B'
    $daslAll = 'http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8215000B'
    $daslKw  = 'http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/Keywords'
    $list = New-Object System.Collections.Generic.List[object]
    $n = 0; $trunc = $false
    # v3.2.1: time the two halves SEPARATELY. "Fast reader, 11 seconds" says
    # nothing about WHERE the time went, and the two halves have completely
    # different costs: the Table half reads whole rows with no per-item COM
    # object, while the object-model half below pays a cross-process call for
    # every property of every item. Which half dominates depends entirely on
    # the MIX of items in the calendar, so it has to be measured, not assumed.
    $tSplit = [Diagnostics.Stopwatch]::StartNew()
    $tableMs = 0; $tableN = 0
    # v3.2.3: whether the one-call row read is ACTUALLY being used is now
    # measured rather than assumed. v3.2.2 introduced it and the next
    # measurement showed no gain - but the untouched half had slowed by the
    # same factor, so the machine was slower and the change was neither
    # confirmed nor cleared. A silent per-row fallback looks exactly like a
    # change that did nothing, so the two are now told apart.
    $bulkOk = $false; $bulkChecked = $false; $bulkRows = 0; $batchCalls = 0

    # --- TIMED non-recurring items via the Table API (no per-item COM object) ---
    # Field lesson from the v1.2.7 mismatch log: which timezone the built-in
    # 'Start'/'End' Table columns use VARIES BY OUTLOOK BUILD (the desktop
    # validated, the laptop returned UTC and every timed item drifted). So the
    # dates come from the DASL dtstart/dtend columns instead - documented UTC on
    # every build - converted locally with the same Windows timezone rules
    # Outlook itself applies. ALL-DAY items are excluded here entirely: they
    # have two storage conventions (local-midnight-as-UTC for user-created
    # items, UTC-midnight for imported holidays - both visible in the log), so
    # no raw conversion reproduces Outlook's all-day normalisation; they take
    # the OM path below, where Outlook does its own normalising.
    # The DASL window is widened by 3 days each side purely as a coarse prefetch;
    # the EXACT year test happens below in PowerShell on the converted values, so
    # DASL date-literal semantics can never cause a boundary bug.
    $daslStart = 'urn:schemas:calendar:dtstart'
    $daslEnd   = 'urn:schemas:calendar:dtend'
    $tf = '@SQL="urn:schemas:calendar:dtstart" < ''{0}'' AND "urn:schemas:calendar:dtend" > ''{1}''' -f `
        $rangeEnd.AddDays(3).ToString('yyyy/MM/dd HH:mm', $inv), `
        $rangeStart.AddDays(-3).ToString('yyyy/MM/dd HH:mm', $inv)
    $tbl = $cal.GetTable($tf, 0)   # 0 = olUserItems
    # v3.2.2: an EXPLICIT column set in a KNOWN ORDER, so a whole row can be
    # fetched in one call and read by position. Keeping the default columns and
    # appending to them would leave the order dependent on the Outlook build -
    # the same class of build-to-build variation that made the built-in Start
    # and End columns unusable in the first place.
    $tbl.Columns.RemoveAll()
    [void]$tbl.Columns.Add($daslStart)   # 0
    [void]$tbl.Columns.Add($daslEnd)     # 1
    [void]$tbl.Columns.Add($daslRec)     # 2
    [void]$tbl.Columns.Add($daslAll)     # 3
    [void]$tbl.Columns.Add($daslKw)      # 4
    [void]$tbl.Columns.Add('Subject')    # 5
    [void]$tbl.Columns.Add('EntryID')    # 6
    # v3.2.4: fetch rows in BATCHES. The measurements settled what the cost
    # actually is: both halves of the reader land at about two and a half
    # milliseconds PER CROSS-PROCESS CALL on the slow machine, so the time is
    # proportional to the NUMBER of calls and nothing else. Reading a whole row
    # in one call (v3.2.2) still cost two calls a row - GetNextRow and then
    # GetValues. GetArray returns up to five hundred rows in ONE call, which
    # turns roughly fourteen hundred calls into three.
    #
    # Both paths below normalise to the SAME seven-element row, so the body
    # that interprets a row exists once and cannot drift between them.
    $ROWBATCH = 500
    $arrayOk = $true
    while (-not $tbl.EndOfTable) {
        Invoke-QuickServe   # keep page GETs responsive during this read
        # v2.1.20: no early exit on Quit - see the note in the long read above.
        $batch = New-Object System.Collections.Generic.List[object]
        if ($arrayOk) {
            $arr = $null
            try { $arr = $tbl.GetArray($ROWBATCH) } catch { $arr = $null }
            # Shape-checked, not trusted: two dimensions and exactly the seven
            # columns asked for. Bounds are read from the array itself because
            # a COM safe array is not obliged to start at zero.
            if ($arr -is [array] -and $arr.Rank -eq 2 -and
                ($arr.GetUpperBound(1) - $arr.GetLowerBound(1) + 1) -eq 7) {
                $cLo = $arr.GetLowerBound(1)
                for ($r = $arr.GetLowerBound(0); $r -le $arr.GetUpperBound(0); $r++) {
                    [void]$batch.Add(@($arr[$r,$cLo],     $arr[$r,$cLo+1], $arr[$r,$cLo+2],
                                       $arr[$r,$cLo+3],   $arr[$r,$cLo+4], $arr[$r,$cLo+5],
                                       $arr[$r,$cLo+6]))
                }
                $bulkRows += $batch.Count
                $batchCalls++
            } else {
                $arrayOk = $false
                Write-Warning 'Batched row read unavailable - reading rows one at a time.'
            }
        }
        if ($batch.Count -eq 0) {
            # Per-row fallback. Prove the POSITIONS on the first row before
            # trusting them: read two columns both ways and require agreement.
            # One row's extra cost, once, and the same validate-then-trust
            # pattern the reader already uses against the classic reader.
            $row = $tbl.GetNextRow()
            if (-not $bulkChecked) {
                $bulkChecked = $true
                try {
                    $probe = $row.GetValues()
                    if ($probe -is [array] -and $probe.Count -eq 7) {
                        $bulkOk = (([string]$probe[5] -eq [string]$row.Item('Subject')) -and
                                   ([string]$probe[6] -eq [string]$row.Item('EntryID')))
                    }
                } catch { $bulkOk = $false }
                if (-not $bulkOk) {
                    Write-Warning 'Row bulk read unavailable - reading columns individually.'
                }
            }
            $v = $null
            if ($bulkOk) {
                try { $v = $row.GetValues() } catch { $v = $null }
                if (-not ($v -is [array] -and $v.Count -eq 7)) { $v = $null }
            }
            if ($v) { $bulkRows++ }
            else {
                $v = @($row.Item($daslStart), $row.Item($daslEnd), $row.Item($daslRec),
                       $row.Item($daslAll),   $row.Item($daslKw), $row.Item('Subject'),
                       $row.Item('EntryID'))
            }
            [void]$batch.Add($v)
        }
        foreach ($v in $batch) {
            if ($v[2] -is [bool] -and $v[2]) { continue }   # recurring: expanded below
            if ($v[3] -is [bool] -and $v[3]) { continue }   # all-day: OM path below
            $st = [DateTime]::SpecifyKind([datetime]$v[0], 'Utc').ToLocalTime()
            $en = [DateTime]::SpecifyKind([datetime]$v[1], 'Utc').ToLocalTime()
            if (-not ($st -lt $rangeEnd -and $en -gt $rangeStart)) { continue }
            if ($n -ge 8000) { Write-Warning 'Stopped at 8000 items (safety cap).'; $trunc = $true; break }
            $kw = $v[4]
            $catStr = ''
            if ($kw -is [array]) { $catStr = ($kw -join '; ') }
            elseif ($kw -is [string]) { $catStr = $kw }
            $catStr = ((($catStr -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join '; ')
            $list.Add([pscustomobject]@{
                s   = $st.ToString('yyyy-MM-ddTHH:mm:ss', $inv)
                e   = $en.ToString('yyyy-MM-ddTHH:mm:ss', $inv)
                t   = [string]$v[5]
                ad  = $false
                cat = $catStr
                id  = [string]$v[6]
                r   = $false
            })
            $n++
        }
        if ($trunc) { break }
    }

    $tableMs = $tSplit.ElapsedMilliseconds
    $tableN  = $n

    # --- recurring series AND all-day items: the OM path. Expansion over
    # masters only; non-recurring all-day items pass through IncludeRecurrences
    # untouched (expanding a non-recurring item yields the item itself), getting
    # Outlook's own all-day normalisation for both storage conventions. ---
    $masters = $cal.Items.Restrict(('@SQL=("{0}" = 1 OR "{1}" = 1)' -f $daslRec, $daslAll))
    $masters.Sort('[Start]')
    $masters.IncludeRecurrences = $true
    $occ = $masters.Restrict($filter)
    foreach ($it in $occ) {
        Invoke-QuickServe   # keep page GETs responsive during this long read
        # v2.1.20: deliberately NO early exit on Quit. v2.1.19 broke out of this
        # loop when Quit arrived, which left the item list SHORT - and the build
        # went on to serialise that partial list, render it and SAVE IT TO THE
        # CACHE, so the next start opened a half-built year. A read either
        # completes or it is not worth having. Quit has already been answered
        # and the window already closed by the time we get here; finishing the
        # read costs a few invisible seconds and leaves a whole, correct cache
        # behind. The main loop exits on the flag as soon as the build returns.
        if ($n -ge 8000) { Write-Warning 'Stopped at 8000 items (safety cap).'; $trunc = $true; break }
        try {
            # v3.2.1: ONE property access, not two. Every one of these is a
            # cross-process call, and testing the property and then reading it
            # paid for the round trip twice on every item in this loop.
            $catRaw = $it.Categories
            $catStr = ''
            if ($catRaw) { $catStr = [string]$catRaw }
            $catStr = ((($catStr -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join '; ')
            $list.Add([pscustomobject]@{
                s   = $it.Start.ToString('yyyy-MM-ddTHH:mm:ss', $inv)
                e   = $it.End.ToString('yyyy-MM-ddTHH:mm:ss', $inv)
                t   = [string]$it.Subject
                ad  = [bool]$it.AllDayEvent
                cat = $catStr
                id  = [string]$it.EntryID
                r   = [bool]$it.IsRecurring
            })
            $n++
        } catch { Write-Verbose ("Skipped one item: {0}" -f $_.Exception.Message) }
    }
    @{
        Events   = $list
        Truncated = $trunc
        TableN   = $tableN
        TableMs  = $tableMs
        BulkN    = $bulkRows
        BatchN   = $batchCalls
        OmN      = ($n - $tableN)
        OmMs     = ($tSplit.ElapsedMilliseconds - $tableMs)
    }
}

function Compare-YearReads([object]$a, [object]$b) {
    # Order-independent comparison of the six CONTENT fields. EntryID is
    # deliberately NOT compared: the Table API can hand out short-term Exchange
    # EntryIDs while the object model reports long-term ones (the 1.2.8 field
    # log showed structurally different id families on every item, with all
    # content identical) - both forms resolve identically through
    # GetItemFromID, so ids are validated FUNCTIONALLY by the caller instead.
    # Returns the differing items (capped each way) so a mismatch verdict names
    # WHAT diverged, not just that something did.
    $us = [string][char]31
    $key = { param($ev) @($ev.s, $ev.e, $ev.t, $ev.ad, $ev.cat, $ev.r) -join $us }
    $ha = New-Object 'System.Collections.Generic.HashSet[string]'
    $hb = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($ev in $a.Events) { [void]$ha.Add((& $key $ev)) }
    foreach ($ev in $b.Events) { [void]$hb.Add((& $key $ev)) }
    $onlyA = New-Object System.Collections.Generic.List[object]
    $onlyB = New-Object System.Collections.Generic.List[object]
    foreach ($ev in $a.Events) { if (-not $hb.Contains((& $key $ev)) -and $onlyA.Count -lt 8) { $onlyA.Add($ev) } }
    foreach ($ev in $b.Events) { if (-not $ha.Contains((& $key $ev)) -and $onlyB.Count -lt 8) { $onlyB.Add($ev) } }
    @{
        Equal = ($a.Events.Count -eq $b.Events.Count -and $onlyA.Count -eq 0 -and $onlyB.Count -eq 0)
        OnlyA = $onlyA
        OnlyB = $onlyB
    }
}

function Get-YearData {
    # v3.0.1: the read-and-serialise half of the former Build-YearViewHtml,
    # lifted out UNCHANGED so the same payload can be produced without a page
    # wrapped around it. Returns the JSON string and the content hash; the
    # CALLER decides whether that hash becomes the change-detection baseline,
    # because a payload fetched for a neighbouring year must not displace the
    # baseline belonging to the year on screen.
    param([int]$Y)
    # Per-phase timing, surfaced in the page stamp so slow builds can be
    # diagnosed with data instead of guesses (added after the fast-JSON change
    # produced no visible speed-up - the cost lives elsewhere).
    $sw = [Diagnostics.Stopwatch]::StartNew()
    # Master category list: name -> colour
    $catMap = @{}
    $catOrder = New-Object System.Collections.Generic.List[string]
    foreach ($c in $ns.Categories) {
        $catMap[$c.Name] = $catColors[[int]$c.Color]
        $catOrder.Add($c.Name)
    }

    $perfCatMs = $sw.ElapsedMilliseconds

    $cal = Get-SelectedCalendar
    $script:calStoreId = [string]$cal.StoreID
    Write-Host ("Reading calendar '{0}' for {1}..." -f $cal.Name, $Y) -ForegroundColor Cyan

    # R-1: the filter string must not depend on the machine's regional settings.
    # Both endpoints are 1 January by construction, so a numeric literal reads the
    # same under dd/MM and MM/dd; date-only (midnight implied) avoids the AM/PM
    # designator, which other locales spell differently or omit; InvariantCulture
    # stops the string following the thread culture (not always the user locale
    # Outlook parses with). Do not widen this range without revisiting all three.
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $rangeStart = [datetime]::new($Y, 1, 1)
    $rangeEnd   = $rangeStart.AddYears(1)
    $filter = "[Start] < '{0}' AND [End] > '{1}'" -f `
        $rangeEnd.ToString('MM/dd/yyyy', $inv), $rangeStart.ToString('MM/dd/yyyy', $inv)
    $perfSetupMs = $sw.ElapsedMilliseconds - $perfCatMs

    # Reader orchestration: fast reader when trusted, classic otherwise; on the
    # first build of the session, run both and require identical results.
    $read = $null; $perfEnumMs = 0
    $perfTableN = 0; $perfTableMs = 0; $perfOmN = 0; $perfOmMs = 0; $perfBulkN = 0; $perfBatchN = 0
    if ($script:fastScanOk) {
        $tF = [Diagnostics.Stopwatch]::StartNew()
        try { $read = Read-YearFast $cal $filter $rangeStart $rangeEnd }
        catch {
            Write-Warning ("Fast scan failed ({0}) - using the classic reader for this session." -f $_.Exception.Message)
            $script:fastScanOk = $false
            $read = $null
            try { Remove-Item (Join-Path (Split-Path -Parent $PSCommandPath) 'yv-fastscan.ok') -Force -ErrorAction SilentlyContinue } catch {}
        }
        if ($read) { $perfEnumMs = $tF.ElapsedMilliseconds }
    }
    if ($read -and $read.ContainsKey('TableN')) {
        $perfTableN  = [int]$read.TableN
        $perfTableMs = [int]$read.TableMs
        $perfOmN     = [int]$read.OmN
        $perfOmMs    = [int]$read.OmMs
        if ($read.ContainsKey('BulkN'))  { $perfBulkN  = [int]$read.BulkN }
        if ($read.ContainsKey('BatchN')) { $perfBatchN = [int]$read.BatchN }
    }
    $perfValidateMs = 0
    if ($read -and -not $script:fastScanChecked) {
        $script:fastScanChecked = $true
        $tC = [Diagnostics.Stopwatch]::StartNew()
        $classic = Read-YearClassic $cal $filter
        $classicMs = $tC.ElapsedMilliseconds
        $perfValidateMs = $classicMs   # first build pays for both readers, ONCE
        $cmp = Compare-YearReads $read $classic
        if ($cmp.Equal) {
            # Content matches; now prove the fast reader's ids WORK - resolve a
            # sample through GetItemFromID (the exact call /open and /delete
            # make). The first events in the list come from the Table path,
            # which is precisely where the suspect ids originate.
            $idOk = $true
            try {
                foreach ($sv in @($read.Events | Select-Object -First 3)) {
                    [void]$script:ns.GetItemFromID($sv.id, $script:calStoreId)
                }
            } catch { $idOk = $false }
            if (-not $idOk) {
                Write-Warning 'Fast scan ids failed to resolve via GetItemFromID - using the classic reader for this session.'
                try { Remove-Item (Join-Path (Split-Path -Parent $PSCommandPath) 'yv-fastscan.ok') -Force -ErrorAction SilentlyContinue } catch {}
                $script:fastScanOk = $false
                $read = $classic
                $perfEnumMs = $classicMs
            } else {
                Write-Host ("Fast scan validated against classic reader: {0} ms vs {1} ms, {2} items identical (ids resolve)." -f `
                    $perfEnumMs, $classicMs, $read.Events.Count) -ForegroundColor Green
                try {
                    [IO.File]::WriteAllText((Join-Path (Split-Path -Parent $PSCommandPath) 'yv-fastscan.ok'),
                        ('{0}|{1}' -f $script:yvVersion, [string]$script:outlook.Version))
                } catch {}
            }
        } else {
            Write-Warning ("Fast scan MISMATCH vs classic reader ({0} vs {1} items) - using the classic reader for this session." -f `
                $read.Events.Count, $classic.Events.Count)
            # Name what diverged (first 5 each way) - to the console AND the log
            # file, so a hidden bridge's verdict is still diagnosable afterwards.
            $diffLines = New-Object System.Collections.Generic.List[string]
            # v3.4.9: dates, flags and an id suffix only - NO titles or
            # categories. This goes to yv-error.log in plain text (the page
            # cache is encrypted; the log is not), and logs get attached to
            # bug reports. The dates and id are enough to find the item.
            foreach ($ev in $cmp.OnlyA) { $diffLines.Add(("  fast-only   : {0} -> {1}  (ad={2} r={3} id..{4})" -f $ev.s, $ev.e, $ev.ad, $ev.r, $ev.id.Substring([Math]::Max(0, $ev.id.Length - 8)))) }
            foreach ($ev in $cmp.OnlyB) { $diffLines.Add(("  classic-only: {0} -> {1}  (ad={2} r={3} id..{4})" -f $ev.s, $ev.e, $ev.ad, $ev.r, $ev.id.Substring([Math]::Max(0, $ev.id.Length - 8)))) }
            foreach ($dl in $diffLines) { Write-Warning $dl }
            try { Remove-Item (Join-Path (Split-Path -Parent $PSCommandPath) 'yv-fastscan.ok') -Force -ErrorAction SilentlyContinue } catch {}
            try {
                $logP = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
                Add-Content -Path $logP -Value (@("[{0}] Fast-scan mismatch ({1} fast vs {2} classic items):" -f `
                    (Get-Date).ToString('s'), $read.Events.Count, $classic.Events.Count) + $diffLines)
            } catch {}
            $script:fastScanOk = $false
            $read = $classic
            $perfEnumMs = $classicMs
        }
    }
    if (-not $read) {
        $tC2 = [Diagnostics.Stopwatch]::StartNew()
        $read = Read-YearClassic $cal $filter
        $perfEnumMs = $tC2.ElapsedMilliseconds
    }

    # Post-pass (pure PowerShell, no COM): colour mapping and legend usage.
    # Uncategorised events are ALWAYS included in the data (since v1.20); their
    # visibility is governed by the client's "(No category)" filter entry.
    $truncated = [bool]$read.Truncated
    $events = New-Object System.Collections.Generic.List[object]
    $usedCats = @{}
    foreach ($rv in $read.Events) {
        $first = ($rv.cat -split '[;,]')[0].Trim()
        $col = '#BFBFBF'
        if ($first -and $catMap.ContainsKey($first)) { $col = $catMap[$first] }
        if ($first) { $usedCats[$first] = $col }
        $events.Add([pscustomobject]@{
            s = $rv.s; e = $rv.e; t = $rv.t; ad = $rv.ad
            cat = $rv.cat; col = $col; id = $rv.id; r = $rv.r
        })
    }
    # Order-independent CONTENT hash - generated timestamps and perf numbers
    # deliberately excluded (a sync-echo rebuild hashes identically), but the
    # mapped COLOUR is included (v1.2.25): with stale-revalidate, a category
    # recolour must register as a change or revalidation would wrongly keep a
    # stale page.
    $keys = New-Object System.Collections.Generic.List[string]
    $us = [string][char]31
    foreach ($ev in $events) { $keys.Add((@($ev.id, $ev.s, $ev.e, $ev.t, $ev.ad, $ev.cat, $ev.col, $ev.r) -join $us)) }
    $keys.Sort([StringComparer]::Ordinal)
    $sha = [Security.Cryptography.SHA1]::Create()
    $contentHash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($keys -join "`n"))))
    $sha.Dispose()
    Write-Host ("{0} appointments found ({1} ms categories, {2} ms restrict, {3} ms enumerate, {4} ms baseline sort)." -f `
        $events.Count, $perfCatMs, $perfSetupMs, $perfEnumMs, $script:lastBaselineMs) -ForegroundColor Cyan

    # Legend: master-list order, used categories only
    $legend = New-Object System.Collections.Generic.List[object]
    foreach ($name in $catOrder) {
        if ($usedCats.ContainsKey($name)) {
            $legend.Add([pscustomobject]@{ name = $name; col = $usedCats[$name] })
            $usedCats.Remove($name)
        }
    }
    foreach ($k in @($usedCats.Keys)) {
        $legend.Add([pscustomobject]@{ name = $k; col = $usedCats[$k] })
    }

    # Fast JSON (v1.2.3): ConvertTo-Json in Windows PowerShell 5.1 costs SECONDS
    # on a few thousand objects and dominated the build time. The payload's shape
    # is fixed and fully known, so it is serialised directly (StringBuilder +
    # regex-escaped strings). Safety net: the first build each session round-trips
    # the result through ConvertFrom-Json; if that ever fails, the session falls
    # back to ConvertTo-Json permanently. Output is equivalent either way, and the
    # </ neutralisation below applies to both.
    $json = $null
    if ($script:fastJsonOk) {
        $sb = New-Object Text.StringBuilder (256 + 160 * $events.Count)
        [void]$sb.Append('{"year":').Append([string]$Y)
        [void]$sb.Append(',"generated":"').Append((Get-Date).ToString('o')).Append('"')
        [void]$sb.Append(',"source":').Append((ConvertTo-JsonStr ([string]$cal.FolderPath)))
        [void]$sb.Append(',"events":[')
        $sep = ''
        foreach ($ev in $events) {
            [void]$sb.Append($sep); $sep = ','
            # dates are our fixed invariant format and col is palette hex - no
            # escapable characters by construction; everything else is escaped
            [void]$sb.Append('{"s":"').Append($ev.s).Append('","e":"').Append($ev.e)
            [void]$sb.Append('","t":').Append((ConvertTo-JsonStr $ev.t))
            [void]$sb.Append(',"ad":').Append($(if ($ev.ad) {'true'} else {'false'}))
            [void]$sb.Append(',"cat":').Append((ConvertTo-JsonStr $ev.cat))
            [void]$sb.Append(',"col":"').Append($ev.col)
            [void]$sb.Append('","id":').Append((ConvertTo-JsonStr $ev.id))
            [void]$sb.Append(',"r":').Append($(if ($ev.r) {'true'} else {'false'})).Append('}')
        }
        [void]$sb.Append('],"categories":[')
        $sep = ''
        foreach ($c in $legend) {
            [void]$sb.Append($sep); $sep = ','
            [void]$sb.Append('{"name":').Append((ConvertTo-JsonStr $c.name))
            [void]$sb.Append(',"col":"').Append($c.col).Append('"}')
        }
        [void]$sb.Append('],"perf":{"reader":"').Append($(if ($script:fastScanOk) {'fast'} else {'classic'})).Append('"')
        [void]$sb.Append(',"catMs":').Append([string]$perfCatMs)
        [void]$sb.Append(',"setupMs":').Append([string]$perfSetupMs)
        [void]$sb.Append(',"enumMs":').Append([string]$perfEnumMs)
        # v3.2.1: zero for the classic reader, which has no such split.
        [void]$sb.Append(',"tableN":').Append([string]$perfTableN)
        [void]$sb.Append(',"tableMs":').Append([string]$perfTableMs)
        [void]$sb.Append(',"omN":').Append([string]$perfOmN)
        [void]$sb.Append(',"omMs":').Append([string]$perfOmMs)
        [void]$sb.Append(',"bulkN":').Append([string]$perfBulkN)
        [void]$sb.Append(',"batchN":').Append([string]$perfBatchN)
        [void]$sb.Append(',"baselineMs":').Append([string]$script:lastBaselineMs)
        [void]$sb.Append(',"validateMs":').Append([string]$perfValidateMs).Append('}')
        [void]$sb.Append(',"truncated":').Append($(if ($truncated) {'true'} else {'false'}))
        [void]$sb.Append(',"uncatDefaultHidden":').Append($(if (-not $IncludeUncategorised) {'true'} else {'false'}))
        [void]$sb.Append('}')
        $json = $sb.ToString()
        if (-not $script:fastJsonChecked) {
            $script:fastJsonChecked = $true
            try { [void]($json | ConvertFrom-Json) }
            catch {
                Write-Warning 'Fast JSON failed validation - using ConvertTo-Json for this session.'
                $script:fastJsonOk = $false
                $json = $null
            }
        }
    }
    if (-not $json) {
        $payload = [pscustomobject]@{
            year       = $Y
            generated  = (Get-Date).ToString('o')
            source     = [string]$cal.FolderPath
            events     = $events
            categories = $legend
            perf       = [pscustomobject]@{ reader = $(if ($script:fastScanOk) {'fast'} else {'classic'})
                                            catMs = [long]$perfCatMs; setupMs = [long]$perfSetupMs
                                            enumMs = [long]$perfEnumMs; baselineMs = [long]$script:lastBaselineMs
                                            validateMs = [long]$perfValidateMs }
            truncated  = [bool]$truncated   # R-4: shown in the page header; the console warning is invisible when hidden
            uncatDefaultHidden = [bool](-not $IncludeUncategorised)   # first-run filter default for "(No category)"
        }
        $json = ($payload | ConvertTo-Json -Depth 6 -Compress)
    }
    $json = $json -replace '</', '<\/'
    # R-17: release the COM wrappers from this enumeration promptly - a long-lived
    # bridge doing many rebuilds otherwise holds Outlook-side objects until the
    # garbage collector happens to run.
    $read = $null; $cal = $null   # reader-held COM wrappers released with their function scopes
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    return [pscustomobject]@{ Json = $json; Hash = $contentHash }
}

function Build-YearViewHtml {
    param([int]$Y)
    # v3.0.1: the PAGE half. Get-YearData now performs the read and the
    # serialisation; everything below is the page wrapper - the literals baked
    # into the template - and is unchanged from v3.0.0. The page build is the
    # one that sets the change-detection baseline, exactly as before.
    $yd = Get-YearData -Y $Y
    $json = $yd.Json
    $script:lastBuildHash = $yd.Hash
    # v3.4.1: a page build has just produced this year's payload, so keep it.
    # Visiting a year directly therefore also warms it for a later spill fetch.
    $script:payloads[$Y] = @{ Json = $json; Hash = $yd.Hash; Ticks = (Get-Date).Ticks }

    # R-10: static snapshots really do get "0"/"" now - there is no bridge for the
    # page to poll or authenticate against, and the token must not be written to disk.
    $ticks = '0'; $tok = ''
    if ($Serve) { $ticks = [string](Get-Date).Ticks; $tok = $script:sessionToken }
    # v3.4.9: ConvertTo-SeedStr, not ConvertTo-JsonStr. This literal sits in
    # the page's SCRIPT block and carries category names; a name containing
    # '</script>' ended the block and ran whatever followed it. '<' and ';'
    # are escaped (as JSON escapes, so the value is unchanged).
    $hideLit = ConvertTo-SeedStr $(if ($null -ne $script:hiddenCsv) { $script:hiddenCsv } else { '' })
    $hideSet = $(if ($null -ne $script:hiddenCsv) { 'true' } else { 'false' })
    $script:pageHideLits[$Y] = $hideLit   # v2.1.22: so /setfilter can re-stamp this page in place
    $confLit = $(if ($script:confirmDelete) { 'true' } else { 'false' })
    $scrlLit = $(if ($script:scrollYears)   { 'true' } else { 'false' })
    $ntLit   = ConvertTo-SeedStr ([string]$script:noTitles)
    $prtLit  = [string][int]$script:printPages
    $plgLit  = $(if ($script:printLegend)   { 'true' } else { 'false' })
    $smLit   = [string][int]$script:startMonth
    # v3.4.2: when the view starts after January it needs the FOLLOWING year,
    # and fetching that after the page had rendered meant the spill months
    # arrived a beat late even when the answer was cached - the round trip and
    # the second render were the cost, not the read. Carrying it IN THE PAGE
    # makes those months part of the first paint, and the page cache then
    # serves both years together exactly as it serves one.
    #
    # The same rule the client applies: only the current calendar year is
    # offset. If this decides wrongly the client simply fetches as before, so
    # a disagreement costs a round trip rather than a wrong page.
    $data2Lit = 'null'
    $script:lastSpillHash = ''
    $effStart = 0
    if ($Y -eq (Get-Date).Year) {
        $effStart = $(if ($script:startMonth -eq -1) { (Get-Date).Month - 1 } else { [int]$script:startMonth })
    }
    if ($effStart -gt 0 -and $effStart -le 11) {
        try {
            $ny = $Y + 1
            $nc = $script:payloads[$ny]
            if ($nc -and [long]$nc.Ticks -ge [long]$script:lastChangeTicks) {
                $data2Lit = [string]$nc.Json
                $script:lastSpillHash = [string]$nc.Hash
            } else {
                # v3.4.5: the hash is KEPT now, but only to answer 'did the
                # spill months change'. It still must NOT become
                # $script:lastBuildHash - that belongs to the year on screen,
                # and a hash from another year there would make the next poll
                # report a change that never happened (v3.0.1).
                $yd2 = Get-YearData -Y $ny
                $data2Lit = $yd2.Json
                $script:lastSpillHash = [string]$yd2.Hash
                $script:payloads[$ny] = @{ Json = $yd2.Json; Hash = $yd2.Hash; Ticks = (Get-Date).Ticks }
            }
        } catch {
            $data2Lit = 'null'   # the client will fetch it instead
            Write-Warning ('Could not pre-read {0} for the spill months: {1}' -f ($Y + 1), $_.Exception.Message)
        }
    }
    # v3.4.9: ONE PASS. This was a chain of .Replace calls, and each one
    # re-scanned everything the earlier ones had inserted - calendar data
    # included. An appointment titled "__HIDE__" therefore had a quoted JSON
    # string spliced into the middle of the data, and the page came up blank;
    # anyone able to send a meeting invitation could do that. A single regex
    # pass inserts every value exactly once and never looks at it again.
    $subst = @{
        DATA = $json; DATA2 = $data2Lit; YEAR = [string]$Y; TICKS = $ticks; TOKEN = $tok
        HIDE = $hideLit; HIDESET = $hideSet; CONFIRMDEL = $confLit; SCROLLYEARS = $scrlLit
        PRINTPAGES = $prtLit; PRINTLEGEND = $plgLit; STARTMONTH = $smLit; NOTITLES = $ntLit
    }
    $fill = {
        param($m)
        $v = $subst[$m.Groups[1].Value]
        if ($null -ne $v) { [string]$v } else { $m.Value }
    }.GetNewClosure()
    return $script:placeholderRx.Replace($template, [Text.RegularExpressions.MatchEvaluator]$fill)
}

if (-not $Serve) {
    $html = Build-YearViewHtml -Y $Year
    [IO.File]::WriteAllText($OutFile, $html, (New-Object Text.UTF8Encoding($false)))
    Write-Host ("Written: {0}" -f $OutFile) -ForegroundColor Green
    # F-3: if we started Outlook ourselves and the user hasn't begun using it
    # (no Explorer or Inspector windows have appeared), close it again.
    if ($script:outlookStarted) {
        try {
            if ($script:outlook.Explorers.Count -eq 0 -and $script:outlook.Inspectors.Count -eq 0) {
                Write-Host 'Closing the Outlook instance this snapshot started.' -ForegroundColor DarkCyan
                $script:outlook.Quit()
            }
        } catch {}
    }
    if (-not $NoOpen) { Start-Process $OutFile }
    return
}

# Splash page: served instantly on '/' and on any page request whose year is not
# yet cached. It holds the session token and POSTs /build (the only route that may
# start a calendar read), then navigates to the finished page. A cross-origin
# caller can force this splash to be served but cannot read it, so the token - and
# with it the ability to trigger a build - never leaves the same-origin boundary.
$splashHtml = @'
<!DOCTYPE html><html><head><meta charset="utf-8"><title>YearView loading</title><link rel="icon" type="image/x-icon" href="data:image/x-icon;base64,AAABAAQAEBAAAAAAIAAsAgAARgAAABgYAAAAACAAvAIAAHICAAAgIAAAAAAgAFoDAAAuBQAAMDAAAAAAIACPCAAAiAgAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAAfNJREFUeJyNkrFrU1EUh79z730vTWKSxmKDcbC2XcQqtmDp0EmxiwhOjoJTBwcnd3HQ/0HwD9DFuVARRURxqgVBwVC0UNRWzEti2rx373FoKQ7t09/+/c7HOUcAAfRp7eRwo1I2CfkpG6v14GVq/cNP9mCeNydvlsU86KuiYASVwyvEF8WYbfWPfZTekWeNiTNx5N7HYkyGkkPuR4GiGNpZetlZwrj53TeJTwO5k/9yQDQYi4J1duHqysSNa+1quV6zYhXJlxCETD2D0Jfv336o27h7uz03Pj74pB/Z6K0T2wKqejAsQhpSjsWjnI3maG9t4forr5xpjsny+hNeb7ygElXx6g8ssGLpZl2mhy8wdfoeO0mCa5QamolnoXmF8/VpIhvnGmQh5Wg0iliLswUcVAl4ar0Z6Mxgbf4CQ4ByCaiABVwCOCO8ayWsft6kWHAcIoAIbA+UyeYQ1xsNPOCqgA/KuVMVTozERNZwCI8AWYBaafdQFnCQENQxduQtY3wBE0NehaYQHwfmgQy35xa0+5Xu5irGDuUWaNjBlbcZqs0HES/OWitoVpX6RS1V5vQff4SiWFdUwKjG3rVa9WR2dvCoWB25ZaOR/3plQDq93nKn8+vNPrC2tnZJIrE+213OQfFAwTlN0zTcX1p6+XBxMf0DX33AxmY0R+wAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAGAAAABgIBgAAAOB3PfgAAAKDSURBVHictZW9a5RBEIef2d37SO5yyXkHGmMRNTFySakIUSSFkF5Io4Xg32BvZSeW9oIgGNBaCzUIKWysDGpIIRH8AKM575Lz/dix2JzJ2d17urDFLru/Z2ZnZgeAm2D4T0MUjIB/cWRqqWjs9QhvFCSLmIIWEYm8rlz8vH5LQQTg5cTUUlncQ4+iA1qsQFkMX31y5+LH9zcEYGV8+lnJmoWWT2NB7KAMZ0TiOGm1vnw77gDESIwxYsASZnZ1ESVJxR4+HE3euysOIGm3xRtD4v2AxoOKwG4Hc2JSKouLOIDTy/cpDw8Rqwf13dBkUFdAsAi/nGPjzWYAjF5aYHhg23tHBDSXnwRA0myipRIbrXest9bImwKq/eWTiBD5X0yXG5wszxB3OlCpBIAXEGtZb7/l0acHjLhRvKZ9AYxYfibbXD56lanRRqguCICDh3KSx0keT/+AnOQxf2W5A7AmbJ6rn2emOoPNmKkpKWO2vgfs8SAs8kmdkU4dI/Rd0QJ4hXwRyO1rOoA0DXJrm01W15oM5y2+T4IR2IlS5hsVzpyq0C2pnhiogveC9/QNQMLdv5PPBXpw51i9yIU5h7PZnihJYaIWbO7W6t5fFBbjNcN4ze0dz4JQuq2lB4APKRl9f02ytYrYofBl9KVv0HQXd2iefPXsn/s9MRDfgXgLKGUCELeDxoERnkjEA0iuhh05jdhiRg86SK7W3dF9D8QMAamrNNRVGv+iqXnEFKy1IbxxFD8lNJpuhAeZBrBJ4led29kWVTUioj9+NG8XioUroAZFFO27KQjixYhEUfLq4+aHa7Ozs996RB4/fz42VCgYmhkeBiiVUrXVqszPzW0BqKr8BrwV5/mNiVTUAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAADIUlEQVR4nL2Xz2sdVRTHP+fc+ZGXJrFBzas/0iJdFKKoRYSQhWThQrAiFSouXLv2H5BQcCFaXIsUdCUSF+JOkO5FlCpoqK5KxEBa25eEPGfezJ17XLz3yoP3bILc1y9cZrhzOd/PvffMmTswkG1sKA9Im5cuueG9ABiIgAHy0/K5F6rAXIkniWTogYSENIRm26fX37y9dTj0FNvYUC5ftp8/+Hil/vTzq1YWq6qK2YAugmxwFaCB7bC+9u7aF599bYaKgX61spK0C/1+0Yfzd33VIBLLe4ylJeqaLC3rplp5afuPmyoQlg6K52bK3vm7vvIi4gR0Ss39Y00926tmpKovClgCkJrMq6phYeqJqCA4Z2TJPNDPMxFtEBUxTET66TEteY/kubQuvtrw0VYfoCkKcaGHtwB1jTVhChAGCJIkhADB+gYJwKmrn9BeepQGo7v1O3Wng7gEzO4b8tgSwAxJEuaef5Ysy9ifbcGV9/sAM+vrzJ+cA+Dk2moc0yNUd7vAYAXY28PmWwhGP02mlIsWCBYQEUJZjgAkCeL61XGn+JNv/vqSVNOo3nWoef2Jt3i8tdzvUB0BGFHXH/LbwXVybWHEyQFB6IWCl9sXxp6NAagomeZkmkUFMAI6YWsnAsy6E2Q6ExXAiTsewJkTZ3nv6Q+J9ykaykh15mgAMcWF2cjmg9jC2LzGAHY7Fd/+2CF1Eq8OCdSN8cqLi5xazO4PYAa92jCLWAgFKm8T400E8N6G1TMagD8uQJ4JZ9opSeQt8I2Rp+OJPQbw8ELKhdVH4jgfQ+NbUO/jD34FcZPG/39ZQ7LwDJI+dBTAHv7va6A5RCpEIBB6uNaTRwMgDknmQLPIAOnEVZ1w9A/QlGAhMkDVj/0fADJ0k7xNtvw2oHEBCEjeBsBGXq9k0NEMR4nmSOt0JONxhRBQVWMwOzUzuXVr55eiKHcBB1QQ/HQataoqIAfd7nfD5XAAOzu7b1RVvW9TVhOC3b5z5x2Azc1NJwMIERG7cePmU0tLi68lWbpQlqVorKNhAFRR8J3D4trZ04/9MPS8N8bMHtjv+ajXv50I0iaYHOK7AAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAIVklEQVR4nM2aTYxkVRXHf+fe+159dE/3DPQMQ+MAMzbNhwiLTjQhCIiABFEUQy/cuJPExARMXJk4gyuXbIzRvZtBEyWCATaAEDFhQgIBw4D2zDgfOsMw1V3d9fHeu/e4eNU1VdWvuqvaSM8/qfSrV/fdd/7n3PN1bwtDoGAAeW7YgM8ITwDPAYsQAN3ygcNgFOT/Ldh2cBTs4L0+QQ+DeSZnyrH9N92q1uxfJUOwO0JIES0DRnT1hdPHjz0DQcFIR0boIaAgAvrK7MH5CdyzIA86EafsnDnW3+1RUN5rqv/J/ef+8VIvCYFc80dA3zwwd8gG87dYzNX14KGH6U5CQSoiEmFYIXviq2c+/t1RsIvgpTPAAPrq7E0vTYt9cEWzRJB4h+UegGYOYwN8UlFuWzh3/CKAOdwxx5uzN8/HyP119eHKEx5AXErwk8bsXYNvC+ir3GvNfbn2UeGOkojVUULVDkHyjwI3AuxiVcwuFgTAh/BFJ4KgVyyBDkSFdP2L6V6I+J2RZ3yIajcwur5fjEHyBJwHVblC8plq/kHBGHoXSZdASFMXWi3SkCEImnk0y3aehCpiDRJFKEowjtBLoH74UeWZY1Sfevqdufu+RAAJQWmfOUPzxClMFAF6OauM9fK+P13IhothENRnRHv2MHHrzSDgEE79+wJ895vw6x+AqlqAd0+dv1dzeL1ykaqqLv3n4hGAt99+O+ouIb1wwYXZqzAiLJ2rsdpIADh47TTVUjS61mVdsYIPgY9PXyLzobsSVcFZw9zn9mCN2WSirmSdZSwIigJhdbX762UC1qrYvNirt1I+XU0QgQMIJnKDs27+ys4aVeBSIyXN+glEzqDGgB2FQA/E5MrpIV4o2Z6JmNiAGMGnKaeXLtGj2k0Fd86yZ+80QZUQlL1TMV61+6gCVoR2OyE1ozmVqmKtpVzaWCAUEqi3PBdX2kxP7Ob5vz7PG+0/MFmZIoRNUoUKxgn1T9f48cLPuH3+C2Q+UGssF1rg0IEKI8q/KQoJCAFLQDVhurKbmbVZyskEqpsUpwrihQoJzkS02m2yLCDqMfRbQFRpNlvYUdsMBWMMcbzRFwsJKALGstZu8tCXv87j5hsEHSElKIgFjSDLMnCSO11Pi6ediGysKXZiyd11g0RDXl5MQCFoPk0WMrS6oZPbhIMiGKLIEQE2KhEISEdWDWAjQ6VgPW8HhQScUSITEBHSJKHZzEltbQFFjGG1vkyz2USB2oUa3veaAKwVTkZrA0tICMEzNbWbuFRG9bLWx3bixEMrE0SgVC5TrYyureXaJX75278gpRlQjw/FfvP6Oxf7FCIC7USZnW7xw+8/TEH/XohCArvKDtE8cmRpQqMZtraAdpIMhjvm99FOFDERIRRX50b6w7IAaeaZ3XsNjUYbRbZvgXYWWGt7jIC1jnIpLiRgCpxwolrh8Ue+0v3+wYlPCsPobTfObKKN0VHsxGLBxCjg4hJmSMDOvC/o37Sr9cwH6mudcNpbSjhDo9nGDWRiJVeSkUHF5P5g7cZlVUigduE0Z8+exTkH9QnKce8wQdVTKle54YY5xJihCVpRIisIMlAL5VWNFjR/quDxA/OAMVsQyDojM4XzH/6JAxOnICvjz3nWtHcyIXKBf54Xrp75OXtndvdFjF5YH8iCkHr6CCBCKY43WGA7cH0XAk7g4J3f4fz5czjnqJYjXO9aF9AQmL+xSqUSs9ZoDrVAUGXPpCME2xdGjRFarVbuyCNAYTwnjqeuZZfdB6LcOXc18TYVFRQunayRphud+OBsaSwLDNtpKCQQ0hZZu4mIUKtFTFaiDR2ZIMjQaiy/H4JS3VUlHewHnMG60bP75RlHJKBi8qJGBOcccVQcRocj11fLJ7z1+z/SulTHOJvf7fTlf5+qDK1v+gQX8JmnetUUD3/vMSIX9e13blkLWWtxzo5FYN2ps2aDj15+DXOxAZHt2004OeLukxhBkwxmJrnn8YeYnoygJ7tvWQslSZuG0TEtAKjiTMS3fvFTGj0Oq4AzhkPX7R6aX/oIACEEJqpVpiengP4EunUtVBqvFupFGgJRPMmEndjgxNfv37etOQcxVi1UiWOM176lsBlsUPbFlixoXz8QOUPSzkh9uunz61BVnLWUy6XRCCReaSaKCERRTKUSIQgvvPEav3rrz0xVJ4dWmYNIs/5x1hhWGqv86O5HePTu+/DBY814EWlLAmnmaacZIpBmGVlmEITMZ2SZJ/Me70choCRp/9mcMZYk7cyTeXzw6JbpYMxayAo40zlzCoHg8yX02D1f47G77h95CRGUD5Y+yfPAuigKcWSZn5sh8xkA3m++rzxSLdTZ+crDX0/94qKIUqecVhSNpKBnLUbqA3U8SRjIxEEJKKVttpXaUwV2CVhr3Xqom646nMSICD5LaTTCtrZGgypXTXSamoFaqN1ukybj10KqijHS3Z5wnd9ZXl4722i20mql7JabGbV67gOz15SoVsbYWuwjALVTK4UNzecPVLezLyQiQpb4JYCFhQV1IhJUVZ588jfHbzi4/+PqdftviQxZyakTgaTdpmVGaCkL4L0SWcUw2NSPvy8EhHK5RG253jp+fOmF9Vfkp5SqVkT8u+9/+MD83KFXSrEjhNzD8mi5/VOnYdl2WK9cDME56wA+/OjE07fMH3x2XebuENU8mL37/keLteWVM5/djvloWKmvriyd/NdTvbLm1HqgqkZEwosvvr731ttvejRy9vpGq4HI//KvBsPyxda9gKrXSqki7SQ9eeLcqZcfuOuuM6oqIjLcfEc7Bx5XIo4ePbpBtg2JbFHEa34KaI8dO3ZFnPItLCzokSNHwuLi4oaM91/n+BV7KaS68AAAAABJRU5ErkJggg=="><style>
html,body{height:100%;margin:0;font:14px "Segoe UI",system-ui,sans-serif;color:#1a1d21;
background:#fff;display:flex;align-items:center;justify-content:center}
.box{text-align:center}
.sp{width:34px;height:34px;border:4px solid #e9ebee;border-top-color:#c11a25;border-radius:50%;
margin:0 auto 14px;animation:r 0.9s linear infinite}
@keyframes r{to{transform:rotate(360deg)}}
</style></head><body><div class="box"><div class="sp"></div>
<div id="m">Building YearView page - please wait...</div></div>
<script>
(function attempt(n){
fetch("/build?year=__Y__", {method:"POST", headers:{"X-YV-Token":"__TOKEN__"}})
  .then(function(r){
    if(r.status===503){ throw 503; }         // bridge busy: retry until it frees
    if(!r.ok) throw 0;
    return r.json();
  })
  .then(function(b){
    // v2.0.28: a 200 can be the retryable "Starting Outlook" fault. The old
    // splash navigated on ANY 200 - for an uncached year that served this
    // splash again, which built again, which answered retry again: a reload
    // loop for the whole Outlook boot. Wait here instead, honestly labelled.
    if(b && b.retry){
      if(n<150){
        var m=document.getElementById("m");
        if(m){ m.textContent=(b.err||"Starting Outlook")+"\u2026"; }
        throw {wait:2000};
      }
      throw 0;
    }
    location.replace("/?year=__Y__");
  })
  .catch(function(e){
    if(e && e.wait){ setTimeout(function(){ attempt(n+1); }, e.wait); return; }
    if(e===503 && n<120){ setTimeout(function(){ attempt(n+1); }, 1500); return; }
    if(e!==503 && n<3){ setTimeout(function(){ attempt(n+1); }, 1500); return; }
    document.body.textContent =
      "YearView is not responding - restart it from the YearView shortcut.";
  });
})(0);
</script></body></html>
'@
function Get-SplashBytes([int]$Y) {
    ,[Text.Encoding]::UTF8.GetBytes(
        $splashHtml.Replace('__Y__', [string]$Y).Replace('__TOKEN__', $script:sessionToken))
}
$splashBytes = Get-SplashBytes $Year

# ---------------- interactive bridge mode ----------------
if (-not ('YVWin' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class YVWin {
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int nMax);
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int nMax);
    // All visible top-level windows of a given window class ("Chrome_WidgetWin_1"
    // for Edge). Used to find OUR new window by set-difference around the launch,
    // which works even while its title is still empty - the state in which the
    // title-based scans missed it and the splash painted at Edge's default size.
    public static IntPtr[] GetTopWindowsByClass(string cls) {
        var found = new List<IntPtr>();
        var sb = new StringBuilder(64);
        EnumWindows((h, l) => {
            if (IsWindowVisible(h)) {
                sb.Clear();
                GetClassName(h, sb, 64);
                if (sb.ToString() == cls) found.Add(h);
            }
            return true;
        }, IntPtr.Zero);
        return found.ToArray();
    }
    // v2.1.11: same enumeration WITHOUT the visibility filter. A tray-hidden
    // window still EXISTS (it is hidden, not destroyed), so counting hidden
    // frames distinguishes "minimised to the notification area" from "the
    // user closed the window" - which a visible-only scan cannot.
    public static IntPtr[] GetAllTopWindowsByClass(string cls) {
        var found = new List<IntPtr>();
        var sb = new StringBuilder(64);
        EnumWindows((h, l) => {
            sb.Clear();
            GetClassName(h, sb, 64);
            if (sb.ToString() == cls) found.Add(h);
            return true;
        }, IntPtr.Zero);
        return found.ToArray();
    }
    public static string GetTitle(IntPtr hWnd) {
        var sb = new StringBuilder(512);
        GetWindowText(hWnd, sb, 512);
        return sb.ToString();
    }
    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);
    [DllImport("kernel32.dll")]
    public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")]
    public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("user32.dll")]
    public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);
    // v2.1.18: close a window we already have a handle for. PostMessage is
    // asynchronous - it queues WM_CLOSE and returns at once - so a slow or
    // busy Edge cannot stall our loop the way a synchronous close can.
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsZoomed(IntPtr hWnd);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
}

function Set-WindowForeground([IntPtr]$hwnd) {
    # Bring an existing top-level window to the front. Restores only if minimised
    # (SW_RESTORE on a maximised window would knock it back to normal size), then
    # borrows foreground permission from the current foreground thread - Windows
    # blocks SetForegroundWindow from background processes otherwise.
    if ($hwnd -eq [IntPtr]::Zero) { return }
    try {
        if ([YVWin]::IsIconic($hwnd)) { [YVWin]::ShowWindow($hwnd, 9) | Out-Null }
        $fg = [YVWin]::GetForegroundWindow()
        $fgThread = [YVWin]::GetWindowThreadProcessId($fg, [IntPtr]::Zero)
        $me = [YVWin]::GetCurrentThreadId()
        $attached = $false
        if ($fgThread -ne 0 -and $fgThread -ne $me) {
            $attached = [YVWin]::AttachThreadInput($me, $fgThread, $true)
        }
        [YVWin]::SetForegroundWindow($hwnd) | Out-Null
        [YVWin]::BringWindowToTop($hwnd) | Out-Null
        if ($attached) { [YVWin]::AttachThreadInput($me, $fgThread, $false) | Out-Null }
    } catch {}
}

function Show-ItemInFront($item) {
    # Display an Outlook inspector and force it above the Edge app window.
    # Windows blocks focus-stealing by background processes; the transient
    # Alt keypress is the documented unlock for SetForegroundWindow.
    $item.Display()
    try {
        $insp = $item.GetInspector
        Start-Sleep -Milliseconds 200
        try { $insp.Activate() } catch {}
        $hwnd = [YVWin]::FindWindow('rctrl_renwnd32', [string]$insp.Caption)
        if ($hwnd -ne [IntPtr]::Zero) {
            [YVWin]::ShowWindow($hwnd, 9) | Out-Null    # SW_RESTORE if minimized
            # Borrow foreground permission from the current foreground thread
            # (no synthetic keystrokes, so no ribbon KeyTips side effect).
            $fg = [YVWin]::GetForegroundWindow()
            $fgThread = [YVWin]::GetWindowThreadProcessId($fg, [IntPtr]::Zero)
            $me = [YVWin]::GetCurrentThreadId()
            $attached = $false
            if ($fgThread -ne 0 -and $fgThread -ne $me) {
                $attached = [YVWin]::AttachThreadInput($me, $fgThread, $true)
            }
            [YVWin]::SetForegroundWindow($hwnd) | Out-Null
            [YVWin]::BringWindowToTop($hwnd) | Out-Null
            if ($attached) { [YVWin]::AttachThreadInput($me, $fgThread, $false) | Out-Null }
        }
    } catch {}
}

function Send-Http($stream, [string]$status, [string]$ctype, [byte[]]$body) {
    # S3 (v2.1.26): X-Content-Type-Options stops a browser second-guessing the
    # declared Content-Type. Free, and it removes a whole class of confusion
    # attack where a response is treated as a type it never claimed to be.
    $hdr = "HTTP/1.1 $status`r`nContent-Type: $ctype`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nConnection: close`r`n`r`n"
    $hb = [Text.Encoding]::ASCII.GetBytes($hdr)
    $stream.Write($hb, 0, $hb.Length)
    $stream.Write($body, 0, $body.Length)
    $stream.Flush()
    $script:respondedTo = $stream   # v2.0.24: identity, not a flag - Invoke-QuickServe
                                    # answers other connections mid-build, and a shared
                                    # boolean would mark THIS request as already answered
}
function JsonBytes([string]$s) { ,[Text.Encoding]::UTF8.GetBytes($s) }

function Save-WindowRectNow {
    # v2.1.21: capture the window geometry AT THIS MOMENT, rather than waiting
    # for the 2s sampler in the watchdog to come round again.
    #
    # Until now the remembered size came only from that sampler, which runs in
    # the accept-wait loop - so it does not run at all while a build is in
    # progress, and it never runs after the window has closed. A resize
    # followed by Quit could therefore be lost, and a resize made DURING an
    # update (the ten-minute revalidation included) was invisible for the
    # whole build. Calling this immediately before closing the window removes
    # the dependency on being sampled at a lucky moment.
    #
    # Mirrors the sampler's rules exactly: skip while launch enforcement is
    # still running (its transition rects are not user choices) and while
    # minimised (bogus -32000 coordinates); a maximised window records the
    # last known NORMAL rect with the max flag set, so restore-down still
    # returns to the user's size.
    try {
        if ($null -ne $script:resizeTarget) { return }
        if ($script:appHwnd -eq [IntPtr]::Zero -or -not [YVWin]::IsWindow($script:appHwnd)) { return }
        if ([YVWin]::IsIconic($script:appHwnd)) { return }
        if ([YVWin]::IsZoomed($script:appHwnd)) {
            if ($script:lastNormalRect) {
                Save-WindowRect ([int]$script:lastNormalRect[0]) ([int]$script:lastNormalRect[1]) `
                    ([int]$script:lastNormalRect[2]) ([int]$script:lastNormalRect[3]) $true
            }
        } else {
            $rc = New-Object YVWin+RECT
            if ([YVWin]::GetWindowRect($script:appHwnd, [ref]$rc)) {
                Save-WindowRect ($rc.Right - $rc.Left) ($rc.Bottom - $rc.Top) $rc.Left $rc.Top $false
            }
        }
    } catch {}
}
function Close-ViewWindow {
    # v2.1.19: shared by BOTH quit paths - the main-loop route and the in-build
    # quick-serve below. Handle first (v2.1.18): $script:appHwnd is captured at
    # startup and tracked by the watchdog, so it is immune to the page retitling
    # itself and to how fast the machine is. PostMessage is asynchronous, so a
    # busy Edge cannot stall us. The title scan remains only for the case with
    # no tracked handle (a hand-typed URL, or -NoOpen).
    #
    # v2.1.21: record the geometry BEFORE the window goes - once it is closed
    # there is nothing left to measure.
    Save-WindowRectNow
    try {
        if ($script:appHwnd -ne [IntPtr]::Zero -and [YVWin]::IsWindow($script:appHwnd)) {
            [void][YVWin]::PostMessage($script:appHwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
            return
        }
        foreach ($attempt in 1..2) {
            $win = Get-Process msedge -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowTitle -like '* YearView' -or
                               $_.MainWindowTitle -like 'YearView (stopped)*' } |
                Select-Object -First 1
            if ($win) { [void]$win.CloseMainWindow(); break }
            Start-Sleep -Milliseconds 500
        }
    } catch {}
}
function Invoke-QuickServe {
    # Cooperative mini-server (v1.2.23), called periodically FROM WITHIN long
    # builds so the app stays responsive while the single-threaded main loop is
    # busy (startup revalidation, upgrade-day dual validation, any build). It
    # serves ONLY instantly-answerable, strictly read-only GETs - cached year
    # pages, the splash for uncached years, the favicon - applying the SAME
    # Host and Sec-Fetch-Site checks as the main loop. Everything else gets
    # 503 {"busy":true} and the caller retries once the build finishes. No
    # route here may touch Outlook, start a build, or mutate any state; window
    # expansion is deliberately skipped (mutating; the window is already up).
    if (-not $script:listener) { return }
    try {
        while ($script:listener.Pending()) {
            $qc = $script:listener.AcceptTcpClient()
            try {
                $qs = $qc.GetStream()
                $qs.ReadTimeout = 500
                $qbuf = New-Object byte[] 8192
                $qtotal = 0; $qtext = ''
                $qDeadline = (Get-Date).AddMilliseconds(800)
                while ($qtotal -lt $qbuf.Length) {
                    if ((Get-Date) -gt $qDeadline) { $qtext = ''; break }
                    $qr = $qs.Read($qbuf, $qtotal, $qbuf.Length - $qtotal)
                    if ($qr -le 0) { break }
                    $qtotal += $qr
                    $qtext = [Text.Encoding]::ASCII.GetString($qbuf, 0, $qtotal)
                    if ($qtext.Contains("`r`n`r`n")) { break }
                }
                if ($qtext) {
                    $ql = $qtext -split "`r`n"
                    $qp = $ql[0] -split ' '
                    if ($qp.Count -ge 2) {
                        $qm = $qp[0].ToUpperInvariant()
                        $qpath = $qp[1]
                        $qhost = ($ql | Where-Object { $_ -match '^(?i)Host:\s*(.+)$' } |
                            ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
                        $qsite = ($ql | Where-Object { $_ -match '^(?i)Sec-Fetch-Site:\s*(.+)$' } |
                            ForEach-Object { $Matches[1].Trim().ToLowerInvariant() } | Select-Object -First 1)
                        if (($qhost -and $qhost -notmatch '^(127\.0\.0\.1|localhost)(:\d+)?$') -or
                            ($qsite -and $qsite -ne 'none' -and $qsite -ne 'same-origin')) {
                            Send-Http $qs '403 Forbidden' 'text/plain' (JsonBytes 'forbidden')
                        }
                        elseif ($qm -eq 'GET' -and ($qpath -eq '/' -or $qpath -eq '/index.html')) {
                            if ($script:pages.ContainsKey($Year)) {
                                Send-Http $qs '200 OK' 'text/html; charset=utf-8' $script:pages[$Year]
                            } else {
                                Send-Http $qs '200 OK' 'text/html; charset=utf-8' $splashBytes
                            }
                        }
                        elseif ($qm -eq 'GET' -and ($qpath.StartsWith('/?') -or $qpath.StartsWith('/index.html?'))) {
                            $qy = Get-QueryYear $qpath $Year
                            if ($script:pages.ContainsKey($qy)) {
                                Send-Http $qs '200 OK' 'text/html; charset=utf-8' $script:pages[$qy]
                            } else {
                                Send-Http $qs '200 OK' 'text/html; charset=utf-8' (Get-SplashBytes $qy)
                            }
                        }
                        elseif ($qpath.StartsWith('/quit')) {
                            # v2.1.19: Quit is honoured DURING a build. It used
                            # to fall through to the 503 below, so pressing
                            # Quit while an update was running was answered
                            # "busy" and the client retried until it gave up -
                            # the window never closed. A fast machine hid this
                            # because builds finish before the retries run out;
                            # a slower laptop does not, which is exactly where
                            # it was reported.
                            #
                            # Quit is the ONE state-changing thing this
                            # mini-server may do, and it is safe precisely
                            # because it changes nothing a build depends on: it
                            # sets the stop flag and closes the window. Unlike
                            # every read-only route here, it MUST carry the
                            # session token - it is not a page GET.
                            $qtok = ($ql | Where-Object { $_ -match '^(?i)X-YV-Token:\s*(.+)$' } |
                                ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
                            if ($qm -eq 'POST' -and $qtok -is [string] -and $qtok -ceq $script:sessionToken) {
                                Send-Http $qs '200 OK' 'application/json' (JsonBytes '{"ok":true}')
                                $script:quitRequested = $true
                                Write-Host 'Quit received during an update - closing the window now; the update will finish first.' -ForegroundColor Yellow
                                Close-ViewWindow
                            } else {
                                Send-Http $qs '403 Forbidden' 'application/json' (JsonBytes '{"ok":false}')
                            }
                        }
                        else {
                            Send-Http $qs '503 Service Unavailable' 'application/json' `
                                (JsonBytes '{"ok":false,"busy":true}')
                        }
                    }
                }
            } catch {} finally { try { $qc.Close() } catch {} }
        }
    } catch {}
}

function Write-CfgAtomic([string]$path, [string]$text) {
    # R3 (v2.1.26): write a small config file ATOMICALLY. These were plain
    # WriteAllText calls, so a crash or power loss between truncate and write
    # left a truncated file. The page cache has always written .tmp then renamed
    # over the target; the config files - filter, window rect, calendar
    # selection - had not. Load-WindowRect validates and falls back safely, but
    # Load-Filter would happily read a truncated CSV and apply a WRONG filter,
    # silently hiding or showing the wrong categories. Same-volume Move is
    # effectively atomic, so a reader sees either the old file or the new one.
    # v2.1.29: VERIFY the write, and say so when it fails. Every caller wraps
    # this in a try/catch that swallows errors, so a failing write was
    # indistinguishable from a successful one - the setting applied for the rest
    # of the session (the page holds its own copy) and was simply absent next
    # start. That is exactly the shape of the "filter forgotten again" report,
    # and it must not be able to fail silently a third time.
    $tmp = $path + '.tmp'
    $ok = $false
    try {
        [IO.File]::WriteAllText($tmp, $text)
        Move-Item -LiteralPath $tmp -Destination $path -Force
        $ok = ((Test-Path $path) -and ([IO.File]::ReadAllText($path) -ceq $text))
    } catch { $ok = $false }
    if (-not $ok) {
        # Fall back to a plain write: a torn file is a far smaller problem than
        # a setting that never persists at all.
        try {
            [IO.File]::WriteAllText($path, $text)
            $ok = ((Test-Path $path) -and ([IO.File]::ReadAllText($path) -ceq $text))
        } catch { $ok = $false }
        try { if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue } } catch {}
        try {
            $lg = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
            Add-Content -Path $lg -Value ("[{0}] config write {1}: {2}" -f `
                (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if ($ok) { 'needed fallback' } else { 'FAILED' }), $path)
        } catch {}
    }
    if (-not $ok) { throw ("Could not save {0}" -f (Split-Path -Leaf $path)) }
}
function Save-Settings {
    # v2.1.29: user settings, same shape as the other config files. One
    # key=value per line so more settings can be added without a format change.
    Write-CfgAtomic (Join-Path (Split-Path -Parent $PSCommandPath) 'yv-settings.cfg') `
        (('confirmdelete=' + $(if ($script:confirmDelete) { '1' } else { '0' })) + "`r`n" +
         ('scrollyears='   + $(if ($script:scrollYears)   { '1' } else { '0' })) + "`r`n" +
         ('notitles='      + [string]$script:noTitles) + "`r`n" +
         ('printpages='    + [string]$script:printPages) + "`r`n" +
         ('printlegend='   + $(if ($script:printLegend)   { '1' } else { '0' })) + "`r`n" +
         ('startmonth='    + [string]$script:startMonth))
}
function Load-Settings {
    # Absent file means defaults, which is what a first run should get. Only an
    # explicit '0' turns confirmation off, so a damaged or partial file leaves
    # the SAFER behaviour in place rather than silently disabling the prompt.
    try {
        $p = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-settings.cfg'
        if (Test-Path $p) {
            foreach ($line in [IO.File]::ReadAllLines($p)) {
                $kvp = $line -split '=', 2
                if ($kvp.Count -ne 2) { continue }
                switch ($kvp[0].Trim()) {
                    'confirmdelete' { $script:confirmDelete = ($kvp[1].Trim() -ne '0') }
                    'scrollyears'   { $script:scrollYears   = ($kvp[1].Trim() -ne '0') }
                    'printlegend'   { $script:printLegend   = ($kvp[1].Trim() -ne '0') }
                    'notitles'      {
                        # v3.4.7: NOT trimmed - a category name is matched exactly.
                        if (Test-TitleCsv $kvp[1]) { $script:noTitles = [string]$kvp[1] }
                    }
                    'startmonth'    {
                        # v3.4.0: -1 (follow the current month) through 11.
                        $sm = 0
                        if ([int]::TryParse($kvp[1].Trim(), [ref]$sm) -and $sm -ge -1 -and $sm -le 11) {
                            $script:startMonth = $sm
                        }
                    }
                    'printpages'    {
                        # v3.2.0: a number, not a flag. Anything unparseable or
                        # out of range leaves the default rather than producing
                        # a scale factor nobody asked for.
                        $n = 0
                        if ([int]::TryParse($kvp[1].Trim(), [ref]$n) -and $n -ge 0 -and $n -le 4) {
                            $script:printPages = $n
                        }
                    }
                }
            }
        }
    } catch {}
}
function Save-Filter([string]$csv) {
    # v2.0.8: persist the hidden-category filter beside the script (replacing
    # browser localStorage, which lived in Edge's profile beyond the reach of
    # uninstall). Stored verbatim as the client's csv; '' is a real value
    # meaning "show all" and is distinct from the file being absent.
    try {
        Write-CfgAtomic (Join-Path (Split-Path -Parent $PSCommandPath) 'yv-filter.cfg') ([string]$csv)
        $script:hiddenCsv = [string]$csv
    } catch {}
}

function Load-Filter {
    try {
        $p = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-filter.cfg'
        if (Test-Path $p) { $script:hiddenCsv = [IO.File]::ReadAllText($p) }
    } catch {}
}

function Save-WindowRect([int]$w, [int]$h, [int]$x, [int]$y, [bool]$isMax = $false) {
    # v2.0.16: remember the geometry the USER settled on. Callers filter out
    # minimised states and the launch-enforcement window; this only
    # sanity-bounds and skips no-change writes. v2.0.17: a fifth field records
    # a maximised state - the w,h,x,y remain the NORMAL rect (what restore-
    # down returns to), never the maximised bounds themselves.
    if ($w -lt 300 -or $h -lt 200 -or $w -gt 10000 -or $h -gt 10000) { return }
    $csv = "$w,$h,$x,$y," + $(if ($isMax) { '1' } else { '0' })
    if ($csv -eq $script:lastWinRect) { return }
    try {
        Write-CfgAtomic (Join-Path (Split-Path -Parent $PSCommandPath) 'yv-window.cfg') $csv
        $script:lastWinRect = $csv
        if (-not $isMax) { $script:lastNormalRect = @($w, $h, $x, $y) }
    } catch {}
}

function Load-StartupState {
    # R7 (v2.1.26): every piece of persisted state, loaded in ONE place that
    # BOTH launch paths call.
    #
    # There are two ways YearView starts. The ordinary path binds the port and
    # runs the whole startup sequence. The ATTACH path - taken when the port is
    # already held by a live bridge with no window, which happens when YearView
    # is restarted within a few seconds of closing - opens a view against that
    # bridge and returns early, skipping everything below it. v2.1.25 fixed the
    # window geometry being missed on that path by adding one call to it; this
    # removes the class of fault rather than that instance, because the next
    # piece of startup state added would have had exactly the same problem.
    #
    # Order matters and is deliberate: the calendar selection is needed before
    # the page cache is restored (the cache gate compares against it), and the
    # filter must be read before Restore-PageCache normalises each cached
    # page's seed against it. Restore-PageCache itself is NOT called here - it
    # populates the server's page cache, which only the serving path has any
    # use for.
    Load-CalSelection
    Load-Filter
    Load-Settings       # v2.1.29: before Restore-PageCache, which normalises each page against it
    Load-WindowRect
}
function Load-WindowRect {
    # Restore only a rect that still makes sense on TODAY'S monitors: sane
    # dimensions, and at least a 100x100 corner visible on some screen's work
    # area (layouts change between sessions - a rect for a detached monitor
    # must not open the window off-screen).
    try {
        $p = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-window.cfg'
        if (-not (Test-Path $p)) { return }
        $parts = ([IO.File]::ReadAllText($p).Trim() -split ',')
        if ($parts.Count -lt 4 -or $parts.Count -gt 5) { return }
        $w = [int]$parts[0]; $h = [int]$parts[1]; $x = [int]$parts[2]; $y = [int]$parts[3]
        $isMax = ($parts.Count -eq 5 -and $parts[4].Trim() -eq '1')   # 4-field files (v2.0.16) load as not-maximised
        if ($w -lt 300 -or $h -lt 200 -or $w -gt 10000 -or $h -gt 10000) { return }
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $visible = $false
        foreach ($scr in [System.Windows.Forms.Screen]::AllScreens) {
            $wa = $scr.WorkingArea
            $ix = [Math]::Min($x + $w, $wa.X + $wa.Width)  - [Math]::Max($x, $wa.X)
            $iy = [Math]::Min($y + $h, $wa.Y + $wa.Height) - [Math]::Max($y, $wa.Y)
            if ($ix -ge 100 -and $iy -ge 100) { $visible = $true; break }
        }
        if (-not $visible) { return }
        $script:savedWinRect = @($w, $h, $x, $y)
        $script:savedWinMax = $isMax
        $script:lastNormalRect = @($w, $h, $x, $y)
        $script:lastWinRect = "$w,$h,$x,$y," + $(if ($isMax) { '1' } else { '0' })
    } catch {}
}

function Save-CalSelection {
    # Persist which calendar is selected (plain identifiers - they are only
    # meaningful inside this Windows profile). Empty selection removes the file.
    try {
        $p = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-calendar.cfg'
        if ($script:selCalId) {
            Write-CfgAtomic $p (($script:selCalId, $script:selCalStoreId, $script:selCalName) -join "`r`n")
        } elseif (Test-Path $p) { Remove-Item $p -Force -ErrorAction SilentlyContinue }
    } catch {}
}

function Load-CalSelection {
    # Startup, BEFORE any Outlook contact: just loads the identifiers; actual
    # resolution (and fallback to default) happens lazily in Get-SelectedCalendar.
    try {
        $p = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-calendar.cfg'
        if (Test-Path $p) {
            $l = @(Get-Content $p -ErrorAction Stop)
            if ($l.Count -ge 2 -and $l[0] -and $l[1]) {
                $script:selCalId = [string]$l[0]
                $script:selCalStoreId = [string]$l[1]
                if ($l.Count -ge 3) { $script:selCalName = [string]$l[2] }
            }
        }
    } catch {}
}

function Assert-SelectedCalendarItem($Item) {
    # v2.1.28: authorise an item BEFORE acting on it.
    #
    # GetItemFromID is scoped to the STORE, not the folder, so an identifier
    # naming an appointment in a DIFFERENT calendar of the same mailbox
    # resolved happily. The MessageClass test caught mail and contacts but not
    # that. Multiple calendars in one store is exactly what the calendar picker
    # offers, so this was reachable in ordinary use.
    #
    # The realistic route to it is not an attacker - reaching /open or /delete
    # already requires the session token, which lives only in our own page - it
    # is a STALE PAGE. Identifiers are baked into a page when it is built, and
    # pages outlive their context: a second window, a page restored from cache,
    # or a delete already in flight when the calendar is switched. Any of those
    # can name an item in the calendar you were looking at BEFORE. Deleting is
    # silent and unconfirmed, so the first sign would be a missing appointment
    # in a calendar you were not even viewing.
    #
    # Fails closed: no item, wrong class, or a parent that is not the selected
    # calendar folder all refuse. Get-SelectedCalendar is the single source of
    # that folder and already handles its own edge cases (no selection yet ->
    # default calendar; folder deleted -> revert and re-persist).
    if (-not $Item) { throw 'Calendar item not found.' }
    if ([string]$Item.MessageClass -notlike 'IPM.Appointment*') { throw 'Not a calendar item.' }
    $cal = Get-SelectedCalendar
    $parent = $null
    try { $parent = $Item.Parent } catch {}
    if (-not $parent) { throw 'Calendar item has no reachable parent folder.' }
    if ([string]$parent.EntryID -ne [string]$cal.EntryID -or
        [string]$parent.StoreID -ne [string]$cal.StoreID) {
        throw 'That item is not in the calendar YearView is showing - reopen YearView and try again.'
    }
    $Item
}
function Get-SelectedCalendar {
    # THE single source of the calendar folder. Every reader, watcher, creator
    # and deleter goes through here. Any resolution failure - folder deleted,
    # store detached - reverts to the default calendar, clears the persisted
    # selection, and says so, rather than erroring per-operation.
    if (-not $script:selCalId) { return $script:ns.GetDefaultFolder(9) }
    try {
        $f = $script:ns.GetFolderFromID($script:selCalId, $script:selCalStoreId)
        if ($f -and $f.DefaultItemType -eq 1) { return $f }
    } catch {}
    Write-Warning ("Selected calendar '{0}' no longer resolves - reverting to the default calendar." -f $script:selCalName)
    $script:selCalId = ''; $script:selCalStoreId = ''; $script:selCalName = ''
    Save-CalSelection
    $script:ns.GetDefaultFolder(9)
}

function Save-PageCache([int]$y) {
    # Persist the freshly built page + its content hash + the session token that
    # is baked into the page bytes, so the NEXT bridge start can serve this year
    # instantly (substituting its own token) and revalidate silently.
    try {
        $dir = Split-Path -Parent $PSCommandPath
        # v1.2.29: DPAPI-encrypt both files (CurrentUser scope - calendar content
        # at rest is no longer plaintext), and write ATOMICALLY: a crash or power
        # loss mid-write must never leave a truncated file to be served. Write to
        # a .tmp sibling, then rename over the target (same-volume, ~atomic).
        $scope = [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        $pageEnc = [System.Security.Cryptography.ProtectedData]::Protect($script:pages[$y], $null, $scope)
        $metaEnc = [System.Security.Cryptography.ProtectedData]::Protect(
            [Text.Encoding]::UTF8.GetBytes(([string]$script:contentHashes[$y] + "`n" + [string]$script:sessionToken + "`n" + $script:yvVersion + "`n" + $script:selCalId)),
            $null, $scope)
        $pgPath = Join-Path $dir ("yv-page-{0}.html" -f $y)
        $mtPath = Join-Path $dir ("yv-page-{0}.meta" -f $y)
        [IO.File]::WriteAllBytes(($pgPath + '.tmp'), $pageEnc)
        [IO.File]::WriteAllBytes(($mtPath + '.tmp'), $metaEnc)
        Move-Item -LiteralPath ($pgPath + '.tmp') -Destination $pgPath -Force
        Move-Item -LiteralPath ($mtPath + '.tmp') -Destination $mtPath -Force
        # prune: keep the ten most recent years on disk (all realistically
        # visited years survive restarts; ~250 KB each)
        Get-ChildItem (Join-Path $dir 'yv-page-*.html') -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -Skip 10 | ForEach-Object {
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                Remove-Item ($_.FullName -replace '\.html$', '.meta') -Force -ErrorAction SilentlyContinue
            }
    } catch {}
}

function Set-SpillPayload([string]$html, [string]$json) {
    # v3.4.3: replace the contents of the second data island in a page that is
    # already built. The payload had '</' neutralised when it was serialised,
    # so nothing inside it can close a script element and the two boundaries
    # are unambiguous - the same reasoning that lets the filter seed be
    # re-stamped in place rather than the page rebuilt (v2.1.22).
    # The boundary search below is only safe while the payload cannot close
    # a script element. Get-YearData already neutralises '</' when it
    # serialises, but this function is not entitled to assume its caller
    # did - so it does it again. Harmless when already done: '<\/' has no
    # '</' left to find.
    $json = $json -replace '</', '<\/'
    $open  = '<script id="yvdata2" type="application/json">'
    $close = '</script>'
    $i = $html.IndexOf($open)
    if ($i -lt 0) { return $html }
    $s = $i + $open.Length
    $e = $html.IndexOf($close, $s)
    if ($e -lt $s) { return $html }
    return $html.Substring(0, $s) + "`r`n" + $json + "`r`n" + $html.Substring($e)
}
function Update-SettingLiterals([string]$html) {
    # v3.1.8: the settings are baked into a page at BUILD time, and
    # yv-settings.cfg is the authority (v2.1.24). This was inline in
    # Restore-PageCache, which meant it only ran when a cached page was read
    # from disk at startup - a setting changed WHILE a page was in memory left
    # every cached page carrying the old value, and a background update that
    # navigated to one of them showed the setting reverted. It is a function
    # now so /setsetting can apply it too, the way /setfilter has re-stamped
    # the filter seed since v2.1.22. The value is a bare true/false, so it
    # ends at the ';'.
    foreach ($pair in @(@('const CONFIRM_DELETE = ', $script:confirmDelete),
                        @('const SCROLL_YEARS = ',   $script:scrollYears),
                        @('const NO_TITLES = ',      [string]$script:noTitles),
                        @('const PRINT_PAGES = ',    $script:printPages),
                        @('const PRINT_LEGEND = ',   $script:printLegend),
                        @('const START_MONTH = ',    $script:startMonth))) {
        # v3.4.9: the marker is matched at the START OF A LINE, as it is
        # written in the client. A bare IndexOf found the FIRST occurrence
        # anywhere - and the data islands come before the code, so an
        # appointment titled "const START_MONTH = " was patched instead, from
        # there to the next ';', and the page came up blank. Calendar text in
        # the page is always inside a JSON string, where a newline is escaped,
        # so a line-anchored marker can only ever match the real declaration.
        $cm = "`n  " + [string]$pair[0]
        $c0 = $html.IndexOf($cm)
        if ($c0 -lt 0) { continue }
        $c1 = $html.IndexOf(';', $c0 + $cm.Length)
        # v3.2.0: the value is a bool for two of these and a NUMBER for the
        # third, so render by type rather than assuming true/false.
        # v3.3.0: bool, number or STRING - a string has to go back quoted, or
        # the re-stamped page would carry a bare word and fail to parse.
        $cur = $(if ($pair[1] -is [bool]) { $(if ($pair[1]) { 'true' } else { 'false' }) }
                 elseif ($pair[1] -is [string]) { (ConvertTo-SeedStr ([string]$pair[1])) }
                 else { [string][int]$pair[1] })
        if ($c1 -gt $c0 -and $html.Substring($c0 + $cm.Length, $c1 - $c0 - $cm.Length) -ne $cur) {
            $html = $html.Substring(0, $c0 + $cm.Length) + $cur + $html.Substring($c1)
        }
    }
    return $html
}
function Restore-PageCache {
    # Load previously built pages so the first request serves INSTANTLY (no
    # splash, no build). Each page's embedded token belongs to the PREVIOUS
    # session, so it is substituted with this session's token; the previous
    # token is dead the moment that session ended, so the on-disk copy holds no
    # live secret. The stale BUILD_TS baked into the page makes the very first
    # /dirty poll report fingerprint-only dirt (lastChangeTicks is set to 'now'
    # below), driving ONE silent revalidating rebuild through the existing
    # quiet path: unchanged calendar -> the page simply stays; changed ->
    # refresh with the "YearView updated" toast.
    $loaded = 0
    try {
        $dir = Split-Path -Parent $PSCommandPath
        foreach ($f in @(Get-ChildItem (Join-Path $dir 'yv-page-*.html') -ErrorAction SilentlyContinue)) {
            try {
                if ($f.Name -notmatch '^yv-page-(\d{4})\.html$') { continue }
                $y = [int]$Matches[1]
                $meta = Join-Path $dir ("yv-page-{0}.meta" -f $y)
                if (-not (Test-Path $meta)) { continue }
                # v1.2.29: files are DPAPI-encrypted; decrypt, falling back to
                # the raw bytes for plaintext caches written by <=1.2.28 (they
                # are re-encrypted on the next save). A wrong-user or corrupt
                # blob throws on BOTH reads and the year is simply skipped.
                $scope = [System.Security.Cryptography.DataProtectionScope]::CurrentUser
                $mBytes = [IO.File]::ReadAllBytes($meta)
                try { $mBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($mBytes, $null, $scope) } catch {}
                $ml = @(([Text.Encoding]::UTF8.GetString($mBytes)) -split "\r?\n" | Where-Object { $_ })
                if ($ml.Count -lt 2 -or -not $ml[1]) { continue }
                # v1.2.31: pages embed the building version's client code; a
                # cache from any OTHER version (or the pre-version 2-line
                # format) is skipped and that year rebuilds fresh - upgrades
                # self-heal instead of running stale JavaScript.
                if ($ml.Count -lt 3 -or $ml[2] -ne $script:yvVersion) {
                    # R4 (v2.1.26): a page built by ANOTHER version can never be
                    # served by this one, so it is dead weight - it used to sit
                    # there until count-based pruning happened to remove it,
                    # wasting space and, worse, showing up in the folder during
                    # debugging as though it were live. Delete it now. (The
                    # calendar-mismatch skip below deliberately does NOT delete:
                    # that page is valid for a different calendar and is wanted
                    # again the moment the user switches back.)
                    Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
                    Remove-Item $meta -Force -ErrorAction SilentlyContinue
                    continue
                }
                # v2.0.0: a cached page renders ONE calendar's content; only
                # restore pages built for the currently selected calendar
                # (empty line 4 = default; empties are filtered, so absent
                # line 4 means default too).
                $mCal = $(if ($ml.Count -ge 4) { [string]$ml[3] } else { '' })
                if ($mCal -ne $script:selCalId) { continue }
                $pBytes = [IO.File]::ReadAllBytes($f.FullName)
                try { $pBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($pBytes, $null, $scope) } catch {}
                $html = [Text.Encoding]::UTF8.GetString($pBytes)
                $html = $html.Replace([string]$ml[1], $script:sessionToken)   # old token -> this session's
                # v2.1.22: recover the HIDE_SEED literal this page carries so a
                # later /setfilter can re-stamp it. The literal is a complete
                # JSON string; it ends at its closing quote, NOT at the ';'.
                # v2.1.24: FORCE the restored page's filter seed to match
                # yv-filter.cfg, which Load-Filter has already read (it runs
                # before this function - deliberately).
                #
                # v2.1.22 tried to keep the cache in step by patching every
                # page at the moment /setfilter ran. That is a WRITE-path fix,
                # and it only holds if that path runs at the right moment with
                # the right page in memory; anything that misses - a page not
                # loaded then, a change that never reached the route, a cache
                # written by some other path - leaves the old seed on disk with
                # nothing to correct it, which is why the symptom came back.
                #
                # The seed is not calendar data. The cache exists to hold
                # CONTENT; the filter is separate state with its own file and
                # its own authority. So rather than trying to keep two copies
                # equal, the copy in the page is simply overwritten with the
                # authoritative one every time a page is loaded. A stale seed
                # then cannot survive a restart no matter how it arose.
                try {
                    # v3.4.9: line-anchored, as in Update-SettingLiterals
                    $mk = "`n  const HIDE_SEED = "
                    $a0 = $html.IndexOf($mk)
                    if ($a0 -ge 0) {
                        # walk the JSON string literal to its CLOSING unescaped
                        # quote (a ';' or '"' inside a category name is escaped
                        # by ConvertTo-JsonStr, so scanning for ';' would truncate)
                        $p = $a0 + $mk.Length
                        if ($html[$p] -eq '"') {
                            $q = $p + 1; $esc = $false
                            while ($q -lt $html.Length) {
                                $ch = $html[$q]
                                if ($esc) { $esc = $false }
                                elseif ($ch -eq '\') { $esc = $true }
                                elseif ($ch -eq '"') { break }
                                $q++
                            }
                            if ($q -lt $html.Length) {
                                $wasLit = $html.Substring($p, $q - $p + 1)
                                $curLit = ConvertTo-SeedStr $(if ($null -ne $script:hiddenCsv) { $script:hiddenCsv } else { '' })
                                if ($wasLit -ne $curLit) {
                                    $html = $html.Substring(0, $p) + $curLit + $html.Substring($q + 1)
                                    Write-Host ("  filter seed of cached {0} refreshed from yv-filter.cfg" -f $y) -ForegroundColor DarkGray
                                }
                                $script:pageHideLits[$y] = $curLit
                            }
                        }
                    }
                } catch {}
                # v2.1.29: same read-time normalisation the filter seed gets
                # (v2.1.24) - the cached page carries whatever the setting was
                # when it was built, and yv-settings.cfg is the authority.
                try { $html = Update-SettingLiterals $html } catch {}
                $script:pages[$y] = [Text.Encoding]::UTF8.GetBytes($html)
                $script:contentHashes[$y] = [string]$ml[0]
                $loaded++
            } catch {}
        }
    } catch {}
    if ($loaded -gt 0) {
        # force one silent revalidation of any restored page that gets served
        $script:lastChangeTicks = (Get-Date).Ticks
        Write-Host ("Restored {0} cached year page(s) from disk - serving instantly, revalidating in the background." -f $loaded) -ForegroundColor Green
    }
}
function JsonError([string]$msg) {
    # R-15: real JSON encoding - control characters, quotes and backslashes in
    # Outlook error text must not be able to break the payload the client parses.
    #
    # v2.0.23: every failed Outlook operation reports through here, so this is
    # also where a DEAD COM LINK is recognised. Ensure-Outlook probes liveness
    # before each COM route, but Outlook can exit BETWEEN that probe and the
    # operation - most easily when YearView started Outlook headless and the
    # user closes the last inspector, leaving it with no UI. No probe can close
    # that time-of-check-to-time-of-use gap; only a retry can.
    #
    # HRESULTs are matched numerically because the message text is localised.
    # Null the link but LEAVE $script:watchItems set, so the next Ensure-Outlook
    # takes its reconnect branch (which forces a structural refresh) rather than
    # its first-init branch. The client is told the call is safe to repeat: a
    # dead link means the operation never reached Outlook, so nothing was
    # created, opened or deleted.
    if ($msg -match '0x800706BA|0x800706BE|0x80010108|0x800401FD|0x80080005|0x80010105') {
        Write-Host 'Outlook link lost mid-operation - resetting for reconnect.' -ForegroundColor Yellow
        Release-OutlookLink   # v2.1.0: bare nulls orphaned live RCWs and pinned Outlook
        return JsonBytes ((@{ ok = $false; retry = $true
                              err = 'Outlook connection was lost' } | ConvertTo-Json -Compress))
    }
    JsonBytes ((@{ ok = $false; err = [string]$msg } | ConvertTo-Json -Compress))
}
function Get-QueryYear([string]$p, [int]$default) {
    if ($p.Contains('?')) {
        foreach ($pair in ($p.Substring($p.IndexOf('?') + 1) -split '&')) {
            $kvp = $pair -split '=', 2
            if ($kvp.Count -eq 2 -and $kvp[0] -eq 'year') {
                $y = 0
                if ([int]::TryParse($kvp[1], [ref]$y) -and $y -ge 1990 -and $y -le 2100) { return $y }
            }
        }
    }
    return $default
}

$url = "http://127.0.0.1:$Port/"
function Open-View {
    if ($App) {
        $edge = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'
        if (-not (Test-Path $edge)) { $edge = Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe' }
        if (Test-Path $edge) {
            # v3.3.2: its OWN profile directory. Without one the app window
            # joins the user's ordinary Edge browser process, and a modal
            # print preview there runs a nested loop that can wedge EVERY Edge
            # window - the reported freeze that needed a reboot to clear. A
            # separate profile means a separate browser process, so YearView
            # printing cannot reach the user's browsing at all. It lives under
            # the data directory; since v3.4.9 Remove-YearViewData.cmd closes
            # this window and removes the profile on uninstall (before that,
            # this comment claimed it did and it did not).
            $edgeProfile = Join-Path (Split-Path -Parent $PSCommandPath) 'edge-profile'
            try { if (-not (Test-Path $edgeProfile)) { [void](New-Item -ItemType Directory -Path $edgeProfile -Force) } } catch {}
            $edgeArgs = @('--app=' + $url,
                          '--user-data-dir=' + $edgeProfile,
                          '--no-first-run',
                          '--no-default-browser-check')
            $bw = 0; $bh = 0; $bx = 0; $by = 0
            $sw = 0; $sh = 0; $sx = 0; $sy = 0
            try {
                $bw = $WindowWidth
                $bh = $WindowHeight
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
                $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
                if ($bw -le 0 -and $bh -le 0 -and $script:savedWinRect) {
                    # v2.0.16: a remembered user geometry replaces the computed
                    # default wholesale (size AND position). Explicit
                    # -WindowWidth/-WindowHeight always win; the rect was
                    # validated against today's monitors at load time.
                    $bw = [int]$script:savedWinRect[0]; $bh = [int]$script:savedWinRect[1]
                    $bx = [int]$script:savedWinRect[2]; $by = [int]$script:savedWinRect[3]
                } else {
                    if ($bw -le 0) { $bw = [Math]::Min([int]($wa.Width  * 0.92), 2200) }
                    if ($bh -le 0) { $bh = [Math]::Min([int]($bw * 10 / 16), [int]($wa.Height * 0.95)) }   # 16:10
                    $bx = $wa.X + [int](($wa.Width  - $bw) / 2)
                    $by = $wa.Y + [int](($wa.Height - $bh) / 2)
                }
                # v2.0.13: the splash opens at the FULL window geometry, not a
                # small centred box. Edge carries the geometry of the window it
                # opened most recently over to the next one, so a 460x210 splash
                # left ordinary Edge sessions started while YearView was running
                # opening at splash size. Using one geometry throughout also
                # removes the mid-load resize: the year view renders into a
                # window that is already the right size, with no visible jump.
                $sw = $bw; $sh = $bh; $sx = $bx; $sy = $by
                $edgeArgs += @("--window-size=$sw,$sh", "--window-position=$sx,$sy")
            } catch {
                Write-Verbose 'Could not read screen metrics; letting Edge choose the size.'
            }
            # Snapshot Edge's top-level windows BEFORE launching: the new window is
            # then identified by set-difference, independent of its title. Title-based
            # capture missed the pre-paint moment whenever the window was created with
            # an empty title (and Process.MainWindowTitle can only ever surface one
            # window per process anyway) - the cause of the large-splash flash.
            $preWins = @()
            try { $preWins = @([YVWin]::GetTopWindowsByClass('Chrome_WidgetWin_1')) } catch {}
            Start-Process $edge -ArgumentList $edgeArgs
            # Edge ignores --window-size when an existing process adopts the window, so
            # the bounds are enforced via MoveWindow once the page has loaded. That is
            # done non-blocking (see Invoke-WindowResize) because the page cannot load
            # until this script starts serving.
            if ($bw -gt 0) {
                $script:lastNormalRect = @($bw, $bh, $bx, $by)   # v2.0.17: known normal rect from the start
                $script:resizeTarget = @{
                    Stage = 'splash'
                    SX = $sx; SY = $sy; SW = $sw; SH = $sh
                    BX = $bx; BY = $by; BW = $bw; BH = $bh
                    # v2.0.17: restore a remembered maximised state. Only on the
                    # saved-geometry path - explicit -WindowWidth/-WindowHeight
                    # mean the user asked for that size, not maximised.
                    MaxFlag = [bool]($WindowWidth -le 0 -and $WindowHeight -le 0 -and $script:savedWinMax)
                    Deadline = (Get-Date).AddSeconds(30)
                    Pre = $preWins
                }
                # Tight catch: shrink the new window the instant it exists - before
                # it has painted at whatever stale size Edge chose - rather than
                # waiting for the serving loop's slower enforcement cadence.
                for ($i = 0; $i -lt 40; $i++) {
                    Start-Sleep -Milliseconds 40
                    if (Invoke-WindowResize) { break }                       # full page found or expired
                    if (-not $script:resizeTarget) { break }
                    if ($script:resizeTarget.Stage -ne 'splash') { break }   # shrunk - done here
                }
            }
            return
        }
        Write-Host 'Edge not found - falling back to default browser.' -ForegroundColor Yellow
    }
    Start-Process $url
}

function Assert-WindowBounds([IntPtr]$hwnd, $t, [int]$x, [int]$y, [int]$w, [int]$h) {
    # v2.0.17: one place decides how bounds are asserted, so the six enforcement
    # call sites cannot drift apart. A plain restore just places the window; a
    # maximised restore has to place the NORMAL rect and then maximise, so that
    # restore-down later returns to the geometry the user actually left.
    if ($t.ContainsKey('MaxFlag') -and $t.MaxFlag) {
        # v2.0.19 moved first and then maximised, guarded by IsZoomed so that a
        # repeat pass could not knock a maximised window back to normal. But
        # that guard also skipped the move when Edge ALREADY opened the adopted
        # window maximised - on Edge's own last-used monitor - which is the
        # intermittent "usually right, sometimes not, particularly if
        # maximised" field report (v2.0.22).
        #
        # So verify the OUTCOME rather than trusting a precondition: the window
        # is right only if it is maximised AND its bounds cover the centre of
        # the rect we intend (a maximised window spans its monitor's work area,
        # so containment of that centre IS the monitor test - no screen
        # enumeration needed). Anything else gets restored down (MoveWindow
        # cannot reposition a maximised window), placed, and maximised.
        #
        # This is self-correcting: a pass that races Edge is fixed by the next
        # one, and once the window is correct every later pass is a no-op, so
        # the enforcement never un-maximises or fights the user.
        $onTarget = $false
        if ([YVWin]::IsZoomed($hwnd)) {
            $rcNow = New-Object YVWin+RECT
            if ([YVWin]::GetWindowRect($hwnd, [ref]$rcNow)) {
                $tcx = $x + [int]($w / 2)
                $tcy = $y + [int]($h / 2)
                $onTarget = ($tcx -ge $rcNow.Left -and $tcx -lt $rcNow.Right -and
                             $tcy -ge $rcNow.Top  -and $tcy -lt $rcNow.Bottom)
            }
        }
        if (-not $onTarget) {
            if ([YVWin]::IsZoomed($hwnd)) { [YVWin]::ShowWindow($hwnd, 9) | Out-Null }   # SW_RESTORE
            [YVWin]::MoveWindow($hwnd, $x, $y, $w, $h, $true) | Out-Null
            [YVWin]::ShowWindow($hwnd, 3) | Out-Null
        }
    } else {
        [YVWin]::MoveWindow($hwnd, $x, $y, $w, $h, $true) | Out-Null
    }
}

function Invoke-WindowResize {
    # Window enforcement. Since v2.0.13 the splash and the year view share one
    # geometry, so the 'splash' stage no longer shrinks anything - it anchors
    # $script:appHwnd (which the watchdog and the expand step both rely on) and
    # re-asserts the intended bounds in case Edge adopted the window into an
    # existing process and ignored --window-size. Returns $true when finished
    # (done or expired).
    if (-not $script:resizeTarget) { return $true }
    $t = $script:resizeTarget
    if ((Get-Date) -gt $t.Deadline) { $script:resizeTarget = $null; return $true }
    try {
        # Primary capture: the window that appeared since launch (set-difference
        # against the pre-launch snapshot). Works title-less and pre-paint. A new
        # window with an unrelated non-empty title (the user opening a browser
        # window in the same second) is skipped, not shrunk.
        if ($t.Stage -eq 'splash' -and $script:appHwnd -eq [IntPtr]::Zero -and $t.ContainsKey('Pre')) {
            $nowWins = @([YVWin]::GetTopWindowsByClass('Chrome_WidgetWin_1'))
            foreach ($h in $nowWins) {
                if ($t.Pre -contains $h) { continue }
                $wt = [YVWin]::GetTitle($h)
                if ($wt -and -not ($wt -like '* YearView' -or $wt -like 'YearView loading*' -or
                                   $wt -like $script:titlePort -or $wt -like '127.0.0.1*')) { continue }
                $script:appHwnd = $h
                Assert-WindowBounds $h $t $t.SX $t.SY $t.SW $t.SH
                $t.Stage = 'sized'
                return $false
            }
        }
        # From here on, once the handle is anchored, drive it directly - no title
        # scans needed for the window we already own.
        if ($script:appHwnd -ne [IntPtr]::Zero -and [YVWin]::IsWindow($script:appHwnd)) {
            if ($t.Stage -eq 'splash') {
                Assert-WindowBounds $script:appHwnd $t $t.SX $t.SY $t.SW $t.SH
                $t.Stage = 'sized'
            }
            return $false
        }
        # Fallbacks (handle not yet anchored, e.g. non-App launch races): title scans.
        $winFull = Get-Process msedge -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle -like '* YearView' } |
            Select-Object -First 1
        if ($winFull -and $winFull.MainWindowHandle -ne [IntPtr]::Zero) {
            $script:appHwnd = $winFull.MainWindowHandle    # anchor for the watchdog
            Assert-WindowBounds $winFull.MainWindowHandle $t $t.BX $t.BY $t.BW $t.BH
            $script:resizeTarget = $null
            return $true
        }
        if ($t.Stage -eq 'splash') {
            $winSplash = Get-Process msedge -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowTitle -like 'YearView loading*' -or
                               $_.MainWindowTitle -like $script:titlePort } |
                Select-Object -First 1
            if ($winSplash -and $winSplash.MainWindowHandle -ne [IntPtr]::Zero) {
                $script:appHwnd = $winSplash.MainWindowHandle   # same top-level window survives navigation
                Assert-WindowBounds $winSplash.MainWindowHandle $t $t.SX $t.SY $t.SW $t.SH
                $t.Stage = 'sized'
            }
        }
    } catch {}
    return $false
}

function Invoke-ExpandNow {
    # Called just before delivering the full page: expand the (still splash-titled)
    # window to full size so the year view renders into an already-correct window.
    if (-not $script:resizeTarget) { return }
    $t = $script:resizeTarget
    try {
        if ($script:appHwnd -ne [IntPtr]::Zero -and [YVWin]::IsWindow($script:appHwnd)) {
            Assert-WindowBounds $script:appHwnd $t $t.BX $t.BY $t.BW $t.BH
            $script:resizeTarget = $null
            return
        }
        $win = Get-Process msedge -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle -like 'YearView loading*' -or
                           $_.MainWindowTitle -like $script:titlePort -or
                           $_.MainWindowTitle -like '* YearView' } |
            Select-Object -First 1
        if ($win -and $win.MainWindowHandle -ne [IntPtr]::Zero) {
            $script:appHwnd = $win.MainWindowHandle
            Assert-WindowBounds $win.MainWindowHandle $t $t.BX $t.BY $t.BW $t.BH
            $script:resizeTarget = $null
        }
    } catch {}
}

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
try {
    # v2.0.2: PORT TAKEOVER. If the bind fails, the likeliest cause is a stale
    # YearView bridge that outlived an upgrade still holding the port - the
    # window then talks to OLD code while the new install sits idle (field
    # evidence: a v2.0.0 bridge surviving into a v2.0.1 install, serving the
    # old behaviour and writing no logs). Identify our own siblings strictly -
    # PowerShell processes whose command line names this script by its full
    # path - stop them, and retry once. Anything else holding the port is
    # reported, never killed.
    try { $listener.Start() }
    catch {
        # First: is the incumbent the SAME version? Then this launch is a
        # legitimate second instance and the single-instance focus behaviour
        # (R-6, outer catch) must win - never kill a healthy current bridge.
        # A DIFFERENT version (or a dead/foreign non-answer plus a matching
        # sibling process) means a stale bridge: take the port.
        $incVer = ''
        try {
            $req = [Net.WebRequest]::Create("http://127.0.0.1:$Port/")
            $req.Timeout = 2000
            $resp = $req.GetResponse()
            $rd = New-Object IO.StreamReader($resp.GetResponseStream())
            $body = $rd.ReadToEnd(); $rd.Close(); $resp.Close()
            # v2.1.23: match the version ALONE. This scrape identifies an
            # incumbent bridge so a healthy same-version one is left alone
            # (R-6 focus) and only a different version is treated as stale.
            # It used to require the attribution words that followed the
            # number; with those removed from the stamp the match would have
            # failed, $incVer would have stayed empty, and a perfectly healthy
            # current bridge would have been judged stale and killed. The
            # looser pattern also still reads pages served by older versions,
            # whose stamp carried the old suffix.
            if ($body -match 'YearView version ([0-9][0-9.]*)') { $incVer = $Matches[1] }
        } catch {}
        if ($incVer -eq $script:yvVersion) { throw }   # same version: defer to R-6 focus
        Write-Host ("Port {0} is busy (incumbent version '{1}') - checking for a stale YearView bridge..." -f $Port, $incVer) -ForegroundColor Yellow
        $me = $PID
        $victims = @()
        try {
            $victims = @(Get-CimInstance Win32_Process -Filter "Name like 'powershell%'" -ErrorAction Stop |
                Where-Object { $_.ProcessId -ne $me -and $_.CommandLine -and
                               $_.CommandLine.IndexOf([string]$PSCommandPath,
                                   [StringComparison]::OrdinalIgnoreCase) -ge 0 })
        } catch {}
        if ($victims.Count) {
            foreach ($v in $victims) {
                Write-Host ("  stopping stale bridge PID {0}" -f $v.ProcessId) -ForegroundColor Yellow
                try { Stop-Process -Id $v.ProcessId -Force -ErrorAction Stop } catch {}
            }
            Start-Sleep -Milliseconds 1200
            $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
            $listener.Start()   # second failure falls to the outer catch below
        } else {
            throw   # port held by something that is not ours - surface it
        }
    }
    $script:listener = $listener   # shared with Invoke-QuickServe (in-build responsiveness)
} catch {
    # R-6: confirm the port is actually held by a YearView bridge before pointing
    # the browser at it - any other local app on this port would get our traffic.
    $isYV = $false
    try {
        $req = [Net.WebRequest]::Create("http://127.0.0.1:$Port/")
        $req.Timeout = 2000
        $resp = $req.GetResponse()
        $body = (New-Object IO.StreamReader($resp.GetResponseStream())).ReadToEnd()
        $resp.Close()
        if ($body -match 'YearView') { $isYV = $true }
    } catch {}
    if (-not $isYV) {
        throw ("Port {0} is in use by another application. Re-run with -Port <n> to pick a free port." -f $Port)
    }
    # A bridge is already serving this port. If its window is open, bring THAT to
    # the front rather than opening a duplicate; only open a new window when none
    # exists (a -NoOpen background bridge, or inside the ~10 s watchdog grace
    # after the window was closed). A second -NoOpen launch touches nothing.
    Write-Host ("Bridge already running on port {0}." -f $Port) -ForegroundColor Yellow
    if ($NoOpen) { return }
    $existingWin = [IntPtr]::Zero
    try {
        foreach ($h in @([YVWin]::GetTopWindowsByClass('Chrome_WidgetWin_1'))) {
            $wt = [YVWin]::GetTitle($h)
            if ($wt -like '* YearView' -or $wt -like 'YearView loading*' -or $wt -like $script:titlePort) {
                $existingWin = $h; break
            }
        }
    } catch {}
    if ($existingWin -ne [IntPtr]::Zero) {
        Write-Host 'Bringing the existing YearView window to the front.' -ForegroundColor Yellow
        Set-WindowForeground $existingWin
        return
    }
    Write-Host 'No open window found - opening the view against the running bridge.' -ForegroundColor Yellow
    # v2.1.25: load the remembered geometry HERE. This branch opens a window and
    # returns, and the normal Load-WindowRect call sits further down - AFTER the
    # block this returns from - so on this path $savedWinRect was still empty and
    # Open-View fell through to its computed default. That is the "restarted
    # within about ten seconds and got the default size" report: a restart that
    # quick meets the previous bridge still alive (since v2.1.20 the process
    # deliberately outlives its window to finish an update, and the window-gone
    # watchdog takes a few seconds more), so the launch takes THIS path instead
    # of the ordinary one. Restart later and the old bridge has gone, the port is
    # free, the ordinary path runs, and the size is right - which is exactly why
    # it looked intermittent.
    Load-StartupState
    Open-View
    while (-not (Invoke-WindowResize)) { Start-Sleep -Milliseconds 250 }
    return
}
Write-Host ''
Write-Host ("Year view running at {0}" -f $url) -ForegroundColor Green
Write-Host 'Click a category in the legend, then a day, to create a new Outlook item.' -ForegroundColor Green
Write-Host 'Keep this window open while using the view. Press Ctrl+C to stop.' -ForegroundColor Yellow
Load-StartupState
if (-not (Show-OutlookNeededPrompt)) {   # v2.1.10: YearView does not start without a windowed Outlook
    Write-Host 'Startup cancelled - Outlook is not running.' -ForegroundColor Yellow
    exit 0
}

# Change detection setup now lives in Ensure-Outlook and runs on the first
# data route (lazy, v1.2.17) - nothing here may touch Outlook, so the window
# below opens against the restored page with no COM wait. Restore stamps
# lastChangeTicks = now, forcing one silent revalidation of restored pages;
# nothing between here and the loop resets it.
Restore-PageCache

if (-not $NoOpen) { Open-View }

# Loop-local timers only: every script-wide variable is initialised in the single
# state block at the top of the script - do not add $script: inits here.
$nextResizeCheck = Get-Date
$nextWindowCheck = Get-Date
$nextLinkCheck  = Get-Date               # v2.1.0: release the COM link when Outlook closes
$nextEventSweep = Get-Date               # R-11
$windowGone = 0
while ($true) {
    while (-not $listener.Pending()) {
        # 25 ms PUMPING wait (never Start-Sleep here): WaitOne on this STA thread
        # lets the CLR service COM messages, which is the only way Outlook's
        # ItemChange/ItemAdd/ItemRemove events reach the PowerShell event queue.
        # Same latency budget as before: ~12 ms average to notice a request.
        [void]$script:comPump.WaitOne(25)
        if ($script:resizeTarget -and (Get-Date) -ge $nextResizeCheck) {
            [void](Invoke-WindowResize)
            $nextResizeCheck = (Get-Date).AddMilliseconds(300)
        }
        # If the app window is closed with its X (instead of Quit), shut the
        # bridge down too. Only in -App mode, where the window is ours alone.
        # Liveness is anchored to the window HANDLE, not the title: .NET's
        # MainWindowTitle needs a *visible* top-level window and reads blank while
        # the workstation is locked, the session is disconnected (RDP), or the
        # browser has put the page to sleep after a long idle. The old title-only
        # check therefore stopped a healthy bridge during idle, leaving the window
        # to show ERR_CONNECTION_REFUSED on the next click. IsWindow() stays true
        # through all of those states and only goes false when the window is truly
        # destroyed. The title scan remains for first capture and for re-adoption
        # when the window is closed and reopened from the shortcut.
        # v2.1.0: if Outlook is closing (its Quit event fired) or its process
        # is gone, release every COM reference we hold within ~2s - our event
        # connection points otherwise pin OUTLOOK.EXE in a half-shutdown
        # state that only End Task clears. Independent of -App: any serve
        # bridge holds the same references.
        if ($null -ne $script:closingAt -and (Get-Date) -ge $script:closingAt) {
            Close-YearViewOutlookGone      # v2.1.6: page has been told; now explain and exit
        }
        if ((Get-Date) -ge $nextLinkCheck) {
            $nextLinkCheck = (Get-Date).AddSeconds(2)
            if ($null -ne $script:outlook) {
                $quitSeen = $null -ne (Get-Event -SourceIdentifier 'YV_AppQuit' -ErrorAction SilentlyContinue)
                $procGone = -not (Get-Process -Name outlook -ErrorAction SilentlyContinue)
                # v2.1.5: closing Outlook's WINDOW does not always end
                # OUTLOOK.EXE - Outlook keeps running to finish a send/receive,
                # and in that state the process exists, the COM link still
                # answers, and Ensure-Outlook early-returns, so no connect
                # verdict is ever produced and nothing noticed the window had
                # gone. Ask WINDOWS for Outlook's frames instead: it is the
                # number of open Outlook windows, so zero means the user has
                # closed Outlook whatever the process table says. It is also
                # correct for "hide when minimised" (a tray-hidden Outlook
                # still HAS an explorer), which a window-handle test is not.
                # Two consecutive checks (~4s) guard against reading zero
                # during a momentary window transition.
                # v2.1.11: ask WINDOWS, not Outlook. The Explorers.Count read
                # used here blocks against an Outlook that is closing while a
                # sync is still running - Outlook is busy, so even a property
                # read queues - and a blocked loop can detect nothing (the
                # lengthy "Refreshing" toast in the field report). The Win32
                # frame scan cannot block, and counting hidden frames keeps
                # tray-minimised Outlook out of the trigger.
                $noWindows = $false
                if (-not $procGone -and -not (Test-OutlookWindowed)) {
                    $script:noExplorerTicks++
                    $noWindows = ($script:noExplorerTicks -ge 2)
                } else { $script:noExplorerTicks = 0 }
                if ($quitSeen -or $procGone -or $noWindows) {
                    # v2.1.2: releasing lets a closing OUTLOOK.EXE finish exiting;
                    # the session then ends with an explanation rather than
                    # limping on through a reconnect that has never been reliable.
                    # v2.1.6: do NOT block on the dialog yet. The page is still
                    # live and polling, and the moment this loop stops serving
                    # it concludes the bridge has died and says so - which is
                    # the stray "Outlook bridge not reachable" toast appearing
                    # seconds before the real explanation. Arm the closure
                    # instead and keep serving for a beat, so the next poll
                    # (2s cadence) collects the closing notice and the page
                    # falls silent; the dialog follows.
                    if ($null -eq $script:closingAt) {
                        Write-Host 'Outlook is closing - telling the page, then closing YearView.' -ForegroundColor Yellow
                        $script:closingAt = (Get-Date).AddSeconds(3)
                        # v2.1.11: NO release here any more. The trigger is now
                        # a Win32 scan, which proves nothing about whether
                        # Outlook can answer COM calls - and releasing against
                        # a sync-busy closing Outlook blocks exactly like the
                        # read this release removed. The process exit drops
                        # every reference; DCOM's dead-client sweep then unpins
                        # Outlook within a few minutes and it finishes its sync
                        # and exits by itself - minutes, where the wedge was
                        # forever.
                    }
                }
            }
        }
        if ($App -and (Get-Date) -ge $nextWindowCheck) {
            $nextWindowCheck = (Get-Date).AddSeconds(1)   # v2.1.21: sample geometry more promptly
            $alive = $false
            try {
                if ($script:appHwnd -ne [IntPtr]::Zero -and [YVWin]::IsWindow($script:appHwnd)) {
                    # R-13: HWNDs are recycled by the OS. Confirm the handle still
                    # belongs to an msedge process before trusting it; if it was
                    # recycled to something else, drop the anchor and recapture.
                    $wpid = [uint32]0
                    [void][YVWin]::GetWindowThreadProcessId($script:appHwnd, [ref]$wpid)
                    $owner = $null
                    if ($wpid) { $owner = Get-Process -Id ([int]$wpid) -ErrorAction SilentlyContinue }
                    if ($owner -and $owner.Name -eq 'msedge') { $alive = $true }
                    else { $script:appHwnd = [IntPtr]::Zero }
                }
                if (-not $alive) {
                    $present = Get-Process msedge -ErrorAction SilentlyContinue |
                        Where-Object { $_.MainWindowTitle -like '* YearView' -or
                                       $_.MainWindowTitle -like 'YearView loading*' -or
                                       $_.MainWindowTitle -like $script:titlePort } |
                        Select-Object -First 1
                    if ($present) {
                        $script:appHwnd = $present.MainWindowHandle
                        $alive = $true
                    }
                }
            } catch { $alive = $null }   # inconclusive - never count as a miss
            if ($alive) {
                $script:windowSeen = $true
                $windowGone = 0
                # v2.0.16: sample the window geometry the user has settled on.
                # Skipped while launch enforcement is still running (its
                # transition rects are not user choices) and while minimised
                # (bogus -32000 coordinates). v2.0.17: a maximised window is
                # RECORDED as such - the saved w,h,x,y stay at the last known
                # normal rect (what restore-down returns to) with the max flag
                # set, so the next launch restores maximised.
                if ($null -eq $script:resizeTarget -and
                    $script:appHwnd -ne [IntPtr]::Zero -and
                    -not [YVWin]::IsIconic($script:appHwnd)) {
                    if ([YVWin]::IsZoomed($script:appHwnd)) {
                        if ($script:lastNormalRect) {
                            # v2.0.18: the MONITOR matters while maximised. The
                            # maximised bounds are valid and identify the
                            # monitor; if the remembered normal rect sits on a
                            # different one (Win+Shift+Arrow moves a maximised
                            # window without ever un-maximising), translate the
                            # normal rect onto this monitor - size kept,
                            # centred, clamped - so the next launch maximises
                            # HERE and restore-down stays sensible.
                            try {
                                $rcM = New-Object YVWin+RECT
                                if ([YVWin]::GetWindowRect($script:appHwnd, [ref]$rcM)) {
                                    $mcx = [int](($rcM.Left + $rcM.Right) / 2)
                                    $mcy = [int](($rcM.Top + $rcM.Bottom) / 2)
                                    $ncx = [int]$script:lastNormalRect[2] + [int]([int]$script:lastNormalRect[0] / 2)
                                    $ncy = [int]$script:lastNormalRect[3] + [int]([int]$script:lastNormalRect[1] / 2)
                                    $curWa = $null; $sameWa = $false
                                    foreach ($scr in [System.Windows.Forms.Screen]::AllScreens) {
                                        $wa = $scr.WorkingArea
                                        $hasMax = ($mcx -ge $wa.X -and $mcx -lt ($wa.X + $wa.Width) -and
                                                   $mcy -ge $wa.Y -and $mcy -lt ($wa.Y + $wa.Height))
                                        if ($hasMax) {
                                            $curWa = $wa
                                            $sameWa = ($ncx -ge $wa.X -and $ncx -lt ($wa.X + $wa.Width) -and
                                                       $ncy -ge $wa.Y -and $ncy -lt ($wa.Y + $wa.Height))
                                            break
                                        }
                                    }
                                    if ($curWa -and -not $sameWa) {
                                        $nw = [Math]::Min([int]$script:lastNormalRect[0], $curWa.Width)
                                        $nh = [Math]::Min([int]$script:lastNormalRect[1], $curWa.Height)
                                        $nx = $curWa.X + [int](($curWa.Width  - $nw) / 2)
                                        $ny = $curWa.Y + [int](($curWa.Height - $nh) / 2)
                                        $script:lastNormalRect = @($nw, $nh, $nx, $ny)
                                    }
                                }
                            } catch {}
                            Save-WindowRect ([int]$script:lastNormalRect[0]) ([int]$script:lastNormalRect[1]) `
                                ([int]$script:lastNormalRect[2]) ([int]$script:lastNormalRect[3]) $true
                        }
                    } else {
                        $rc = New-Object YVWin+RECT
                        if ([YVWin]::GetWindowRect($script:appHwnd, [ref]$rc)) {
                            Save-WindowRect ($rc.Right - $rc.Left) ($rc.Bottom - $rc.Top) $rc.Left $rc.Top $false
                        }
                    }
                }
            } elseif (($null -ne $alive) -and $script:windowSeen) {
                $windowGone++
                if ($windowGone -ge 5) {
                    Write-Host 'YearView window closed - stopping bridge.' -ForegroundColor Yellow
                    $script:quitRequested = $true
                }
            }
        }
        # R-11: a headless bridge (-Startup -NoOpen) may have no page polling
        # /dirty, so queued COM events would otherwise accumulate indefinitely.
        # Sweep every 30 s; found events mark the cache stale exactly as a poll
        # would, so the next page request rebuilds.
        if ((Get-Date) -ge $nextEventSweep) {
            $nextEventSweep = (Get-Date).AddSeconds(30)
            try {
                $evts = @(Get-Event -ErrorAction SilentlyContinue |
                    Where-Object { $_.SourceIdentifier -like 'YV_Item*' })
                if ($evts.Count) {
                    $evts | Remove-Event -ErrorAction SilentlyContinue
                    $script:lastChangeTicks = (Get-Date).Ticks   # stale-revalidate (v1.2.25)
                }
            } catch {}
        }
        if ($script:quitRequested) { break }
    }
    if ($script:quitRequested) { break }
    $client = $listener.AcceptTcpClient()
    try {
        $script:respondedTo = $null
        $stream = $client.GetStream()
        $stream.ReadTimeout = 2000
        $buf = New-Object byte[] 16384
        $total = 0
        $reqText = ''
        # R-2: ReadTimeout only bounds a single Read. A client trickling one byte
        # per call could otherwise hold this single-threaded loop - and with it the
        # UI, polling and the window watchdog - open for hours. Overall budget:
        $connDeadline = (Get-Date).AddSeconds(5)
        while ($total -lt $buf.Length) {
            if ((Get-Date) -gt $connDeadline) { $reqText = ''; break }
            $read = $stream.Read($buf, $total, $buf.Length - $total)
            if ($read -le 0) { break }
            $total += $read
            $reqText = [Text.Encoding]::ASCII.GetString($buf, 0, $total)
            # v2.0.4: requests may carry a small Content-Length body; the loop
            # reads on until it arrives (absent or zero means headers-only).
            # Kept as general protocol support (v2.0.6) - no current route
            # sends a body, but the capability is correct and harmless.
            $hdrEnd = $reqText.IndexOf("`r`n`r`n")
            if ($hdrEnd -ge 0) {
                $bodyLen = 0
                if ($reqText.Substring(0, $hdrEnd) -match '(?im)^Content-Length:\s*(\d+)\s*$') {
                    $bodyLen = [int]$Matches[1]
                }
                if ($bodyLen -gt 8192) { $reqText = ''; break }   # never a legitimate YearView request
                if ($total -ge ($hdrEnd + 4 + $bodyLen)) { break }
            }
        }
        if (-not $reqText) { continue }
        $reqBody = ''
        $hdrEnd = $reqText.IndexOf("`r`n`r`n")
        if ($hdrEnd -ge 0 -and ($hdrEnd + 4) -lt $reqText.Length) {
            $reqBody = $reqText.Substring($hdrEnd + 4)
        }
        $lines = $reqText -split "`r`n"
        $parts = $lines[0] -split ' '
        if ($parts.Count -lt 2) { continue }
        $method = $parts[0].ToUpperInvariant()
        $path = $parts[1]

        # Host-header check: defeats DNS-rebinding (rebound requests carry a foreign host)
        $hostHdr = ($lines | Where-Object { $_ -match '^(?i)Host:\s*(.+)$' } |
            ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
        if ($hostHdr -and $hostHdr -notmatch '^(127\.0\.0\.1|localhost)(:\d+)?$') {
            try {
                $logP = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
                Add-Content -Path $logP -Value ("[{0}] request rejected ({1}): {2} {3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), 'host gate', $method, (($path -split '\?')[0]))
            } catch {}
            Send-Http $stream '403 Forbidden' 'text/plain' (JsonBytes 'forbidden')
            continue
        }

        # S-1: browsers label requests started by another site 'cross-site'.
        # Legitimate traffic is 'none' (address bar / shortcut) or 'same-origin'
        # (our own page's fetches and navigations). Browsers that predate the
        # header omit it; the token check below still covers those.
        $fetchSite = ($lines | Where-Object { $_ -match '^(?i)Sec-Fetch-Site:\s*(.+)$' } |
            ForEach-Object { $Matches[1].Trim().ToLowerInvariant() } | Select-Object -First 1)
        if ($fetchSite -and $fetchSite -ne 'none' -and $fetchSite -ne 'same-origin') {
            try {
                $logP = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
                Add-Content -Path $logP -Value ("[{0}] request rejected ({1}): {2} {3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), 'sec-fetch gate', $method, (($path -split '\?')[0]))
            } catch {}
            Send-Http $stream '403 Forbidden' 'text/plain' (JsonBytes 'forbidden')
            continue
        }

        # S-2/S-3: the session token travels in a request header, never in a URL
        # (so never in history or the address bar), and is compared exactly.
        # Every non-page route - including /build, the only way to start a
        # calendar read (S-1) - requires POST plus the token. Page routes only
        # serve from cache; a cache miss returns the splash, which owns the token
        # and performs the authenticated build.
        $tokHdr = ($lines | Where-Object { $_ -match '^(?i)X-YV-Token:\s*(.+)$' } |
            ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
        $authed = ($tokHdr -is [string] -and $tokHdr -ceq $script:sessionToken)

        $isPage = ($path -eq '/' -or $path.StartsWith('/?') -or $path.StartsWith('/index.html'))
        if (-not $isPage) {
            if ($method -ne 'POST' -or -not $authed) {
                try {
                    $logP = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
                    Add-Content -Path $logP -Value ("[{0}] request rejected ({1}): {2} {3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), 'token gate', $method, (($path -split '\?')[0]))
                } catch {}
                Send-Http $stream '403 Forbidden' 'application/json' (JsonBytes '{"ok":false,"err":"bad token"}')
                continue
            }
        }
        # Routes below may touch Outlook COM - connect lazily/reconnect. The
        # page, favicon and quit routes never need Outlook, so serving a
        # restored page at startup cannot block on connecting; the first data
        # route (normally the fresh page's own /dirty poll, ~2 s in) pays the
        # one-off connect cost in the background instead.
        # v2.1.6: once the closure is armed, every route answers the same
        # notice. The page uses it to stop polling and suppress its own error
        # toasts, so the only thing the user sees is the dialog.
        # v2.1.9: EXCEPT page routes. A navigation that lands here while armed
        # (the build path navigates on any non-retry 200) had the browser
        # render the JSON notice AS the page - the field-reported raw
        # {"ok":false,"closing":true,...} display. Pages serve normally; the
        # freshly loaded page's first poll collects the notice and goes quiet.
        if ($null -ne $script:closingAt -and
            -not ($path -eq '/' -or $path -eq '/index.html' -or
                  $path.StartsWith('/?') -or $path.StartsWith('/favicon'))) {
            Send-Http $stream '200 OK' 'application/json' `
                (JsonBytes '{"ok":false,"closing":true,"dirty":false,"err":"Outlook is not running - YearView is closing"}')
            continue
        }
        # v2.0.28: the dirty poll is BACKGROUND traffic and must never launch,
        # attach or reconnect Outlook. It used to: with Outlook absent, the 2s
        # poll - not the user's action - triggered the launch, and closing
        # Outlook deliberately resurrected it (2s poll vs 10s storm guard).
        # Worse, once the launched process existed, the next poll paid the
        # attach and could wedge the loop mid-boot. With no LIVE link the poll
        # now answers "nothing changed" and stands down; the next USER action
        # connects, after which polling resumes on the fresh link.
        if ($path.StartsWith('/dirty')) {
            $dLive = $false
            try { if ($script:outlook) { [void]$script:outlook.Version; $dLive = $true } } catch {}
            if (-not $dLive) {
                Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"dirty":false}')
                continue
            }
        }
        if (-not ($path -eq '/' -or $path -eq '/index.html' -or $path.StartsWith('/?') -or
                  $path.StartsWith('/favicon') -or $path.StartsWith('/quit') -or
                  $path.StartsWith('/cached'))) {
            # v2.1.8: CHECK BEFORE READING. A user action arriving just after
            # the Outlook window closes used to start a long COM enumeration
            # against an Outlook that is shutting down. Those calls hang, and
            # because this loop is single-threaded the 2s watchdog cannot run
            # while it is blocked - so nothing detected anything, the progress
            # toast sat there, and only killing OUTLOOK.EXE (which makes the
            # hung calls fail and releases the loop) produced the dialog. The
            # detection was never wrong; it simply never got to run.
            #
            # So ask before committing to the read. A Win32 frame scan is one
            # cheap property, and zero means the user has closed Outlook.
            # A single reading is decisive HERE - unlike the background poll,
            # which keeps its two-tick guard below - because a user action
            # arriving during a momentary window transition is far rarer than
            # a poll landing in one, and arming still gives the page three
            # seconds to be told before anything closes.
            #
            # /dirty is excluded deliberately: it is background traffic and
            # must not end a session on one reading.
            if ($null -ne $script:outlook -and -not $path.StartsWith('/dirty')) {
                # v2.1.11: Win32 scan, not the COM read - see the watchdog
                # note. This is the check that answers a Refresh pressed just
                # after the Outlook window is closed mid-sync.
                $winsGone = -not (Test-OutlookWindowed)
                if ($winsGone) {
                    if ($null -eq $script:closingAt) {
                        Write-Host 'Outlook window closed - ending the session before starting a read.' -ForegroundColor Yellow
                        $script:closingAt = (Get-Date).AddSeconds(3)
                        # v2.1.11: no release here - the Win32 trigger proves
                        # nothing about COM responsiveness, and releasing
                        # against a sync-busy closing Outlook blocks. The
                        # process exit unpins it (DCOM sweep, a few minutes).
                    }
                    Send-Http $stream '200 OK' 'application/json' `
                        (JsonBytes '{"ok":false,"closing":true,"dirty":false,"err":"YearView is closing"}')
                    continue
                }
            }
            # v2.0.24: connecting can fail or be interrupted - a cold start that
            # errors, a profile prompt, an Outlook already shutting down. This
            # call sits OUTSIDE every per-route handler, so an exception here
            # escaped to the outer catch, which closed the connection without
            # writing anything: the browser saw a network failure and the client
            # could only report "Bridge not reachable" - blaming the bridge for
            # an Outlook fault. Answer with a retryable JSON fault instead.
            try { Ensure-Outlook }
            catch {
                $starting = ($_.Exception.Message -match 'YV_OUTLOOK_STARTING')
                $absent   = ($_.Exception.Message -match 'YV_OUTLOOK_ABSENT')
                $stuck    = ($_.Exception.Message -match 'YV_OUTLOOK_STUCK')
                if (-not $starting -and -not $absent -and -not $stuck) {
                    Write-Host ('Outlook connect failed: ' + $_.Exception.Message) -ForegroundColor Yellow
                    try {
                        # v2.0.25: the console is hidden in production - a connect
                        # failure must land somewhere a field report can find it.
                        $logP = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
                        Add-Content -Path $logP -Value ("[{0}] Outlook connect failed on {1}: {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), (($path -split '\?')[0]), $_.Exception.Message)
                    } catch {}
                }
                Release-OutlookLink   # v2.1.0: no-op when nothing is held
                if (-not $starting) {
                    # v2.1.7: this verdict ends the session, so do NOT send the
                    # "please start Outlook" instruction - it contradicts the
                    # dialog that follows a moment later and reads as two
                    # different answers to the same question. Arm the closure
                    # and answer the same closing notice every armed route
                    # sends, so the page simply falls silent and the dialog is
                    # the only thing the user is asked to read.
                    if ($null -eq $script:closingAt) { $script:closingAt = (Get-Date).AddSeconds(3) }
                    Send-Http $stream '200 OK' 'application/json' `
                        (JsonBytes '{"ok":false,"closing":true,"dirty":false,"err":"YearView is closing"}')
                    continue
                }
                Send-Http $stream '200 OK' 'application/json' `
                    (JsonBytes ((@{ ok = $false; retry = $true
                                    err = 'Outlook is starting' } | ConvertTo-Json -Compress)))
                # v2.1.3: UNCONDITIONAL. v2.1.2 required a link to have existed
                # earlier in the session, which is why pressing Refresh with
                # Outlook closed sat in a retry loop instead of ending the
                # session. The rule is now exactly the one the user stated:
                # Outlook is not running (or not usable) => YearView closes and
                # says so. 'starting' remains excluded - a booting Outlook has a
                # process and resolves by waiting, which is what the user is
                # doing when they start Outlook after the launch dialog.
                continue
            }
        }

        if ($path -eq '/' -or $path -eq '/index.html') {
            # The ENTRY url. v1.2.16's restored cache never got served here -
            # this branch was hard-wired to the splash from the S-1 design, so
            # every launch still paid splash + build. Serve the cached page for
            # the default year when we have one (restored from disk or built
            # this session); the splash remains the no-cache bootstrap.
            if ($script:pages.ContainsKey($Year)) {
                Invoke-ExpandNow
                Send-Http $stream '200 OK' 'text/html; charset=utf-8' $script:pages[$Year]
            } else {
                Send-Http $stream '200 OK' 'text/html; charset=utf-8' $splashBytes
            }
        }
        elseif ($path.StartsWith('/?') -or $path.StartsWith('/index.html?')) {
            $y = Get-QueryYear $path $Year
            if ($script:pages.ContainsKey($y)) {
                Invoke-ExpandNow
                Send-Http $stream '200 OK' 'text/html; charset=utf-8' $script:pages[$y]
            } else {
                # S-1: a page request never builds. The splash carries the token,
                # POSTs /build, then navigates back here to the cached result.
                Send-Http $stream '200 OK' 'text/html; charset=utf-8' (Get-SplashBytes $y)
            }
        }
        elseif ($path.StartsWith('/setsetting')) {
            # v2.1.29: POST + token enforced by the gate above.
            # v2.1.30: read the BODY, not the query string. The client's api()
            # helper sends its second argument as a form-encoded POST body -
            # which is what /setfilter above has always parsed - but this route
            # was written to parse $path, where nothing ever arrives. The value
            # was therefore always absent, the route always threw, and the page
            # always reported that it could not save the setting.
            try {
                # v3.0.0: one parser for every setting, so adding another is a
                # table entry rather than a new branch that can drift.
                $applied = $false
                $smChanged = $false
                foreach ($k in @('confirmdelete', 'scrollyears', 'printlegend')) {
                    if ($reqBody -match ('(?:^|&)' + $k + '=([^&]*)')) {
                        $raw = $Matches[1]
                        try { $raw = [uri]::UnescapeDataString($raw.Replace('+', '%20')) } catch {}
                        $on = ([string]$raw -ne '0')
                        # v3.4.6: a switch, so a further flag cannot be sent to
                        # the wrong variable by a two-way if/else.
                        switch ($k) {
                            'confirmdelete' { $script:confirmDelete = $on }
                            'scrollyears'   { $script:scrollYears   = $on }
                            'printlegend'   { $script:printLegend   = $on }
                        }
                        $applied = $true
                    }
                }
                # v3.2.0: the one numeric setting. Range-checked here as well as
                # at load, so a hand-edited or malformed value can never reach a
                # page as a scale factor - it is REFUSED, which the client shows
                # by putting the control back where it was.
                if ($reqBody -match '(?:^|&)startmonth=([^&]*)') {
                    $sn = -99
                    if (-not [int]::TryParse($Matches[1].Trim(), [ref]$sn) -or $sn -lt -1 -or $sn -gt 11) {
                        throw 'Unknown setting.'
                    }
                    $smChanged = ($script:startMonth -ne $sn)
                    $script:startMonth = $sn
                    $applied = $true
                }
                if ($reqBody -match '(?:^|&)printpages=([^&]*)') {
                    $pn = -1
                    if (-not [int]::TryParse($Matches[1].Trim(), [ref]$pn) -or $pn -lt 0 -or $pn -gt 4) {
                        throw 'Unknown setting.'
                    }
                    $script:printPages = $pn
                    $applied = $true
                }
                # v3.4.7: the one STRING setting - which categories show no
                # titles. Set from the Filter panel, not the Settings panel,
                # but saved, re-stamped and refused exactly like the others.
                if ($reqBody -match '(?:^|&)notitles=([^&]*)') {
                    $nt = $null
                    try { $nt = [uri]::UnescapeDataString($Matches[1].Replace('+', '%20')) } catch { $nt = $null }
                    if (-not (Test-TitleCsv $nt)) { throw 'Unknown setting.' }
                    $script:noTitles = [string]$nt
                    $applied = $true
                }
                if (-not $applied) { throw 'Unknown setting.' }
                Save-Settings
                # v3.1.8: RE-STAMP the pages already built, exactly as
                # /setfilter has done for the filter seed since v2.1.22.
                # Without this the file was correct and the pages were not:
                # a page built before the change still carried the old value,
                # so a background update that navigated to it put the setting
                # back - which is what "it did not stick" looked like. The
                # window is widest during a build, because the save waits the
                # build out while the build is baking the OLD value into the
                # very page it is about to serve.
                try {
                    foreach ($py in @($script:pages.Keys)) {
                        $h0 = [Text.Encoding]::UTF8.GetString($script:pages[$py])
                        $h1 = Update-SettingLiterals $h0
                        if ($h1 -ne $h0) {
                            $script:pages[$py] = [Text.Encoding]::UTF8.GetBytes($h1)
                            Save-PageCache $py
                        }
                    }
                } catch {}
                # v3.4.3: a page built before this setting changed carries the
                # wrong spill payload - usually none at all - so the first load
                # after the change had to fetch it and only a later rebuild
                # embedded it, which is why it took more than one load to go
                # instant. Patch the island in place instead. Only on a real
                # change of this setting: it can cost a calendar read, and the
                # other settings have no business paying for one.
                if ($smChanged) {
                    try {
                        $cy = (Get-Date).Year
                        $eff = $(if ($script:startMonth -eq -1) { (Get-Date).Month - 1 } else { [int]$script:startMonth })
                        $spill = 'null'
                        if ($eff -gt 0 -and $eff -le 11) {
                            $nc = $script:payloads[$cy + 1]
                            if ($nc -and [long]$nc.Ticks -ge [long]$script:lastChangeTicks) {
                                $spill = [string]$nc.Json
                            } else {
                                $yd2 = Get-YearData -Y ($cy + 1)   # hash discarded: not a page build
                                $spill = $yd2.Json
                                $script:payloads[$cy + 1] = @{ Json = $yd2.Json; Hash = $yd2.Hash; Ticks = (Get-Date).Ticks }
                            }
                        }
                        if ($script:pages.ContainsKey($cy)) {
                            $h0 = [Text.Encoding]::UTF8.GetString($script:pages[$cy])
                            $h1 = Set-SpillPayload $h0 $spill
                            if ($h1 -ne $h0) {
                                $script:pages[$cy] = [Text.Encoding]::UTF8.GetBytes($h1)
                                Save-PageCache $cy
                            }
                        }
                    } catch {
                        Write-Warning ('Could not re-stamp the spill months: {0}' -f $_.Exception.Message)
                    }
                }
                Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"ok":true}')
            } catch {
                Send-Http $stream '200 OK' 'application/json' (JsonError $_.Exception.Message)
            }
        }
        elseif ($path.StartsWith('/setfilter')) {
            # v2.0.8: persist the hidden-category filter. The body carries
            # "hide=<url-encoded csv>"; an empty value is a real "show all".
            # POST + token gated by the global gate; no Outlook involved.
            $csv = ''
            if ($reqBody -match '(?:^|&)hide=([^&]*)') {
                try { $csv = [uri]::UnescapeDataString($Matches[1].Replace('+', '%20')) }
                catch { $csv = $Matches[1] }
            }
            Save-Filter $csv
            # v2.1.22: RE-STAMP the cached pages. The filter seed is baked into
            # each page's bytes at BUILD time, and the cache is served verbatim
            # on the next start - so a filter changed after the last build was
            # persisted correctly in yv-filter.cfg but the reopened page still
            # carried the OLD seed. The startup revalidation could not rescue
            # it: its "changed" test compares calendar CONTENT only, so a
            # rebuild differing solely in the filter reported no change and
            # the stale page stayed. (Refresh forces a navigation to the fresh
            # page, which is why Refresh made a hide stick.) Substituting the
            # new literal for the old one in every cached page - the same
            # in-memory patch Restore-PageCache already does for the session
            # token - keeps cache and filter permanently in step, whichever
            # order the user changes and closes in.
            try {
                $newLit = ConvertTo-SeedStr ([string]$script:hiddenCsv)
                foreach ($py in @($script:pages.Keys)) {
                    $oldLit = [string]$script:pageHideLits[$py]
                    if (-not $oldLit -or $oldLit -eq $newLit) { continue }
                    $html = [Text.Encoding]::UTF8.GetString($script:pages[$py])
                    $marker = "`n  const HIDE_SEED = "   # v3.4.9: line-anchored
                    $at = $html.IndexOf($marker + $oldLit + ';')
                    if ($at -lt 0) { continue }
                    $html = $html.Substring(0, $at + $marker.Length) + $newLit + $html.Substring($at + $marker.Length + $oldLit.Length)
                    $script:pages[$py] = [Text.Encoding]::UTF8.GetBytes($html)
                    $script:pageHideLits[$py] = $newLit
                    Save-PageCache $py
                }
            } catch {}
            Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"ok":true}')
        }
        elseif ($path.StartsWith('/calendars')) {
            # v2.0.5: this ONE route both lists and switches. Listing caches
            # the enumeration server-side; the client then switches with a
            # tiny "?set=N" index into that cache - after two field failures
            # ("Failed to fetch") where only the id-carrying request died, no
            # folder identifier travels on the wire in either direction's
            # switch path, and the switch reuses the request shape that
            # provably works. Arrival is logged so "did it reach the bridge"
            # is a recorded fact.
            $setIdx = -1
            if ($path -match '[?&]set=(\d{1,3})') { $setIdx = [int]$Matches[1] }
            if ($setIdx -ge 0) {
                try {
                    $logP = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
                    Add-Content -Path $logP -Value ("[{0}] switch request arrived: set={1} (cached list: {2})" -f `
                        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $setIdx, @($script:calList).Count)
                } catch {}
                $okSet = $false; $nm = ''; $scStage = 'lookup'; $scErr = ''
                try {
                    $entry = $null
                    if ($setIdx -lt @($script:calList).Count) { $entry = $script:calList[$setIdx] }
                    if ($entry) {
                        $scStage = 'resolve'
                        $f = $script:ns.GetFolderFromID($entry.id, $entry.sid)
                        $scStage = 'typecheck'
                        if ($f -and $f.DefaultItemType -eq 1) {
                            $scStage = 'apply'
                            $nm = [string]$f.Name
                            if ($entry.id -eq [string]$script:ns.GetDefaultFolder(9).EntryID) {
                                $script:selCalId = ''; $script:selCalStoreId = ''; $script:selCalName = ''
                            } else {
                                $script:selCalId = $entry.id; $script:selCalStoreId = $entry.sid; $script:selCalName = $nm
                            }
                            Save-CalSelection
                            Get-EventSubscriber -ErrorAction SilentlyContinue |
                                Where-Object { $_.SourceIdentifier -like 'YV_*' } |
                                Unregister-Event -ErrorAction SilentlyContinue
                            $script:watchItems = $null
                            Ensure-Outlook
                            $script:pages.Clear()
                            $script:payloads.Clear()   # v3.4.1: another calendar entirely
                            $script:contentHashes.Clear()
                            $script:lastChangeTicks = (Get-Date).Ticks
                            Write-Host ("  ~ Calendar switched to '{0}'" -f $nm) -ForegroundColor DarkCyan
                            $okSet = $true
                        } else { $scErr = 'not a calendar folder' }
                    } else { $scErr = ('no cached entry at index {0} - reopen the picker' -f $setIdx) }
                } catch { $scErr = $_.Exception.Message }
                if (-not $okSet) {
                    $reason = ('calendar switch failed at stage [{0}]: {1}' -f $scStage, $scErr)
                    try {
                        $logP = Join-Path (Split-Path -Parent $PSCommandPath) 'yv-error.log'
                        Add-Content -Path $logP -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $reason)
                    } catch {}
                    Write-Host ("  ! {0}" -f $reason) -ForegroundColor Yellow
                }
                Send-Http $stream '200 OK' 'application/json' `
                    (JsonBytes ('{"ok":' + $(if ($okSet) {'true'} else {'false'}) + ',"name":' + (ConvertTo-JsonStr $nm) +
                        ',"err":' + (ConvertTo-JsonStr $(if ($okSet) {''} else {('[' + $scStage + '] ' + $scErr)})) + '}'))
            } else {
                $entries = New-Object System.Collections.Generic.List[string]
                $cache = New-Object System.Collections.Generic.List[object]
                try {
                    $defId = [string]$ns.GetDefaultFolder(9).EntryID
                    $stack = New-Object 'System.Collections.Generic.Stack[object]'
                    $stack.Push(@($ns.DefaultStore.GetRootFolder(), 0))
                    while ($stack.Count -gt 0 -and $entries.Count -lt 40) {
                        $fr = $stack.Pop(); $depth = [int]$fr[1]
                        foreach ($sub in @($fr[0].Folders)) {
                            try {
                                if ($sub.DefaultItemType -eq 1) {
                                    $sid = [string]$sub.EntryID
                                    $isDef = ($sid -eq $defId)
                                    $isSel = $(if ($script:selCalId) { $sid -eq $script:selCalId } else { $isDef })
                                    $entries.Add('{"i":' + $cache.Count +
                                        ',"name":' + (ConvertTo-JsonStr ([string]$sub.Name)) +
                                        ',"def":' + $(if ($isDef) {'true'} else {'false'}) +
                                        ',"sel":' + $(if ($isSel) {'true'} else {'false'}) + '}')
                                    $cache.Add([pscustomobject]@{ id = $sid; sid = [string]$sub.StoreID })
                                }
                                if ($depth -lt 3) { $stack.Push(@($sub, ($depth + 1))) }
                            } catch {}
                        }
                    }
                } catch {}
                $script:calList = $cache.ToArray()
                Send-Http $stream '200 OK' 'application/json' `
                    (JsonBytes ('{"ok":true,"calendars":[' + ($entries -join ',') + ']}'))
            }
        }
        elseif ($path.StartsWith('/cached')) {
            # Cache probe for instant year navigation: reports whether a year's
            # page is already in memory. Touches nothing but the page table -
            # no Outlook, no build - so it answers in ~milliseconds.
            $y = Get-QueryYear $path $Year
            Send-Http $stream '200 OK' 'application/json' `
                (JsonBytes ('{"cached":' + $script:pages.ContainsKey($y).ToString().ToLower() + '}'))
        }
        elseif ($path.StartsWith('/build')) {
            # POST + token enforced by the gate above; the only route that reads
            # the calendar. Replaces the old GET /refresh.
            $y = Get-QueryYear $path $Year
            # The build blocks this loop, so settle the window geometry NOW,
            # before starting it - otherwise the enforcement (and the handle
            # anchoring it performs) lands only after the build completes.
            $tries = 0
            while ($script:resizeTarget -and $script:resizeTarget.Stage -eq 'splash' -and $tries -lt 10) {
                [void](Invoke-WindowResize)
                if ($script:resizeTarget -and $script:resizeTarget.Stage -eq 'splash') {
                    Start-Sleep -Milliseconds 150
                }
                $tries++
            }
            # Our own /new and /delete raise ItemAdd/ItemRemove like any external
            # change. If those events are still queued after this build, the next
            # /dirty poll "discovers" them, clears the cache, and the page request
            # that follows flashes the splash. Drain the queue, build, then
            # re-baseline the count and modification fingerprint so the bridge never
            # treats its own edits as news. Changes arriving DURING the build stay
            # queued and simply trigger one further seamless refresh.
            Get-Event -ErrorAction SilentlyContinue |
                Where-Object { $_.SourceIdentifier -like 'YV_Item*' } |
                Remove-Event -ErrorAction SilentlyContinue
            # Baseline BEFORE the build (equally sound: changes landing mid-build
            # stay queued and trigger one poll refresh either way) so its duration
            # is known when the payload is serialised. This sort walks the WHOLE
            # calendar, not just the year - a prime suspect for slow builds, now
            # measured per build and shown in the page stamp.
            $bsw = [Diagnostics.Stopwatch]::StartNew()
            try {
                $script:lastCount = [int]$script:watchItems.Count
                $script:watchItems.Sort('[LastModificationTime]', $true)
                $newest = $script:watchItems.GetFirst()
                if ($newest) { $script:lastModTicks = $newest.LastModificationTime.Ticks }
            } catch {}
            $script:lastBaselineMs = $bsw.ElapsedMilliseconds
            $script:pages[$y] = [Text.Encoding]::UTF8.GetBytes((Build-YearViewHtml -Y $y))
            # changed = the CONTENT hash moved (timestamps excluded); lets the
            # client skip navigation for rebuilds that alter nothing visible
            $prevHash = [string]$script:contentHashes[$y]
            $script:contentHashes[$y] = $script:lastBuildHash
            # v3.4.5: and the spill months, when this view has any. They are on
            # screen, so a change to them is a change to what the reader sees -
            # but the year-on-screen hash cannot notice it, which is why an edit
            # after December announced an update and then did nothing.
            $prevSpill = [string]$script:spillHashes[$y]
            $script:spillHashes[$y] = $script:lastSpillHash
            $same = ($prevHash -and $prevHash -eq $script:lastBuildHash -and
                     $prevSpill -eq $script:lastSpillHash)
            $chg = if ($same) { 'false' } else { 'true' }
            Save-PageCache $y   # instant serving on the next bridge start
            Send-Http $stream '200 OK' 'application/json' `
                (JsonBytes ('{"ok":true,"changed":' + $chg + ',"ts":"' + [string](Get-Date).Ticks + '"}'))
        }
        elseif ($path.StartsWith('/data')) {
            # v3.0.1: a year's payload WITHOUT a page around it - the same JSON
            # the page carries in __DATA__, produced by the same routine. The
            # year travels in the QUERY STRING (route convention: Get-QueryYear
            # parses $path; only /setfilter and /setsetting read the body), and
            # POST + token is enforced by the gate above like every other
            # non-page route. Nothing calls this yet - it exists so a later
            # client can hold more than one year at a time.
            #
            # It deliberately touches NONE of the change-detection state: no
            # event drain, no watch re-baseline, no $script:contentHashes entry,
            # no $script:lastBuildHash. A payload fetched for a neighbouring
            # year must leave the displayed year's baseline exactly as /build
            # left it, or the next /dirty poll would compare against the wrong
            # year and report a change that never happened. It also does not
            # populate $script:pages or the disk cache: those hold PAGES, and a
            # payload is not one.
            #
            # v3.4.1: answered from the payload cache when one is held and
            # nothing has changed since it was made. The staleness test is
            # deliberately CRUDE - any item event anywhere retires every cached
            # payload - because this route serves a year that nothing is
            # watching, so there is no per-year signal to be cleverer with. One
            # extra read after a calendar change is a fair price for not
            # serving a year that is quietly out of date.
            $y = Get-QueryYear $path $Year
            $pc = $script:payloads[$y]
            if ($pc -and [long]$pc.Ticks -ge [long]$script:lastChangeTicks) {
                Send-Http $stream '200 OK' 'application/json' `
                    (JsonBytes ('{"ok":true,"cached":true,"data":' + [string]$pc.Json + '}'))
            } else {
                $yd = Get-YearData -Y $y
                $script:payloads[$y] = @{ Json = $yd.Json; Hash = $yd.Hash; Ticks = (Get-Date).Ticks }
                Send-Http $stream '200 OK' 'application/json' `
                    (JsonBytes ('{"ok":true,"cached":false,"data":' + $yd.Json + '}'))
            }
        }
        elseif ($path.StartsWith('/quit')) {
            Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"ok":true}')
            $script:quitRequested = $true
            try {
                Start-Sleep -Milliseconds 300   # let the page receive the response first
                # v2.1.18: close the window we KNOW about. $script:appHwnd is
                # the handle the geometry logic captured at startup and the
                # watchdog has tracked ever since, so Quit no longer has to
                # rediscover the window by its title.
                #
                # Title matching was the wrong mechanism and had failed twice:
                # the page retitles itself to "YearView (stopped)" the instant
                # the quit answer arrives, so the scan raced the retitle (fixed
                # by matching both forms in v2.1.14, still a race), and on a
                # slower machine that race simply widens - Get-Process must
                # enumerate every Edge process and marshal MainWindowTitle for
                # each, and a title read can come back empty while the window
                # is busy repainting. A handle has none of those properties: it
                # does not change when the title changes, costs nothing to
                # test, and cannot be confused with another application's
                # window. PostMessage is asynchronous, so a busy Edge cannot
                # stall this loop either.
                Close-ViewWindow
            } catch {}
        }
        elseif ($path.StartsWith('/dirty')) {
            try {
                # v2.1.1: ITEM events only. This filter used to be YV_*, so the
                # 2s dirty poll consumed the YV_AppQuit quit signal before the
                # link check could act on it - which is why v2.1.0's release
                # never fired and Outlook still zombified - and misreported the
                # quit as a structural calendar change into the bargain.
                $evs = @(Get-Event -ErrorAction SilentlyContinue |
                    Where-Object { $_.SourceIdentifier -like 'YV_Item*' })
                $changed = $false
                $structural = $false   # events or count: definitely a real add/remove/change
                if ($evs.Count -gt 0) {
                    $evs | Remove-Event -ErrorAction SilentlyContinue
                    $changed = $true; $structural = $true
                }
                $cnt = -1
                try { $cnt = [int]$script:watchItems.Count } catch {}
                if ($cnt -ge 0 -and $cnt -ne $script:lastCount) {
                    $script:lastCount = $cnt
                    $changed = $true; $structural = $true
                }
                # fingerprint: most recently modified item - THE edit detector.
                # COM ItemChange events proved undeliverable in this hosting
                # (two pumping attempts, v1.2.10), so detection is poll-driven by
                # design now. R-12's fixed throttle protected large mailboxes from
                # an O(n log n) COM sort per poll - but the cost is measured, not
                # assumed: the first sort is timed, and if cheap (<=100 ms, and it
                # measures ~0 ms on this calendar) the fingerprint runs EVERY
                # poll, giving edit detection within one poll interval. Slower
                # mailboxes self-tune to a 5 s or 10 s cadence.
                if ($script:modCheckInterval -lt 0 -or (Get-Date) -ge $script:nextModCheck) {
                    try {
                        $fsw = [Diagnostics.Stopwatch]::StartNew()
                        $script:watchItems.Sort('[LastModificationTime]', $true)
                        $newest = $script:watchItems.GetFirst()
                        if ($newest) {
                            $lm = $newest.LastModificationTime.Ticks
                            if ($script:lastModTicks -gt 0 -and $lm -ne $script:lastModTicks) { $changed = $true }
                            $script:lastModTicks = $lm
                        }
                        if ($script:modCheckInterval -lt 0) {
                            $ms = $fsw.ElapsedMilliseconds
                            $script:modCheckInterval = if ($ms -le 100) { 0 } elseif ($ms -le 500) { 5 } else { 10 }
                            Write-Host ("Fingerprint sort: {0} ms -> checking every {1}." -f `
                                $ms, $(if ($script:modCheckInterval -eq 0) {'poll'} else {"$($script:modCheckInterval) s"})) -ForegroundColor DarkCyan
                        }
                    } catch {}
                    $script:nextModCheck = (Get-Date).AddSeconds($script:modCheckInterval)
                }
                if ($changed) {
                    # v1.2.25: STALE-REVALIDATE, not clear-everything. The
                    # detection signals cannot say WHICH year changed, but
                    # clearing every cached year forced a full foreground build
                    # on the next year-switch. Instead the pages stay; any
                    # cached page built before this stamp serves INSTANTLY and
                    # then self-revalidates (quiet rebuild + content-hash
                    # compare): unchanged years dismiss silently, affected
                    # years refresh with "YearView updated". /delete keeps its
                    # explicit clear - a post-delete reload must never show the
                    # deleted item, even briefly.
                    $script:lastChangeTicks = (Get-Date).Ticks
                    Write-Host '  ~ Calendar change detected' -ForegroundColor DarkCyan
                }
            } catch {}
            $since = 0L
            if ($path.Contains('since=')) {
                [void][long]::TryParse(($path -replace '.*since=([0-9]+).*', '$1'), [ref]$since)
            }
            $isDirty = ($since -gt 0 -and $script:lastChangeTicks -gt $since)
            Send-Http $stream '200 OK' 'application/json' `
                (JsonBytes ('{"dirty":' + $isDirty.ToString().ToLower() +
                    ',"fp":' + ($(if ($isDirty -and -not $structural) {'true'} else {'false'})) + '}'))
        }
        elseif ($path.StartsWith('/new?')) {
            $qs = $path.Substring($path.IndexOf('?') + 1)
            $kv = @{}
            foreach ($p in ($qs -split '&')) {
                $pair = $p -split '=', 2
                if ($pair.Count -eq 2) { $kv[$pair[0]] = [Uri]::UnescapeDataString($pair[1]) }
            }
            try {
                $d = [datetime]::ParseExact($kv['date'], 'yyyy-MM-dd', $script:invCulture)
                $dEnd = $d
                if ($kv['end']) { $dEnd = [datetime]::ParseExact($kv['end'], 'yyyy-MM-dd', $script:invCulture) }
                if ($dEnd -lt $d) { $t = $d; $d = $dEnd; $dEnd = $t }
                # v2.0.0: create IN the selected calendar (CreateItem always
                # saved to the default); Items.Add(1) targets the folder itself
                $appt = (Get-SelectedCalendar).Items.Add(1)   # olAppointmentItem
                $appt.AllDayEvent = $true
                $appt.Start = $d
                $appt.End = $dEnd.AddDays(1)            # all-day End is exclusive midnight
                # v2.0.20: an empty 'cat' is the "(No category)" chip - leave
                # Categories untouched rather than assigning an empty string.
                if ($kv['cat']) { $appt.Categories = $kv['cat'] }
                Show-ItemInFront $appt
                $catLabel = $(if ($kv['cat']) { "'" + $kv['cat'] + "'" } else { 'uncategorised' })
                Write-Host ("  + New {0} item opened: {1:d} -> {2:d}" -f $catLabel, $d, $dEnd) -ForegroundColor Cyan
                Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"ok":true}')
            } catch {
                Send-Http $stream '200 OK' 'application/json' (JsonError $_.Exception.Message)
            }
        }
        elseif ($path.StartsWith('/goto?')) {
            # v2.1.12: open the Outlook calendar AT a date (blank-day right
            # click). Folder switch first, then CalendarView.GoToDate; if the
            # current view were somehow not a calendar view the navigation is
            # skipped but the folder switch still lands the user in the
            # calendar. ActiveExplorer can be null when only inspectors are
            # open - then the calendar's own explorer is displayed instead.
            $qs = $path.Substring($path.IndexOf('?') + 1)
            $kv = @{}
            foreach ($p in ($qs -split '&')) {
                $pair = $p -split '=', 2
                if ($pair.Count -eq 2) { $kv[$pair[0]] = [Uri]::UnescapeDataString($pair[1]) }
            }
            try {
                if ($kv['date'] -notmatch '^\d{4}-\d{2}-\d{2}$') { throw 'Bad date.' }
                $gd = [datetime]::ParseExact($kv['date'], 'yyyy-MM-dd', $script:invCulture)
                $cal = Get-SelectedCalendar
                $exp = $script:outlook.ActiveExplorer()
                if (-not $exp) { $exp = $cal.GetExplorer(); $exp.Display() }
                $exp.CurrentFolder = $cal
                try { $exp.CurrentView.GoToDate($gd) } catch {}
                try { $exp.Activate() } catch {}
                Write-Host ("  > Outlook calendar opened at {0}" -f $kv['date']) -ForegroundColor Cyan
                Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"ok":true}')
            } catch {
                Send-Http $stream '200 OK' 'application/json' (JsonError $_.Exception.Message)
            }
        }
        elseif ($path.StartsWith('/open?')) {
            $qs = $path.Substring($path.IndexOf('?') + 1)
            $kv = @{}
            foreach ($p in ($qs -split '&')) {
                $pair = $p -split '=', 2
                if ($pair.Count -eq 2) { $kv[$pair[0]] = [Uri]::UnescapeDataString($pair[1]) }
            }
            try {
                # S-5 / v2.1.28: resolve, then AUTHORISE against the selected
                # calendar - GetItemFromID is store-scoped, not folder-scoped.
                $item = Assert-SelectedCalendarItem ($ns.GetItemFromID($kv['id'], $script:calStoreId))
                if ($kv['start'] -and $item.IsRecurring) {
                    try {
                        $occStart = [datetime]::ParseExact($kv['start'], "yyyy-MM-dd'T'HH:mm:ss", $script:invCulture)
                        $item = $item.GetRecurrencePattern().GetOccurrence($occStart)
                    } catch {
                        Write-Verbose 'Falling back to recurring master.'
                    }
                }
                Show-ItemInFront $item
                Write-Host ("  > Opened '{0}'" -f [string]$item.Subject) -ForegroundColor Cyan
                Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"ok":true}')
            } catch {
                Send-Http $stream '200 OK' 'application/json' (JsonError $_.Exception.Message)
            }
        }
        elseif ($path.StartsWith('/delete?')) {
            $qs = $path.Substring($path.IndexOf('?') + 1)
            $kv = @{}
            foreach ($p in ($qs -split '&')) {
                $pair = $p -split '=', 2
                if ($pair.Count -eq 2) { $kv[$pair[0]] = [Uri]::UnescapeDataString($pair[1]) }
            }
            try {
                # S-5 / v2.1.28: see /open - resolve, then authorise against the
                # selected calendar before deleting anything.
                $item = Assert-SelectedCalendarItem ($ns.GetItemFromID($kv['id'], $script:calStoreId))
                $mode = [string]$kv['mode']
                # v2.1.28: whitelist the mode. These three are the only values
                # the page sends; anything else is a malformed request and must
                # not fall through to a branch that deletes MORE than was asked.
                if ($mode -ne 'item' -and $mode -ne 'occ' -and $mode -ne 'series') {
                    throw 'Invalid delete mode.'
                }
                # v2.1.28: "this occurrence" REQUIRES a recurring item. Before,
                # a non-recurring one fell through to the else branch and the
                # whole item was deleted - the same outcome by luck for a plain
                # item, but the wrong direction to fail in: a request to remove
                # one occurrence must never remove more. ('series' on a
                # non-recurring item does delete that item, which is exactly
                # what "entire series" means when there is only one.)
                if ($mode -eq 'occ' -and -not $item.IsRecurring) {
                    throw 'That item is not recurring - reopen YearView and try again.'
                }
                # The store applies deletions with a small lag. If this route replies
                # before the change is visible to a fresh enumeration, the client's
                # rebuild (~700 ms later) can read the item still present and /build's
                # re-baseline then masks the change FOR GOOD - nothing is left to
                # trigger a heal, and any browser reload serves the stale cache: the
                # "ghost bar that survives refresh" report. So: confirm the deletion
                # is visible before replying (cap ~2 s), and drop the page cache.
                if ($mode -eq 'occ') {
                    # delete ONLY the clicked occurrence; never fall back to the master
                    $occStart = [datetime]::ParseExact($kv['start'], "yyyy-MM-dd'T'HH:mm:ss", $script:invCulture)
                    $occ = $item.GetRecurrencePattern().GetOccurrence($occStart)
                    $subj = [string]$occ.Subject
                    $occ.Delete()
                    for ($w = 0; $w -lt 20; $w++) {
                        # a deleted occurrence becomes an exception; GetOccurrence throws once registered
                        try { [void]$item.GetRecurrencePattern().GetOccurrence($occStart) }
                        catch { break }
                        Start-Sleep -Milliseconds 100
                    }
                    Write-Host ("  - Deleted occurrence of '{0}'" -f $subj) -ForegroundColor Cyan
                } else {
                    $subj = [string]$item.Subject
                    $item.Delete()
                    # the move to Deleted Items invalidates the old EntryID; wait for
                    # it to stop resolving inside the calendar
                    $calFolderId = [string](Get-SelectedCalendar).EntryID
                    for ($w = 0; $w -lt 20; $w++) {
                        $still = $false
                        try {
                            $chk = $ns.GetItemFromID($kv['id'], $script:calStoreId)
                            if ([string]$chk.Parent.EntryID -eq $calFolderId) { $still = $true }
                        } catch {}
                        if (-not $still) { break }
                        Start-Sleep -Milliseconds 100
                    }
                    Write-Host ("  - Deleted '{0}'" -f $subj) -ForegroundColor Cyan
                }
                # server knows the content changed: a reload arriving before the
                # client's rebuild must get the splash + a fresh build, never the
                # pre-delete page from cache
                $script:pages.Clear()
                Send-Http $stream '200 OK' 'application/json' (JsonBytes '{"ok":true}')
            } catch {
                Send-Http $stream '200 OK' 'application/json' (JsonError $_.Exception.Message)
            }
        }
        elseif ($path.StartsWith('/categories')) {
            # v2.0.14: the FULL master category list (name + palette colour,
            # master-list order) for the category picker - the page payload
            # carries only categories in use this year. Read-only; POST +
            # token enforced by the global gate, like every non-page route.
            $entries = New-Object System.Collections.Generic.List[string]
            try {
                foreach ($c in $ns.Categories) {
                    $hex = $catColors[[int]$c.Color]
                    if (-not $hex) { $hex = '#BFBFBF' }
                    $entries.Add('{"name":' + (ConvertTo-JsonStr ([string]$c.Name)) + ',"col":"' + $hex + '"}')
                }
            } catch {}
            Send-Http $stream '200 OK' 'application/json' `
                (JsonBytes ('{"ok":true,"cats":[' + ($entries -join ',') + ']}'))
        }
        elseif ($path.StartsWith('/newcat?')) {
            $qs = $path.Substring($path.IndexOf('?') + 1)
            $kv = @{}
            foreach ($p in ($qs -split '&')) {
                $pair = $p -split '=', 2
                if ($pair.Count -eq 2) { $kv[$pair[0]] = [Uri]::UnescapeDataString($pair[1]) }
            }
            try {
                $name = ([string]$kv['name']).Trim()
                # v2.0.21: reject names that would corrupt storage, BEFORE they
                # reach Outlook. Comma and semicolon are Outlook's own
                # separators within the Categories property, and yv-filter.cfg
                # is a comma-delimited list of hidden names read back as raw
                # text - either character in a name breaks the category filter.
                # Control characters would corrupt that file and the diagnostic
                # log; 255 is Outlook's own name limit. The client applies the
                # same rules before sending, so these are the trust-boundary
                # guard rather than the message a user normally sees.
                if (-not $name) { throw 'Empty category name.' }
                if ($name.Length -gt 255) { throw 'Category names cannot be longer than 255 characters.' }
                if ($name -match '[,;]') { throw 'Category names cannot contain a comma or semicolon.' }
                if ($name -match '[\x00-\x1F\x7F]') { throw 'Category names cannot contain control characters.' }
                $existing = $null
                foreach ($c in $ns.Categories) {
                    if ($c.Name -ieq $name) { $existing = $c; break }
                }
                if (-not $existing) {
                    $usedCols = @{}
                    foreach ($c in $ns.Categories) { $usedCols[[int]$c.Color] = $true }
                    $pick = 0
                    foreach ($i in 1..25) {
                        if (-not $usedCols.ContainsKey($i)) { $pick = $i; break }
                    }
                    if ($pick -eq 0) { $pick = Get-Random -Minimum 1 -Maximum 26 }
                    $existing = $ns.Categories.Add($name, $pick)
                    Write-Host ("  + Category '{0}' created" -f $existing.Name) -ForegroundColor Cyan
                }
                $hex = $catColors[[int]$existing.Color]
                if (-not $hex) { $hex = '#BFBFBF' }
                # R-15: ConvertTo-Json copes with anything Outlook allows in a
                # category name, tabs and other control characters included.
                Send-Http $stream '200 OK' 'application/json' `
                    (JsonBytes ((@{ ok = $true; name = [string]$existing.Name; col = [string]$hex } | ConvertTo-Json -Compress)))
            } catch {
                Send-Http $stream '200 OK' 'application/json' (JsonError $_.Exception.Message)
            }
        }
        else {
            Send-Http $stream '404 Not Found' 'text/plain' (JsonBytes 'not found')
        }
    } catch {
        # v2.0.24: a request that ends here used to be dropped in silence - the
        # finally below closed the socket with no reply, which the browser
        # reports as a network failure and the client can only call "Bridge not
        # reachable". Answer instead, unless this request was already answered
        # (in which case a second write would corrupt the response).
        Write-Verbose $_.Exception.Message
        if ($null -ne $stream -and $script:respondedTo -ne $stream) {
            try {
                Send-Http $stream '500 Internal Server Error' 'application/json' `
                    (JsonError $_.Exception.Message)
            } catch {}
        }
    } finally {
        $client.Close()
    }
    if ($script:quitRequested) {
        Write-Host 'Quit requested from the year view - stopping bridge.' -ForegroundColor Yellow
        break
    }
}

}
catch {
    Write-Host ''
    Write-Host ('ERROR: ' + $_.Exception.Message) -ForegroundColor Red
    if ($_.InvocationInfo) {
        Write-Host ('At line {0}: {1}' -f $_.InvocationInfo.ScriptLineNumber, $_.InvocationInfo.Line.Trim()) -ForegroundColor DarkGray
    }
    Write-Host ''
    if ($HiddenChild) {
        try {
            $log = Join-Path (Split-Path $PSCommandPath -Parent) 'yv-error.log'
            ('{0}  {1}' -f (Get-Date -Format 's'), $_.Exception.Message) | Out-File $log -Append -Encoding utf8
            Add-Type -AssemblyName System.Windows.Forms
            [void][System.Windows.Forms.MessageBox]::Show(('YearView error: ' + $_.Exception.Message), 'YearView')
        } catch {}
    } else {
        Read-Host 'Press Enter to close'
    }
    exit 1
}
finally {
    if ($listener) { $listener.Stop() }
    Release-OutlookLink   # v2.1.0: unregisters YV_* events and drops every RCW
}
