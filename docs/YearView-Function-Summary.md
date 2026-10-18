# YearView — Function Summary

**Version 3.4.9 · 17 October 2026**

YearView is a single-file PowerShell application that renders a full
year-at-a-glance view of the classic Outlook calendar in a dedicated window. It
puts an entire year of categorised commitments on one screen, with instant
startup, live refresh, in-view event creation and deletion, and persistent
category filtering.
Everything ships as one script (`New-OutlookYearView.ps1`) with no dependencies
beyond Windows PowerShell 5.1, classic Outlook, and Microsoft Edge.

## The year view

The window shows all twelve months as parallel columns of day cells, with
weekends banded, the current day marked by a red outline, and each appointment
drawn as a colour-coded bar in the cells it spans. Bar colours follow the
Outlook master category list; multi-day events flow across days and clip
correctly at month boundaries; overlapping events stack into lanes. Within each
day, bars are ordered deterministically: all-day events at the top, longer
spans above shorter within each group (so multi-day bars read as continuous
ribbons), then timed events earliest to latest. Events with no category draw as
hollow bars so they are distinguishable from genuinely grey categories.
Hovering a bar shows the subject and times. The header stamp reads
"Generated … · calendar · N items · Outlook read X.X s · YearView version
x.x.x", with a per-phase performance breakdown — including which
calendar reader served the build — available by hovering the stamp.

## Interacting with the calendar

The view is not read-only. Clicking a category chip in the legend arms it;
clicking empty day space then creates a new appointment in that category on
that date, opening the Outlook inspector — including a drag-across-days
range mode for multi-day entries. Clicking an existing bar always opens the
item in Outlook, whether or not a category is armed (an armed category
governs only empty day space), correctly resolving the clicked occurrence of
a recurring series. Item actions issued while the view is mid-update wait
the update out with an explanatory notice and complete first — even when the
update finishes mid-wait, the item is created or opened before the view
refreshes. Right-clicking a bar
deletes it: single occurrences are deleted as exceptions without touching the
master, and the server confirms the deletion has registered in the store before
the page refreshes, so deleted events can never linger. A "(No category)" chip sits at the end of the
legend for items that should carry no category at all. The "+ New category"
picker lists every category already defined in Outlook — select one to arm it, or type a name to
create a genuinely new one, checked for characters that would break Outlook’s
own category separators or the stored filter list — and the ‹ › controls move
between years, as does the mouse wheel or a vertical swipe on a touch screen.
Wheel and swipe only take over once the page has nothing left to scroll in
that direction, so on a window too short to show the whole year the view
scrolls first and moves between years only at the edge; the behaviour can be
switched off in Settings.

## Choosing the calendar

The Calendar button (beside Refresh) lists every calendar in the default
Outlook store — the default calendar plus any calendar-type folders in the
mailbox tree — with the current one marked. Selecting another switches the
entire application to it: the display, click-to-create (new items are created
in the selected folder), open and delete, live change detection, and the
header stamp all follow the chosen calendar, and the choice persists across
restarts. Switching is by a compact index into a bridge-held list, so folder
identifiers never travel between page and bridge. If a selected calendar ever
stops resolving — deleted, or its store detached — the app reverts to the
default calendar, clears the saved selection, and notes why, rather than
failing per-operation. Shared and delegate calendars belonging to other
mailboxes are out of scope by design in this version.

## Instant startup and cached years

Every successful build is persisted to disk (encrypted — see the security
section) beside the script, keeping the ten most recently built years. On the
next launch the bridge restores these pages before the window opens and serves
the entry request straight from cache, so the app opens into last session's
year in roughly one to two seconds — before Outlook has even been contacted.
YearView never starts Outlook itself, and does not start without it: if
classic Outlook is not running with a window when YearView launches, a
dialog asks you to start it. OK re-checks, allowing up to a minute for a
just-started Outlook to show its window, and the dialog returns if Outlook
still is not running; Cancel closes YearView instead (`-NoOutlookPrompt`
skips the gate). If Outlook is closed when a
calendar action needs it, the action shows a persistent request to start
Outlook and completes by itself once Outlook is up; background polling never
starts or wakes Outlook, so closing it stays closed until you act. Closing Outlook’s window ends the YearView session — including when
OUTLOOK.EXE keeps running to finish a send/receive, since YearView asks
Outlook how many windows it has open rather than trusting the process list.
Closing Outlook mid-session ends the YearView session — the message is acknowledged
before the view closes: the bridge releases its hold
on Outlook within a couple of seconds so Outlook exits fully, then the view
closes and a message explains that Outlook should be started before YearView
is re-launched. Reconnecting to a restarted Outlook is deliberately not
attempted — a fresh start is reliable where a reconnect is not. YearView only
ever connects to an Outlook that has a window and can answer, and its checks
on Outlook’s window ask Windows rather than Outlook, so a busy or closing
Outlook can never stall the view. Right-clicking a blank part of a day
offers to open the Outlook calendar on that date. The cached view opens
instantly and then quietly revalidates itself — shortly after each start,
and every ten minutes thereafter — so changes made while YearView was not
running appear without a manual Refresh. Otherwise the
Outlook connection is lazy: nothing before the window touches it, and the first
data request pays the one-off connect cost invisibly in the background. Each
restored page then revalidates silently; if the calendar changed while the app
was closed, the view refreshes with a "Calendar updated" toast, otherwise it
simply stays. Cached pages are version-stamped and calendar-stamped: a cache written by a
different version of the script, or for a different calendar than the one
selected, is discarded at startup (one splash build in either case),
guaranteeing restored pages always run current code and always show the
selected calendar.
The splash appears only on a true first run or immediately after an upgrade.

## Continuous scrolling across years (currently unavailable)

The option to switch this on was removed from Settings at v3.3.6. The
machinery below is intact and unchanged, and restoring the option is a
single table entry; nothing else was taken out. Until then the view shows
one year and the ‹ › buttons change it.

Turned on in Settings — it is off by default — the view becomes one unbroken
timeline rather than a single year. Scrolling past
December continues straight into January of the following year, so part of
two years is visible at once and the boundary travels up the screen like any
other month. A slim band names the year each run of twelve month rows
belongs to; it appears only with continuous scrolling on, since with a
single year on the page the heading already answers the question.

Neighbouring years are fetched in the background as either end comes within
reach, with a progress message, and only in the direction of travel — each
year costs a full Outlook read, so fetching the year behind a reader heading
forwards would spend that time on something they may never look at. About
four years are held at once and the far one is released, always from the end
opposite the arrival so a year cannot be discarded the moment it lands, and
never while it is still close to the view. A year is fetched only once the
reader has genuinely moved since the last one arrived, which is what stops
the loading chasing its own tail whatever the window size. The heading and
the window title follow whichever year fills the view, and the year buttons
scroll to an already-loaded neighbour instead of rebuilding.

The controls stay put while the timeline moves: the title, year buttons and
the Calendar, Refresh, Print, Settings, Help and Quit buttons remain pinned
at the top with the weekday strip directly beneath them, and the category
legend, filter and "+ New category" stay pinned at the foot. Both are
measured rather than assumed, so they stay aligned at any display scaling or
window width. On paper nothing is pinned — every row prints in flow.

A category used only in a neighbouring year joins the legend and the filter
as that year arrives, in the same alphabetical order, so it can be hidden
like any other. Turning the option off collapses the view back to a single
year, where the wheel and touch scroll the page and nothing else — the year
changes by the ‹ › buttons alone.

Printing follows the view. Paper and orientation belong to the print dialog;
a page cannot ask what was chosen there, and declaring a size in the
stylesheet — the documented way to influence it — is ignored outright by
this browser. **Print — fit to** therefore scales against an assumed A4
landscape sheet — the paper most printers hold, in the orientation a year
wants. On other paper the sheet count will be out; the dialog's own scale
control is the remedy, or print to PDF and enlarge, which from A4 landscape
is an exact step up to A3.

Printing follows the view, and how much it is squeezed is a setting.
**Print — fit to** takes one, two, three or four pages, or "As many as
needed"; whatever is on the page — a single year or a timeline — is scaled
onto that many sheets, with month rows kept whole across a break. One page
is the default, which is the year-at-a-glance sheet the application has
always produced. "As many as needed" scales nothing and lets the page run
on at full size, which is usually what suits a timeline.

**Print category legend**, on by default, decides whether the row of
categories at the foot of the page goes onto the paper. Turned off, the
printout is the calendar alone, and the fit-to-pages scale is worked out
without the legend, so the months get the room it would have taken. The
legend stays on screen either way.

## Starting the year at a chosen month

**Year starts at** in Settings takes January (the ordinary whole year),
**Current month**, or any named month. Choose anything but January and the
view still shows twelve month rows, beginning at that month and running on
into the following year — so from July you see July to December of this
year and January to June of the next. The heading says the span rather than
a single year.

Only the current year is affected. Stepping forward or back with the ‹ ›
buttons always lands on a plain January-to-December year, which keeps at
most two years in view at once.

The spill months belong to the following year, and the page is built
carrying that year's events as well as its own — so they are part of the
first paint, and a cached page serves both years together exactly as it
serves one. Only the first build after a change pays to read the second
year; its payload is then kept and reused, and any change to the calendar
retires it. Should a page somehow arrive without it — the setting changed
since that page was built, or the pre-read failed — the missing year is
fetched in the background instead. Change detection watches both years, so
an edit in the spill months updates the view exactly as one in the current
year does. An
appointment that crosses New Year is drawn on both sides of the boundary
but counted once, and a category used only in the following year joins the
legend and the filter as that year arrives.

## Titles per category

The Filter panel has two columns. **Show** chooses which categories appear
at all; **Titles** chooses which of them carry their item titles. Untick a
category's Titles box and its items are drawn as coloured bars alone —
same colour, size, outline and position, with no text — so a busy
category can recede while the ones that matter stay labelled. **All
titles** and **No titles** set every category at once, and (No category)
has its own box like any other.

Nothing else about a bar changes: hovering still shows its subject and
times, clicking still opens it and right-clicking still deletes it. The
change takes effect at once, applies to printing as it does on screen,
and is saved with the other settings, so it survives restarts and is
re-applied to pages already built.

## Instant year switching

Previously opened years switch instantly: navigation probes a millisecond
cache-lookup route and jumps straight to any cached year; only a never-built
year pays a build, with the "Loading 20XX…" progress toast. Calendar changes do
not invalidate other years' caches — the bridge uses stale-revalidate rather
than clear-everything, so after an edit the next year-switch is still instant,
followed by a silent background check that refreshes only if that year was
actually affected. Year navigation remains fully usable while a background
update is running: a cooperative quick-serve inside the calendar readers
answers page requests from cache even mid-build, and an uncached year requested
during a build gets a splash that retries until the bridge frees.

## Live refresh and change detection

A local bridge process watches the calendar with a two-second poll combining an
item-count check with an adaptive last-modification fingerprint (its cost is
measured on first run; on this calendar it is effectively free and runs every
poll). Edits, adds, and deletions made in Outlook surface in the view within
about two to six seconds. Notifications are honest and non-repeating: real
changes show "Updating calendar…" during the background rebuild and "Calendar
updated" on the refreshed page; the store's sync echo after a save (which
re-touches modification times without changing content) is detected by an
order-independent content hash — covering event fields and category colours —
and dismissed silently; and each detected change costs exactly one
revalidation cycle, after which the page's dirty baseline advances and polling
goes quiet. External adds/removes show "Calendar changed — updating…"; in-view
deletes show "Deleted … — updating…". The Refresh button forces a full rebuild
on demand.

## Performance

Reading the year uses a hybrid strategy: timed non-recurring items come from
Outlook's fast Table API (with build-independent UTC date columns converted
locally), while recurring series and all-day items go through the object
model's own recurrence and all-day normalisation. The fast reader is validated
item-for-item against the classic single-path reader on the live calendar, with
functional resolution checks on item identifiers; any error or mismatch falls
back to the classic reader with a logged, field-diagnosable diff. The
validation verdict is persisted per (script version, Outlook version) with a
seven-day expiry, so routine sessions skip the dual read and revalidate weekly.
Field results: a 753-item year dropped from ~11 s to ~4 s on a laptop and
~4.7 s to ~2 s on an i9 desktop. `-ClassicScan` forces the original reader.

## Application window behaviour

In `-App` mode the view runs in a chromeless Edge app window with the YearView
icon, sized to about 92% of the work area (or explicit `-WindowWidth` /
`-WindowHeight`) and centred. A loading splash appears only when no cache
exists, and it opens at that same full size — one geometry from launch to
year view, so there is no resize partway through and no small window for Edge
to carry over to the next browser session it opens. Resize, move, or maximise the window and the state is remembered and restored
on future launches — a maximised session reopens maximised on the same
monitor (even when moved between monitors while maximised), and restoring it
down returns to your remembered size (validated against the current monitor layout; explicit
`-WindowWidth`/`-WindowHeight` still override it). Launching again while a
window is open focuses the existing window rather than duplicating it.
A handle-anchored watchdog shuts the bridge down about ten seconds after the
window is genuinely closed, while surviving locks, RDP disconnects, and Edge
tab-sleep. Shortcuts launch through a windowless wscript/VBS chain so no
console box ever flashes, and the installers heal legacy shortcuts — including
pinned-taskbar copies.

## Security posture

The bridge is a single-threaded HTTP listener bound strictly to 127.0.0.1
(default port 8377, changeable with `-Port`). Every state-changing action is a
POST carrying a per-session token, with Host-header and Sec-Fetch-Site checks
rejecting cross-origin and DNS-rebinding attempts — applied identically by the
in-build quick-serve. All values crossing the HTTP boundary use
invariant-culture formats and fully escaped JSON. The on-disk page cache is
encrypted with Windows DPAPI (CurrentUser scope — decryptable only within the
owner's Windows account), and cache writes are atomic (write-then-rename), so
a crash can never leave a truncated file to be served. Request reads carry a
hard deadline, responses declare Connection: close, and diagnostics go to a
size-bounded yv-error.log beside the script.

## Modes, deployment and files

Without `-Serve`, the script writes a static, self-contained HTML snapshot of
the year. With `-Serve -App` it runs the full interactive application; adding
`-NoOpen` starts the bridge silently, which is how the optional sign-in
autostart works. Deployment is via a per-user MSI (no elevation; stops any
running bridge before replacing files; upgrades any earlier version in place;
heals legacy shortcuts) or an equivalent self-contained installer script; both
install six files to `%LOCALAPPDATA%\YearView` — the application script, the
windowless VBS launcher, a command-line launcher, a bridge-stop utility, the
shortcut healer, and the icon — plus Desktop, Start Menu, and Startup
shortcuts. Runtime data (encrypted page cache, validation verdict, calendar selection,
saved category filter, error log, and the app window's own Edge profile)
lives beside the script and survives upgrades; a genuine uninstall closes the
app window and removes all of it, leaving nothing behind. YearView is designed
for single-user machines; on a PC shared by several signed-in accounts,
another account could reach the running bridge (see the Security Design
note). The full switch reference:
`-Year`, `-OutFile`, `-NoOpen`, `-IncludeUncategorised`, `-ClassicScan`,
`-Serve`, `-App`, `-Visible`, `-WindowWidth`, `-WindowHeight`, `-Port`
(plus internal `-HiddenChild`); installer: `-InstallDir`, `-Startup`,
`-Uninstall`, `-NoShortcuts`, `-Quiet`.
