# YearView — Release History

**Version 3.7.6 · 27 November 2026**

A consolidated history from the v2.0.12 production baseline to the current
release. Entries state what changed and, where it matters, why — the root
cause rather than the symptom. Work before the baseline (the calendar
picker, server-side filter persistence, clean uninstall, and the interaction
fixes that led to 2.0.12) and the 1.x series are covered by their own
changelogs.

## The baseline

Version 2.0.12 was declared the production baseline: calendar switching with
per-calendar cached startup, server-side category filtering, create, open and
delete from the grid, and every programmatic navigation routed through a
deferral gate so that a completion could no longer destroy the loop waiting
for it. The MSI pipeline already used explicit sequence numbers throughout,
after the toolchain's relative-ordering resolver had been caught producing
different sequences from identical source.

## Window memory (2.0.13 – 2.0.22)

A chain of five releases, each exposed by field-testing the last. The loading
splash was made to open at full window geometry, because Edge carries a small
window's geometry to the next window opened from the profile (2.0.13). Size
and position were then remembered and validated against the current monitor
layout on restore (2.0.16); the maximised state was remembered separately,
keeping the last normal rectangle so restore-down returns to the user's size,
with six enforcement sites consolidated into one helper (2.0.17); the monitor
a maximised window was moved to was tracked, since moving a maximised window
never un-maximises it (2.0.18); and restore was fixed by placing the window
before maximising, because Edge ignores the command-line position when it
adopts a window into an existing browser process (2.0.19). The remaining
intermittency fell in 2.0.22, when the code stopped trusting a precondition
and began verifying the outcome instead: maximised, and covering the intended
rectangle's centre.

## Categories (2.0.14 – 2.0.21)

The bare name prompt became a picker showing Outlook's full master category
list (2.0.14), opening upward from its chip (2.0.15). A "(No category)" chip
was added for uncategorised items, reusing the empty-string key the filter
already used rather than inventing a parallel notion (2.0.20). New category
names gained validation — rejecting commas and semicolons, which are
Outlook's own separators and the delimiter of the stored filter, along with
control characters and excessive length — with identical rules mirrored
client-side and proven equivalent over an adversarial corpus (2.0.21).

## The Outlook lifecycle chain (2.0.23 – 2.1.11)

Thirteen releases, all tracing one question: how should a long-running bridge
relate to an application the user starts and stops independently?

It began with a dead COM link after an item inspector was closed, answered
with retry-on-failure because no liveness probe can close a
time-of-check-to-time-of-use gap (2.0.23). A failed connect at the dispatch
site was found to escape every per-route handler and close the connection
with no response at all, which the browser could only report as the service
being unreachable (2.0.24). Beneath that lay a structural fault: a cold COM
start blocked the single-threaded loop for tens of seconds, wedging every
queued request (2.0.25). Versions 2.0.26 to 2.0.29 refined how and when
Outlook was started and attached to — minimised launches, running-instance
binds, maturity and then readiness gates — and along the way found that the
background poll had been the thing launching Outlook, and that a page
outliving a bridge restart holds a dead token whose rejection surfaced as a
connection error.

Version 2.0.30 was the first design pivot, at the user's direction:
**YearView does not start Outlook.** Six releases of heuristics had been
attempts to judge from outside whether an externally launched Outlook was
ready; an Outlook the user starts is fully initialised, registered and
windowed. A dialog asks; the user obliges.

Version 2.1.0 found why Outlook so often ended up half-dead: the bridge's own
held references — above all the item-event connection points — pin a closing
Outlook in a windowless state that only Task Manager clears. Version 2.1.1
found that the quit signal added to fix that was being eaten by three
existing event consumers, and that binding a half-shut-down Outlook blocks
the loop outright; a window became required before any bind. Version 2.1.2
was the second pivot: losing Outlook mid-session ends the session with an
explanation rather than attempting a reconnect that had never been reliable.
Versions 2.1.3 to 2.1.9 made that behaviour hold — removing a condition that
let Refresh slip through, stopping a progress toast from overwriting the
message that mattered, replacing a cleanup that could itself block with an
outright process exit, detecting an Outlook that keeps running to finish a
send/receive, telling the page the session is ending before the dialog
appears, and reducing three competing messages to the one dialog the user
needs. Version 2.1.10 made the startup dialog a gate: OK is verified rather
than trusted, the dialog returns while Outlook remains windowless, and
Cancel declines to start YearView at all. The shortcut label also became
"YearView". Version 2.1.11 removed the last blocking read from detection
itself: asking Outlook how many windows it has is a COM call, and against an
Outlook closed mid-sync even that blocks — so detection now asks Windows for
Outlook's frames (hidden ones included, keeping tray-minimised Outlook
safe), which cannot block regardless of Outlook's state. Dialogs also became
per-monitor DPI aware, curing fuzzy text on high-DPI displays.

## Interface and freshness (2.1.12 – 2.1.15)

Version 2.1.12 added a blank-day right-click that opens the Outlook calendar
on that date, via a strictly validated token-gated route. The window title
moved the year to the front — "2026 YearView" — with all seven window-title
matchers renamed to follow and a build assertion proving the old form
appears nowhere. The same release closed the offline-staleness gap: change
detection is event-driven, so edits made while YearView was not running
raised no event and the cache stayed stale; the cached view still opens
instantly and now revalidates itself a few seconds after each start and
every ten minutes thereafter, with the poll chain threaded so exactly one
exists however the revalidation was invoked.

Version 2.1.13 settled the toast vocabulary — "YearView" always means the
page and "Outlook calendar" always means Outlook — ending the ambiguity of
"Updating calendar". Version 2.1.14 made Quit close the window reliably: the
title flip had broken the quit scan whenever the page retitled itself to
"(stopped)" before the deliberately delayed scan ran, so the scan now
matches both title states and sweeps twice. The same release removed the
word "bridge" from every message a user can see, it being an implementation
detail rather than anything the user chose to run. Version 2.1.15 reworded
the new-category tooltip to "Select an Outlook category…", describing what
the control mostly does — the panel is primarily a picker, with creation as
its secondary option. Version 2.1.16 made the red today box follow the
date: it was stamped at render time, and the background update
deliberately re-renders only when calendar data changes, so a page left
open past midnight kept yesterday’s box — a once-a-minute check now
moves it in place, with no rebuild and no server call. Version 2.1.17
sorted the category picker and the filter panel alphabetically — Outlook
returns its master list in its own order, which reads as no order at all
beyond a handful of entries. Version 2.1.18 extended that ordering to the
legend row and replaced the mechanism behind Quit: the window is now
closed by the handle the geometry logic already tracks, rather than
rediscovered by title — a title scan that raced the page’s own retitle
and lost more often on slower machines. Version 2.1.19 finished the job:
Quit pressed DURING a background update had been answered "busy" by the
in-build mini-server and retried until the client gave up, which a fast
machine hid and a slow laptop did not. Quit is now honoured mid-build —
the one state-changing thing that mini-server may do, and safe because
it changes nothing the build depends on. Version 2.1.20 corrected an
overreach in that change: it had also abandoned the calendar read
part-way, and the build then cached the truncated result, so the next
start opened a half-built year. A read now always completes — Quit is
answered and the window closed immediately, and the update finishes
unseen so the cache left behind is whole. Version 2.1.21 made the remembered
window geometry independent of timing: it had been captured only by a
periodic sampler, which does not run during a build and never runs after
the window has closed, so a resize made just before quitting — or during
an update — could be lost. The geometry is now also captured at the
moment of closing. Version 2.1.22 fixed the category filter not surviving a
restart when the last change was a hide: the filter seed is baked into
each cached page at build time and the cache is served verbatim on the
next start, while the startup revalidation compares calendar content
only and so never replaced a page that differed solely in its filter.
Every cached page is now re-stamped in place, and re-persisted, the
moment the filter changes. That proved insufficient — a write-path fix only
holds if it runs at the right moment with the right page loaded — so
version 2.1.24 moved the guarantee to the read path: a restored page’s
filter seed is overwritten from the filter file every time the cache is
loaded, and a stale seed can no longer survive a restart however it arose.
Version 2.1.25 fixed a restart within about ten seconds of closing opening
at the default size: that quick a restart meets the previous bridge still
alive, taking a second launch path that opened the window without ever
loading the remembered geometry. Version 2.1.26 acted on a robustness and
security review: config files are now written atomically, cache files from
other versions are deleted rather than left to accumulate, both launch
paths load persisted state through one shared routine, responses carry a
no-sniff header, and category colours are validated at the point they are
placed into markup. Version 2.1.27 corrected the quick-help panel: a
JavaScript escape had been left in HTML, where it showed literally, and
the right-click menu had grown two entries the help never mentioned.
Version 2.1.28 added an authorisation step to the item routes: an identifier
is resolved against the whole mailbox, so opening or deleting now also
requires the item to sit in the calendar being shown — protecting against a
stale identifier from a page built under a different calendar selection
rather than against any attacker. Version 2.1.29 added a Settings button
with a "confirm before deleting" option, on by default, and made the config
writer verify what it wrote — a silently failing write had been
indistinguishable from a successful one, which is how the category filter
came to be forgotten across restarts a second time. Version 2.1.30 fixed
the new settings route, which read its value from the URL while the page
sent it in the request body, so saving a setting always failed. Version
2.1.31 found the deeper cause behind both that and the recurring lost
filter: while a calendar read is running the in-build server answers
"busy" and expects a retry, and the settings saves never retried — the
filter save discarded the answer in silence. Both now wait the build out. Version 2.1.32 replaced the browser’s own
confirmation dialog with an in-page one, so the heading reads "YearView"
rather than the loopback address and port, which no script can change.
Version 2.1.33 removed every reference to the predecessor product from the
script, the documents and the installer metadata.

## Version 3

The 3.x series is the bulk of this history. It is set out below by sub-series,
each named for what that run of releases was mostly about.

### Scrolling and the continuous timeline (3.0.0 – 3.1.9)

Version 3.0.0 opened a new series with year navigation by mouse wheel and by
touch. Both defer to the page: they take effect only once there is nothing
left to scroll in that direction, so a legitimate scroll on a window too
short to show the whole year is never stolen, and Ctrl+wheel remains the
browser’s zoom. A vertical swipe suppresses the synthetic mouse events a
touch would otherwise generate, so it cannot start a drag-create. The
behaviour is on by default and can be switched off in Settings, which became
table-driven in the same release so that adding a setting is one entry
rather than another copy of the save-and-revert logic.

Version 3.0.1 changed no behaviour at all. It separated the calendar read
and payload serialisation from the page builder into their own routine and
added a token-gated route that returns a year's payload without a page
around it — the groundwork for holding more than one year in the client at
once. It shipped alone, and deliberately so: the code it touches owns the
hybrid reader, the perf timings, the truncation flag and the change-detection
baseline, none of which can be exercised without Outlook, so any change in
behaviour over the following days would have exactly one candidate. The
content hash was the one place the split needed judgement rather than a cut:
the routine now returns it and the caller decides whether it becomes the
baseline, so a payload fetched for a neighbouring year cannot displace the
baseline of the year on screen.

Version 3.1.0 completed the change v3.0.1 had prepared: the page became one
continuous timeline, scrolling without interruption from one year into the
next. Two things had to stop being global. Dates resolved from a single
page-wide year, so a click in a neighbouring year would have created an
appointment twelve months away — they now resolve from the block the click
landed in. And a year arriving above the reading position adds height there,
so the view is anchored on the topmost visible month row and corrected in
the same frame; without it every prepended year would have thrown the reader
a year backwards. The trimming rule needed a second look for a similar
reason: releasing whichever year is furthest from the one in view sounds
right, but while the timeline grows forwards the furthest year is the one
just fetched, which would have been discarded on arrival and fetched again
on the next scroll — tens of seconds of Outlook reading, repeatedly, for
nothing. Years are released from the end opposite the arrival instead.

Version 3.1.1 pinned the controls: the header and the category legend stay in
place while the timeline scrolls between them. It was done with sticky
positioning rather than a fixed shell around a scrolling pane, deliberately —
a shell would have moved the scrolling from the document to a container and
quietly invalidated the year loading, the scroll anchoring and the print
scaling, all of which had just been built and verified against the document
being the scroller.

Version 3.1.2 made the timeline smooth and fixed a defect the previous two
releases had carried. A year arriving had been rebuilding every month of
every loaded year — three years of grids destroyed and recreated in the
middle of a scroll — so the arrival lurched and the scroll stuttered. Only
the new block is built now, and it is spliced into place; a block inserted
above the reading position is paid for with a scroll correction in the same
task, so the browser never paints the intermediate state. The scroll handler
also measures elements, and was answering every scroll event rather than
once a frame. The defect was subtler: an appointment spanning New Year sits
in both adjacent years' payloads, because each year's Outlook restriction
admits an item that starts before that year ends and finishes after it
begins — so combining the payloads drew it twice, in two lanes, on both
sides of the boundary. It went unnoticed because the first simulations used
test data in which no item crossed a year.

Version 3.1.3 found why a year's arrival still jumped, and it was a
consequence of the previous fix. The browser adjusts the scroll position
itself when content changes above the reading position. While v3.1.0 and
v3.1.1 rebuilt the whole timeline, that heuristic had no surviving element
to anchor to and stood down, leaving the correction to the application.
v3.1.2 kept the surrounding blocks alive — which was the point — and in
doing so handed the browser exactly the anchor it needed, so both
compensations applied and the view moved by precisely one year. The cure is
to leave one mechanism in charge: the browser's is switched off, and the
application's now measures where a month row actually ended up rather than
assuming how far it should have moved, which is correct whether or not
anything else touches the scroll position.

Version 3.1.4 stopped a loading loop, and its two causes are both instances
of a pattern already in this history. The first: the correction that keeps
the view still moves the scroll position, and the event that raises is
indistinguishable from the reader scrolling — so a correction downwards was
read as travelling forwards and fetched the year ahead, whose release then
read as travelling backwards and fetched the year behind, indefinitely. The
second was geometry: with three years held and a reach of nine hundred
pixels, a tall window could be within reach of both ends at once, so
releasing a year at one end immediately re-armed the other. The fix is
partly numbers — four years held, a shorter reach, and no release of a
block still within a window of the view — and partly a rule that terminates
regardless: a year is fetched only after the reader has genuinely moved,
measured from a position recorded after the corrections rather than before.

Version 3.1.5 removed the year band from the single-year view, where it said
nothing the heading had not already said. Two smaller things came out of
doing it. The flag the renderer now consults had been declared alongside the
other settings, far below the first render that would read it, so it had to
move up — and it derives from the one seeded literal rather than a second
copy of its own, because the cached-page normaliser rewrites that
declaration and would not have rewritten a duplicate. And turning the
setting off while a year was in flight would have spliced that year back
into a view that had just collapsed to one, which is now refused on arrival.

Version 3.1.6 made continuous scrolling opt-in and fixed what printing did
while it was on. The print handler had restricted the sheet to the year in
view and then scaled that year to fit — but it measured the height to scale
by while the print stylesheet was not yet in force, so the years it had just
hidden were still counted, and the survivor was shrunk against a height
several times its own. A fraction of a year reached the paper. Nothing is
hidden or scaled now when the timeline is on: the page prints as it stands
and runs onto further sheets, with month rows kept whole across the break. A
single year is still scaled onto one sheet, which is the path almost every
print now takes.

Version 3.1.7 retired the wheel-and-swipe year paging that had opened the v3
series. It was built to reach for something continuous scrolling now does
properly, and keeping it as the fallback meant a single-year page could
still jump a year under a stray scroll — while the touch half had to
suppress the synthetic mouse events a swipe produces, so that it could not
start a drag-create. All of it is gone: with continuous scrolling off, the
wheel and touch scroll the page and nothing more, and the year changes by
the buttons alone. Which makes this the third feature in this history
retired rather than repaired once the design behind it moved on.

Version 3.1.8 closed the last two ways a settings change could be lost, and
both were the same old shapes. A save can spend a minute waiting out a
build; a background update finishing inside that minute navigated the page
away, and the retry loop stands down deliberately once a navigation is
committed — so the save was simply abandoned. Config saves now count as
pending actions, using the same gate item creation has used since the
identical fault there. And the settings are baked into each page at build
time, so a page built before the change still carried the old value and an
update that navigated to it put the setting back; the route now re-stamps
the pages already built, exactly as the filter route has since v2.1.22. The
substitution that does it is one function now rather than two copies, which
is what let the second copy be missing in the first place.

Version 3.1.9 fixed a wrong-direction fetch: scrolling forwards from the top
could load the year BEHIND first. The cause was a single stray upward wheel
delta — the small negative number a mouse or trackpad commonly emits at the
start of a gesture — reaching a trigger that asked only whether the movement
was upward and the page was at the top. Both conditions are true for anyone
about to scroll forwards, because that is where they are sitting. The
backward reach now needs a sustained upward gesture that any forward
movement or a pause clears, and the direction of travel is taken from the
reader's input rather than inferred from the scroll position. The travel
requirement had to be waived for that one gesture: at the top there is
nowhere to travel backwards to, so a rule meant to prove intent would have
made the earlier year unreachable instead — the same trap as the capped
distance a few releases earlier.

### Printing, and settings that carry more than a yes or no (3.2.0 – 3.2.4)

Version 3.2.0 added a setting for how many sheets a print is squeezed onto,
and in doing so made the settings mechanism carry something other than a
yes or no for the first time. The table that drives the panel gained a type
rather than a second code path: one entry declares itself a list of choices
and everything else — opening, changing, saving, putting the control back
when a save fails — runs through the same single route it always did. The
value is range-checked twice, at load and at the route, so a hand-edited
settings file cannot put an arbitrary scale factor onto a page. The choice
also replaced a hard-coded decision: printing a timeline had flowed onto as
many sheets as it needed since v3.1.6, and that is now simply one of the
options rather than something the application decided on the reader's
behalf.

Version 3.2.1 answered a performance report with instrumentation rather than
a guess. A refresh on the slower machine was taking nine to thirteen seconds
for around seven hundred items, and the existing timings named the calendar
read but went no further. The read has two halves with completely different
costs — a table pass that fetches whole rows with no per-item object, and an
object-model pass that pays a cross-process call for every property of every
item — and which one dominates depends entirely on the mix of items in the
calendar, since all-day and recurring items are deliberately routed to the
slower half. The two are now timed and counted separately and reported in
the same tooltip. One saving went in alongside: both readers were testing
the categories property and then reading it, paying for that round trip
twice on every item.

Version 3.2.2 acted on what v3.2.1 measured, and the measurement overturned
the expectation behind it. The guess had been that the fast table reader was
barely being used, because it routes all-day items to the slower path and a
year planner is mostly all-day bars. The split said otherwise: the table
half was carrying most of the items and still taking most of the time, at
around six milliseconds a row for a read that fetches no per-item object at
all. The reason was in the loop rather than the strategy — each of the seven
columns was being read with its own cross-process call, so the bulk read was
paying per-item costs after all. A row is now fetched once, by position,
against an explicitly ordered column set. Two releases, two wrong theories,
and a number that settled it in a single refresh.

Version 3.2.3 came out of a measurement that could not be read either way.
The one-call row read shipped in v3.2.2 appeared to have done nothing — but
the half of the reader it did not touch had slowed by the same factor in the
same run, so the machine was simply slower and the change was neither
confirmed nor cleared. A silent per-row fallback and a change that achieved
nothing look identical from the outside, so the two are now told apart: the
column positions are proved on the first row by reading two of them both
ways, and the count of rows actually read in one call is reported alongside
the timings. The wider lesson is about the numbers rather than the code —
a single before-and-after reading on a machine whose speed varies by a
factor of two says nothing at all, and the untouched half of a system is
the control that reveals it.

Version 3.2.4 finally named the cost the previous three releases had been
circling. With both halves of the reader timed and counted, they agreed:
each landed at roughly two and a half milliseconds per cross-process call
on the slow machine, and perhaps a fifth of that on the fast one. The time
was proportional to the NUMBER of calls and to nothing else — not to the
items, not to the columns, not to the strategy. That reframed the problem
from making calls cheaper to making far fewer, and the table can hand back
five hundred rows in a single call, which turns fourteen hundred calls into
three. The sequence is worth remembering: two wrong theories offered
confidently, then instrumentation, then a control that showed a regression
was really a slower machine, then a number that pointed straight at the
answer. The measurements were quicker than the guessing and considerably
more accurate.

### Print fidelity and the reading of the grid (3.3.0 – 3.3.6)

Version 3.3.0 dealt with two print complaints that turned out to be one
mechanism. The dialog showed the page as portrait when it printed
landscape, and its orientation control could not be changed — both because
a page that names its size in CSS takes those controls out of the dialog's
hands. The rule was also nested inside a print media block, which is
redundant at best and may be why the dialog ignored what it said. The rule
moved to the top level, and the size it names now comes from a setting, so
the choice exists where it can be remembered rather than in a dialog that
forgets it every time. It is the third setting added to a table that was
built for checkboxes, and the first whose value is a string — which needed
care in exactly two places: the select must not coerce it to a number, and
the re-stamping of cached pages must write it back quoted.

Version 3.3.1 corrected the previous release's assumption rather than its
reasoning. Naming the paper as a size plus an orientation — the form the
specification is written around — turned out to be ignored outright by the
print preview, which went on offering a live Layout control that read
Portrait and printed on a portrait sheet, while the page had been scaled
for a landscape one; so a year that should have filled a single sheet ran
onto two. Declaring two explicit lengths instead is the form browsers
honour most consistently. The diagnosis came from a screenshot rather than
from the code: the Layout control being live, rather than greyed out, said
the declaration was not reaching the dialog at all, which is the opposite
of what the previous release had assumed was happening.

Version 3.3.2 gave up on a battle worth losing, and won a different one. Two
releases had tried to tell the print dialog what paper to use — first by
name and orientation, then by explicit measurements — and both were ignored
outright. Continuing to declare a size was actively harmful: the page went
on scaling itself for a sheet the printer was never going to use, so a year
that fits one landscape sheet ran onto two portrait ones. The declaration
is gone, the dialog owns paper and orientation as it always really did, and
the setting now says what it actually does — which sheet to scale for. The
same release traced a freeze that took every Edge window down with it to
the application window having no profile of its own: it joined the user's
ordinary browser process, where a modal print preview runs a nested loop.
A private profile directory separates the two entirely.

Version 3.3.3 removed the paper setting that three releases had spent
themselves arriving at. Once the print dialog was left to own paper and
orientation — the only arrangement this browser would accept — a second
place to state the same thing was worse than useless, since the two could
disagree. What remains is the number of sheets to fit, scaled against an
assumed A3 landscape page, exactly as the application assumed for years
before the attempt to do better. Four releases to return to where it
started, minus one wrong belief: the page had been silently trusting a
declaration the browser was discarding, and that was the actual fault all
along.

Version 3.3.4 moved the assumed sheet from A3 landscape to A4 landscape, on
the reasoning that A4 is what printers actually hold and that anyone wanting
A3 can print to PDF and enlarge — from A4 landscape that is an exact step
up. The change is one number, but it is worth recording a coincidence that
had been quietly flattering the previous assumption: A4 portrait and A3
landscape are both 297mm tall, so every release until now had been scaling
correctly for A4 portrait without anyone intending it. That is why the
prints looked right on a dialog set to Portrait, and it is the sort of
accident that makes a wrong assumption survive testing.

Version 3.3.5 responded to a report of the Windows push-notification user
service consuming enough memory to make a 64GB machine sluggish, suspected
of being connected to printing. That service is a Windows component shared
by everything in the session, and YearView raises no notifications, so it
cannot be a direct client — but the browser window YearView opens is a
heavy one. The window now starts with notifications, background networking
and sync switched off, none of which have any purpose in a window that only
ever loads one page on the loopback address. This is elimination rather
than diagnosis, and it is recorded as such: it removes YearView's browser
from the set of possible causes rather than proving it was one.

Version 3.3.5 was a precaution rather than a fix, and is recorded as such.
Reports of the machine running out of memory pointed at the Windows push
notification service — a component this application neither uses nor can
control, and one with a long public history of growing without bound. What
could be done was to stop giving it reason to notice: the private browser
profile introduced two releases earlier means notification registration,
background networking, sync and component updates can all be switched off
for YearView's window alone, without touching how ordinary browsing
behaves. Whether that changes anything is a question only days of use can
answer, which is the honest position when the suspect belongs to someone
else's code.

Version 3.3.6 withdrew continuous scrolling from Settings without withdrawing
it from the application. Every part of it remains — the year blocks, the
neighbour loading, the scroll anchoring — and restoring the option is one
table entry. The only judgement the change required was what to do about a
settings file that already said the feature was on: honouring it would have
left a reader inside a continuous timeline with no control anywhere to leave
it, so the stored value is now ignored rather than obeyed, and left in place
rather than deleted. This release is built on v3.3.4, so the browser flags
added in v3.3.5 are not in it; the number still ascends, because it always
must.

### Starting the year where you like, and publication (3.4.0 – 3.4.9)

Version 3.4.0 let the year view begin at a month other than January. The
feature needed two years' events on one page, which is what the multi-year
machinery built for continuous scrolling already did — so the release cost
far less than it would have a month earlier, and the code withdrawn from
the settings panel two releases before proved its keep. The only real
refactor was lifting a single month row out of the year-block builder so
that twelve rows could be assembled from any sequence of year-and-month
pairs. The first attempt failed immediately on the temporal dead zone this
history has already recorded twice: state placed above the literal it
reads. It now sits directly beneath that literal, and a build assertion
checks the order rather than trusting it.

Version 3.4.1 gave the payload route a cache. The route had been written
deliberately to keep nothing — pages are what the cache holds, and a payload
is not a page — which was right when nothing called it, and wrong the moment
a view starting after January needed the following year on every page load.
Every background update navigates and so does every year switch, so a
several-second calendar read was being paid repeatedly for data that had not
changed. Payloads are now kept per year and retired by any change to the
calendar at all — a crude test, chosen because this route serves a year
nothing is watching, so there is no finer signal available to it. The cache
lives only in memory, so a restart still pays once; making it survive that
means a second encrypted file on disk with the same version and calendar
guards as the page cache, and was left for when the need is proven rather
than assumed.

Version 3.4.2 put the spill year into the page. Caching its payload a release
earlier had removed the calendar read but not the delay, because the delay
was never the read — it was a round trip after the page had already drawn,
and a second render when the answer came back. The months of the current
year appeared instantly precisely because they travel inside the page, so
the fix was to let the following year travel the same way. The page cache
then serves both years together with no further work, which is the neatest
part: nothing about caching had to change at all. The fetch remains as a
fallback for a page built before the setting was chosen, so a page that
lacks the second year still fills in rather than showing empty months.

Version 3.4.3 fixed two things that had each been visible for a while without
being understood. A pinned taskbar shortcut had been losing its icon on
every upgrade for many versions: the shortcuts take their icon from the
installer's icon table, which Windows extracts to a folder named after the
product code — and that code changes with every release by design, so the
pinned copy was left pointing into a directory that no longer existed.
Unpinning and repinning worked because it rebuilt the shortcut from the
new one. The installed icon file sits at a path that never changes, so the
post-install helper now points every YearView shortcut at that instead.
And a view starting mid-year needed two page loads before it became
instant, because changing the setting did not change the page already in
the cache: it was the rebuild that eventually embedded the following year.
The settings route now patches the spill payload into the cached page in
place, the way the filter seed has been patched since v2.1.22.

Version 3.4.4 moved "Year starts at" to the top of the settings panel. The
table that drives the panel has been its display order since v3.0.0, so
this is a reordering of three entries and nothing else — but it is worth a
line for the ordering principle it settles: the setting that changes what
is on screen comes before the two that change what happens later.

Version 3.4.5 taught change detection to watch the spill months. Editing an
appointment after December announced an update and then did nothing, while
editing one before December updated correctly and swept up the earlier edit
as a side effect — a pair of symptoms that describes the cause exactly. The
test asked one question, whether the content hash of the year on screen had
moved, and an edit in the following year moves no hash it was watching. A
view that shows two years has to watch two. The care needed was in keeping
the second hash strictly separate: it answers whether the spill months
changed and must never become the baseline for the displayed year, which is
the trap noted when the payload route was first written.

Version 3.4.6 added **Show item titles** to Settings: off, each item is
its coloured bar alone. It hides the text rather than removing it, because
opening and deleting an item both read the bar's own text — so a root class
clears the ink, the tooltip and everything else are untouched, and nothing
is rebuilt. The new
flag was seeded beneath its literal and applied before the first render,
by the rule the dead-zone faults established, so a page opening with titles
off never flashes them; and the third yes/no setting turned the server's
two-way if/else into a switch, which would otherwise have sent it to the
wrong variable.

Version 3.4.7 made titles a choice per category rather than all or
nothing. The single Settings checkbox gave way to a **Titles** column in
the Filter panel, beside the existing Show column. The choice is a list
of category names, which made it the first string setting the seeded-
literal machinery had carried since it learned to render strings, and so
the first to test what a string can hold. Category names reach a script
block and a re-stamper that reads up to the next semicolon, so the seeded
literal escapes both "<" and ";" — ordinary JSON escapes that decode to
the same text — and the value is checked at load and at the route. Each
checkbox has its own label and acts on its change event, because one
label around two boxes would send every click on a name to the first.

Version 3.4.8 added **Print category legend** to Settings, so the row
of categories can be left off the paper. The one subtlety was the
fit-to-pages scale, which is measured when printing begins, before the
print stylesheet is in force — the trap the v3.1 print fault recorded. A
print-only rule would have hidden the legend but left its height in the
sum, shrinking the months for room they no longer shared. So the legend
is removed by a class set at the start of printing, before the
measurement, and the class is cleared when printing ends.

Version 3.4.9 was the review before publishing the code. Three things
were wrong. The saved category filter went into the page's script block
escaped for JSON but not for HTML, so a hidden category named with
"</script>" could run script - the third seeded-literal channel, which the
earlier injection audit had not covered. Page assembly filled its
placeholders by repeated replacement, re-reading inserted calendar text,
and the settings re-stamper matched the first occurrence of a declaration
anywhere in the page; either let an appointment title - which anyone can
send as an invitation - blank the view. The fixes were a single-pass fill,
line-anchored markers, and the same escaping the titles list already had,
and they were checked by running the real PowerShell for the first time
rather than a transcription of it. The review also found the uninstall
leaving the app window's Edge profile behind, contrary to a comment that
said it did not, and the shared-machine exposure documented as view-only
when the page carries the token; the note was corrected and the
single-user requirement made explicit. Appointment titles were taken out
of the error log, and the publisher name became the GitHub handle.

### Page shading (3.5.0 – 3.5.1)

Version 3.5.0, the first feature after publication, added **Page shading**:
White, Light, Medium or Darker grey for the page's white areas. All of them
were already one colour variable, so each level is a set of variable overrides
on the page root, live and without a rebuild. The care was in what had to move
with the paper: at the darkest level the original weekend tint and grid lines
all but vanished into the grey, so each level carries its own darker weekend,
out-of-month, month-label and line colours, checked by rendering every level in
a real browser rather than by reasoning about hex values. The new seeded
setting followed the line-anchored rule from 3.4.9 and was exercised by the
same real-PowerShell harness.

Version 3.5.1 followed a day later in response to feedback: the lines should
darken steadily with the greys to keep their contrast. Measured, they had not:
line- to-page contrast held roughly level across the four levels and even
dipped slightly at Light, which reads as lines fading on a darker page. Each
level's line and outline colours are now solved for a rising contrast ratio
against its page, keeping the original blue-grey tint, and White is unchanged.

### The page review, and what measuring it found (3.6.0 – 3.6.13)

Version 3.6.0 came out of a review of the page as a whole and added seven
things: a Monday week start, a Today button and plain-key shortcuts, Find, a
hover card in place of the browser tooltip, hatching for tentative items, and
two options - item counts in the legend and a stripe for a second category.
Tentative marking needed one more field from Outlook, read as an optional
eighth Table column; if Outlook refuses the column the feature switches itself
off in every reader, so the fast path and its validation are never the price of
a cosmetic.

The release also found a fault that had sat unnoticed since v3.2.4. The reader
functions were run for the first time, against a mock Outlook, and the batched
row read threw on its first row: in PowerShell the comma binds tighter than
'+', so an index written [$r,$cLo+1] is a three-element index. The caller
caught the error and fell back to the classic reader for the session, which is
correct but slower - so, unless Windows PowerShell 5.1 parses that line
differently from PowerShell 7 (it should not), the batching credited with the
v3.2.4 speed-up has not been running. The indices are now parenthesised. It is
the clearest case yet for the rule this history keeps arriving at: code that
has only been read has not been tested.

Version 3.6.1 added the one item held back from 3.6.0 because it changes
calendar data in a new way: dragging a bar to another day to move the item. It
is hedged accordingly - a confirmation every time, and refusal of repeating
items, meetings and stale pages. Testing in a real browser found what
inspection would not have: after a double-click the bar's text is selected, and
the browser then begins its own text drag on the next press and stops sending
mouse movements. Bar text is no longer selectable.

Version 3.6.2 fixed the first field report on moving: the item moved in
Outlook, and the page went on showing it on its old day. The move route had
confirmed the save by asking Outlook for the item - which Outlook answers from
the object just saved - while the rebuild that follows reads the folder's
listing, which can trail it. The page was rebuilt from the old listing, and
because a build re-baselines change detection, nothing was left to notice the
difference afterwards: the trap the delete route's own comments had described
years earlier. The route now waits until a fresh listing shows the item at its
new date, and schedules one more quiet re-read if it does not. A second cause
was closed at the same time: the months after December are drawn from a kept
copy of the following year, which neither a move nor a delete was discarding.

Version 3.6.3 removed an occasional full "loading" screen after a move. The
page waited 700 ms before rebuilding, a delay inherited from the delete flow
and no longer needed once the server waited for the move to show. In that gap
the change poll could notice the edit and start a background update; the
rebuild then arrived mid-build, was told the bridge was busy, navigated, and
met a page cache the move had just emptied. The rebuild now starts at once,
which closes the gap. The collision was reproduced in a browser test that
models the single-threaded bridge before the fix was accepted.

Version 3.6.4 made the move confirmation a setting, **Confirm before moving**,
on by default and beside its delete counterpart. With it off, the toast names
the new date, since that is then the only statement of where the item went.

Version 3.6.5 made the hatching of tentative items a setting, **Show tentative
items hatched**, on by default. Off, they are drawn like any other item; the
hover card still names them.

Version 3.6.6 gave the split year a place in the sequence of years. Going
forward from a view such as November to October already showed the following
year from January; going back skipped to the year before, so this year's
January to October could not be reached at all. The arrows now walk last year,
this year whole, the split view, next year, each step undone by the opposite
arrow, and the two middle steps redraw in place without a reload. The change
also fixed a latent fault: with the following year held in the page for the
split view, showing a single year drew both.

Version 3.6.7 closed a measurement gap that had been open since the per-phase
timings were added in v3.2.1. Those timings covered the Outlook calls, so a
one-second read could sit under a three-second update with nothing to account
for the difference - which is what was asked about. The phases between the last
COM call and the finished payload are now timed too, and the stamp carries
them: the serialisation the v1.2.3 rewrite was assumed to have made negligible
turns out, on a container at least, to cost about as much as the read itself.
Two notes on the mechanics. A number appended into the payload is part of the
payload's SYNTAX, so the new figures are coerced rather than interpolated raw -
a regression test caught a missing value blanking the page, which is the JSON
equivalent of the unsubstituted placeholder trap. And the fallback serialiser
cannot time itself, because it builds the perf block before converting it; it
reports -1 and the page says so instead of showing a zero.

Version 3.6.8 acted on what v3.6.7 had just measured. The reported figures
from a real machine - 793 items, one second to read Outlook and one second to
serialise - matched the container's, so the serialisation was worth attacking.
The cause was not the volume of data. Each item's title, category and
identifier passes through an escaper that calls Regex.Replace with a
PowerShell scriptblock as the evaluator, and PowerShell converts that
scriptblock into a MatchEvaluator delegate on every call, whether or not the
pattern matches. Ordinary calendar text matches nothing, so the callback never
ran; the cost was entirely in preparing a callback that was never used. Three
calls an item made it about a second a year. The comment above the function
had asserted the opposite, reasoning from the callback never firing - correct
about the callback, wrong about where the time went, and it had stood
unchallenged since v1.2.3 because nothing measured it. The fix tests for
escapable characters directly and calls the escaper only when the answer is
yes: 764 ms to 6 ms on ordinary text, 1271 ms to 73 ms on text hostile
throughout.

Because the change claims to alter speed and nothing else, the verification
was built to prove the output had not moved rather than to exercise features.
A complete page was built by the real PowerShell from both versions over the
same items and compared byte for byte; the serialiser was run from both over
a corpus of injection attempts, control characters and null fields; and both
escapers were fuzzed from both versions over every code point to 0x2FF and
4,000 random strings. All identical. That last test matters more than it
looks, because the escaper has callers well outside the page build, and one
fuzz run covered them all without a test being written for each.

The same reported figures left two larger costs clearly visible and
unaddressed: the object-model reader at roughly seventeen times the per-item
cost of the bulk path, and first-build validation at three seconds a session.
Both had been suspected for several releases; neither could be ranked against
the other until something measured them. They also confirmed, by way of "bulk
788 in 2 calls", that the batched read repaired in v3.6.0 was running on
Windows PowerShell 5.1 at last - it had been failing over silently since
v3.2.4, which is the fault that prompted the per-phase timings in the first
place.

Version 3.6.9 closed two reported faults that proved to be one chain. A delete
sometimes ended in a full page rebuild, and after that rebuild the page came
back in the split view even though the delete had been made on a whole year.
The first was the v3.6.3 move fault verbatim, fixed there and not here: the
700 ms pause before refreshing let the change poll notice the user's own
deletion, start a background update, and leave the refresh to arrive
mid-build, be told busy, navigate, and meet an empty page cache. The pause had
never been needed, because the delete route already waits for the item to
leave the folder listing before replying - the same property that made the
move fix safe. The second fault lived in the loading splash, which rebuilt its
destination as "/?year=Y" and so discarded every other query parameter.
Carrying `whole` had been added to every navigation path in v3.6.6, but the
splash is not a navigation path; it is a separate page with its own three
lines of script, and it was missed. It was also quietly losing the armed
category; the hidden-category filter escaped only because it is persisted
server-side and reseeded from there.

Two things are worth drawing out. A fix applied to one of two paths that share
a shape is a fix half made, and the second path will be reported later as a
new fault - the v3.6.3 comment even described the mechanism precisely enough
to have found the delete case by reading. And state threaded through
navigation needs checking at every place a URL is CONSTRUCTED, not every place
it is navigated to; the splash's one-line destination was the only constructor
outside the client's navSuffix, and being in a different here-string put it
out of sight. Both faults were reproduced before being fixed - the delete race
3 times out of 3, the splash deterministically - and both verified by the same
harness afterwards.

Version 3.6.10 renamed one setting: "Year starts at" became "Current year
starts at". The setting had only ever governed the current year, every other
year being shown January to December, and the label had never said so. Worth a
line here only because the rename had to reach the help panel and the function
summary as well as the control, while the release history keeps the old wording
where it records what earlier versions did - a label is not one string once it
has been documented.

Version 3.6.11 came out of re-examining whether page load and refresh were
actually optimised, now that v3.6.8 had removed the serialisation cost. Timing
the phases the stamp still did not cover put most of the remaining suspicion to
rest: the R-17 garbage collect after each read costs about 11 ms, writing the
encrypted page cache 3 ms, and the prune scan beside it 3 ms, while the client
reaches first paint 216 ms after a 265 KB page starts loading. None of that is
worth touching.

What the review did find was waste of a different kind. The dual-reader
validation that costs three seconds already saved its verdict to disk so it
need not repeat - but keyed on $script:yvVersion, so every release discarded
it. Hashing the three reader functions across 3.6.7, 3.6.8, 3.6.9 and 3.6.10
showed them byte-identical: four releases had each re-earned the same verdict
because a number beside the code had moved. The key is now a hash of the
reader's own text, so the verdict is void exactly when the thing it attests to
changes and stands otherwise. The seven-day expiry, the delete-on-failure and
the first-run-on-a-machine cost are all unchanged, and any failure to compute
the signature falls back to validating - the fallback is the old behaviour,
never a skipped check.

Two notes on the method. The signature is taken by brace-matching the script's
own text rather than parsing it, which is about five times quicker at startup
but could in principle be defeated by an unbalanced brace inside a string; so
the build now computes it both ways and refuses to package unless they agree,
moving a runtime hazard to build time. And the finding itself only existed
because v3.6.7 had made the cost visible: an invisible three seconds had sat
in every session since the validation was introduced, attributed to nothing in
particular, and the fix was not to make the check faster but to notice it was
being asked for again and again with no new question to answer.

The review also restated the one large cost that remains and deliberately
left it alone. The object-model reader takes 0.9 s of a 1.0 s read, 160 items
against 633 on the bulk path, roughly seventeen times the cost each. The
all-day items go down it by design, to get Outlook's own normalisation across
two storage conventions, so moving them is a correctness risk and not merely a
speed change - and precisely the kind that cannot be settled without Outlook
to test against.

Version 3.6.12 gave the Find box AND and OR. The query parses into an OR of
ANDs, so AND binds first and any number of terms works, and a term still
matches against the title and the category together.

The design question was not the parsing but the operators' case. Lower-case
"and" is ordinary English in a subject line, so treating it as an operator
would silently break searches that had always worked; requiring capitals both
avoids that and gives a way to search for the word itself. The same reasoning
settled what a half-typed query should do: "abc AND" with nothing after it is
not an operator but the literal text "abc and", so the view does not flash
empty between keystrokes. That last behaviour was what the test suite caught -
the expectation written for it was wrong rather than the code, because the
corpus happened to contain the words "abc and" and the match was correct.

The change also became the first demonstration of v3.6.11: it does not touch
the reader, so the signature is unchanged and the release costs none of the
three-second validation.

Version 3.6.13 outlined the legend chip of a category the search matches. It
began as a question about whether Find covered categories at all - it had since
v3.6.0, and a test against the shipped build confirmed every case including an
item's second category - so the feature asked for already existed. What did
not exist was any sign of it: a search for a category name lit the right bars
and left the reader to infer why. The gap was in the feedback, not the
function, and that is worth recording, because a feature nobody can see
working is reported as a missing feature.

The rule chosen for the chip is the bars' own rule applied to the category
name alone, which means "travel AND work" marks no chip even though bars
light, since no single category satisfies it. Marking both would have looked
more responsive and said something untrue. Chips are also not dimmed the way
non-matching bars are: the legend is the control surface, and fading it during
a search would obstruct filtering exactly when it is wanted.

### Reading all-day items in bulk (3.7.0)

Version 3.7.0 moved non-recurring all-day items onto the table path, which
they had been kept off since the beginning because an all-day item has more
than one storage convention and a raw read cannot tell them apart. The
objection was sound and was never tested; a read-only diagnostic run against
the real calendar settled it. Across 2024, 2025 and 2026, 140 non-recurring
all-day items, every one stored the same way, and the table's values
reproduced the object model's exactly.

What made the change safe was not that finding but the shape of the guard. The
reader does not assume the convention it measured; it checks the OUTCOME -
both boundaries must land on a midnight, or within two hours of one - and
sends anything that fails to the object model, one item at a time, exactly as
before. An item stored against another time zone's midnight reads hours away
from local midnight, fails, and is fetched properly. So the optimisation is
correct everywhere and faster only where the values happen to verify, which is
the right way round.

The diagnostic also found a fault that had been shipping unnoticed. An all-day
item spanning a daylight-saving change comes back from the object model an
hour off midnight - late where the clocks go back, early where they go forward
- and since only an exact midnight counts as belonging to the previous day,
the late ones gained a whole day on the grid. Two items in the 2026 calendar,
none in 2024 or 2025, which is why nobody had seen it. Both boundaries are now
snapped to the nearest midnight, bounded at two hours so that the correction
can only ever repair a clock-change artefact and never relocate a value that
is wrong for some other reason.

Three things are worth keeping from how this was done. The objection in the
comment - two conventions, indistinguishable - had stood unexamined for many
releases because nothing could test it without Outlook; a sixty-line read-only
script answered it in an afternoon, and the answer was not the one the comment
assumed. The guard was written to verify rather than to trust, which is why a
finding from one calendar in one time zone could be shipped to anyone.
And the change had to be made in BOTH readers at once: Compare-YearReads
checks them against each other, so changing one would have made every
clock-spanning all-day item differ, failed validation, and dropped the session
onto the slow reader permanently and silently - the same failure that hid the
batched-read fault from v3.2.4 to v3.6.0. The test asserts the two readers
agree before it asserts anything else.

Version 3.7.1 undid a speed regression introduced by 3.7.0, reported from the
field as a twenty-second first build on the slower laptop. The new code tested
AllDayEvent to decide whether to correct an item's dates and then read the
same property again to record it - two cross-process calls where there had
been one, for every item in both readers. The classic reader runs over the
whole calendar as the validation half of a session's first build, so at about
two and a half milliseconds a call that was some two seconds added to the
slowest moment of a session.

The instructive part is that the function already said so. A v3.2.1 comment
two lines above reads "one property access, not two - the classic reader runs
in full on the FIRST build of every session as the validation half", written
about Categories for precisely this reason. The rule was stated, in the right
place, and still broken, because it was read as a note about one property
rather than as the invariant it is. A comment that records a conclusion is
weaker than a test or a gate; this one has now been broken once and that is
worth remembering when the next property is added.

Version 3.7.2 withdrew 3.7.0 and 3.7.1 and restored the v3.6.13 reader, after
the field reported a twenty-second first build and the tooltip showed the
reason: "reader: classic" on both machines. The fast reader had produced 768
items where the reference reader produced 769, validation had failed, and
every session was paying for both reads and then using the slower one. No
wrong date ever reached the screen - the comparison did exactly what it exists
for - but the optimisation had made things worse than doing nothing.

The mismatch log named both faults. First, an item carrying the all-day flag
in its stored properties that Outlook's object model reports as
AllDayEvent = False, because the April clock change makes its duration six
days and an hour rather than a whole number of days; the table trusted the
stored flag, the object model did not, and the readers disagreed about what
the item was. Second, an item starting at exactly the first instant after the
year, which the table path's own date test excludes and Outlook's Restrict
includes - a boundary disagreement that had been harmless since the beginning
because only an all-day item starts at exactly midnight, so only an all-day
item could ever land on it.

Both faults are instructive about the limits of the testing that preceded
them. The change had been exercised against a stand-in for Outlook, across
twelve time zones and 4,380 synthetic days, and it passed everything; a
calendar of 769 real items broke it twice in ways no synthetic case had
modelled, because both faults are properties of how Outlook itself behaves
rather than of arithmetic. A mock can only contain the behaviours its author
already knows about. The diagnostic script that preceded the change was the
right instinct and was not taken far enough: it compared dates, and the first
fault is a disagreement about a FLAG.

What held was the architecture rather than the judgement. The dual-reader
validation caught a defect its author did not anticipate, named the three
diverging items in the log with dates and flags and no private content, and
degraded to the slow-but-correct path on its own. The verdict being keyed to
the reader's code hash (v3.6.11) then made the restoration checkable rather
than asserted: v3.7.2's hash is byte-identical to v3.6.13's.

Version 3.7.3 fixed the boundary disagreement that v3.7.0 had exposed and
v3.7.2 had left dormant. The bulk reader decided year membership itself, in
PowerShell, while the reference reader asked Outlook; the two differed about
an item starting at exactly the first instant after the year. Outlook includes
it. The change is one character, -lt to -le.

What made it worth the afternoon was the END of the same test. The obvious
move is to make both comparisons inclusive, and it is wrong: measured on the
real calendar, "[Start] < 'D'" is inclusive of an item starting exactly at D
while "[End] > 'D'" is strict. Symmetry here would have pulled every item
finishing at midnight on 1 January into the previous year - a new silent fault
introduced while fixing one. Both sides were measured; only one was changed.

The measuring tool needed correcting twice before it could be trusted, and
both corrections are the point. Its first version inferred the operator by
probing, which meant enumerating the whole calendar and issuing filters such
as "[Start] < '01/01/2025'" - every appointment ever recorded, one
cross-process call each - and it ran for many minutes before being abandoned.
Its second version returned three verdicts from three dates, two of which
disagreed with the third; the cause was that it formatted probe dates as
MM/dd/yyyy on a machine that reads dd/MM, so "02/06/2026" was being read as
2 June and the probe was testing a boundary four months away. The source
comment above the filter had warned of exactly this, and explained that
YearView only ever uses 1 January because it reads the same either way.

Three things follow. A diagnostic is software and deserves the same suspicion
as the code it examines - this one was wrong twice and would have produced a
confident, wrong change both times. Disagreement between probes is worth more
than agreement, because only the disagreement revealed the locale fault; a
tool that had probed a single date would have reported "the whole day is
included" with no hint that anything was amiss. And a warning already written
in a comment is not a safeguard: the locale hazard was documented, in the
right file, a few lines from the code that walked into it.

Version 3.7.4 fixed a confidentiality fault in the per-category title filter,
reported from the field: a year printed to PDF with titles turned off still
contained them, invisible but selectable, and copying a bar revealed its
title. The v3.4.7 implementation wrote the title into the bar as usual and set
its ink to transparent, and the comment beside it said why - opening and
deleting an item both read the title back out of the element. That reasoning
was sound for the screen and silently wrong for paper. Transparent text is
still text: it prints, it answers the browser's Find, it is plain in the page
source and a screen reader would announce it. The title is now simply not
written, and the two readers take it from the item record the page has always
held beside each bar.

The lesson is about where a property is enforced. "Hidden" was implemented as
a presentation attribute, and presentation attributes do not survive a change
of medium - the PDF kept the content and discarded the styling that concealed
it. A confidentiality property has to be enforced by absence, not by appearance.

The testing is worth recording too, because the first attempt passed falsely.
Printing the page at a small page size makes each bar a few pixels wide;
Chromium then drops the clipped text, and both the faulty and the fixed
version reported no leak. The test only became a test once the sample items
were long enough for their titles to be legible on paper - the condition under
which the fault had been reported in the first place. A green result from a
test that cannot observe the fault is worth less than no test at all, because
it is believed.

Version 3.7.5 made the status line optional, and off by default. The line
reports what the page is and how it was built - generated time, mailbox, item
count, read timings, version - which is what makes it useful while something
is being diagnosed and what makes it unwanted the rest of the time, because it
names the mailbox on a page that is often shown to other people or printed.

The implementation took its instruction from 3.7.4 rather than repeating it.
The obvious way to hide a line is to hide the element, and that is precisely
the mistake the previous release had just finished undoing: a hidden element
keeps its text, and text survives the change of medium into a PDF. So "off"
here means the line is never written - no text, no tooltip, and removed from
the layout - and the test asserts the mailbox cannot be found in the page, in
what a copy would pick up, or in the text extracted from a printed PDF. The
same test checks the opposite case, that turning it on restores it without a
reload, since a setting that needs a restart to take effect is a different
feature from the one asked for.

One further path had to be covered. Settings are baked into each page when it
is built, and pages are cached; a change is applied to a cached page by
re-stamping the literal in place. Had that been missed, turning the line off
would have left it on screen until the cache happened to be rebuilt - the
setting would have appeared to work and then not have. That re-stamp is
tested directly, in both directions, along with a control proving an unchanged
setting rewrites the page byte for byte.

Version 3.7.6 added the running version to the foot of the Settings panel.
Turning the status line off by default in the release before had removed the
only place the version appeared, which is a reasonable thing to want back and
an easy thing to overlook when the change that removed it was about something
else.

The version had been written twice: once in `$script:yvVersion` and once in
the page's stamp string, with a comment on the first telling whoever changed
it to remember the second. A comment is not a gate — the same lesson 3.7.3
had already provided, where a documented locale hazard sat a few lines above
the code that walked into it. Rather than add a third copy for the panel, the
version is now substituted into the page from the single declaration, and the
comment that asked a human to keep two copies in step is gone with it.

## What the history shows

Three patterns are worth carrying forward. Field reports that say "usually,
but sometimes not" almost always name a condition rather than a defect — the
intermittent wrong-monitor restore, the intermittent connection error and
the intermittent Quit all resolved that way. Fixes that add cleverness to
protect an edge case have repeatedly had to be removed again, because the
qualifier itself became the fault. And the two design pivots, both proposed
by the user, each retired an entire class of problem that incremental fixes
had been chasing for several releases.

The verification discipline that grew alongside this is worth keeping: every
edit asserted to match exactly once with all writes gated behind every
assertion, the embedded client diffed line by line against the previous
release, the MSI sequence table and purge condition read back from the built
package, and behaviour modelled by simulation before release rather than
inferred. Several defects in this history — including two introduced by the
tooling itself — were caught by those gates before they ever reached a
build.
