# YearView — Security Design Note

**Version 3.4.9 · 17 October 2026**

This note documents the localhost endpoint design, authentication model, and
data-protection measures of the YearView bridge, for security review. It
reflects an independent review cycle (July 2026) and a completed
context-by-context injection audit of every point where calendar-derived
text enters the rendered page.

## Architecture and exposure

YearView runs a single-threaded HTTP bridge bound exclusively to the
loopback interface (127.0.0.1, default port 8377, configurable). It is never
reachable from the network. The bridge serves one HTML application to a
Microsoft Edge app window and mediates all Outlook access via COM; the page
itself holds no credentials and no Outlook access. Exposure is therefore
limited to processes on the same machine.

**Single-user machines only.** Any process on the machine can request the page,
and the page carries the session token (see below). On a machine shared by
several signed-in Windows accounts - a Remote Desktop or multi-session host, or
a PC using fast user switching - another account could therefore read the
calendar YearView is showing and also use the token-gated routes: create, open
and delete items. The fixed port adds an accidental form of the same exposure:
a second user who starts YearView while the first user's bridge is running is
shown the first user's calendar. Earlier editions of this note described the
exposure as view-only; that was wrong, because the page that can be read
contains the token. YearView should not be run on shared or multi-session
hosts. The remedy, if that use is ever needed, is to gate page requests on a
per-user launch secret (for example a cookie set from a key file only that user
can read) and to choose the port per user.

## Request validation

Every request passes, in order: a read deadline (5 s overall) with a 16 KB
request cap and an 8 KB bound on any declared body; a Host header allow-list
(127.0.0.1/localhost only), defeating DNS-rebinding; a Sec-Fetch-Site check
(same-origin or none), defeating cross-site request forgery from browsers
that send fetch metadata; and route dispatch against fixed path prefixes
with a 404 default. The in-build cooperative server (quick-serve) applies
the identical Host and Sec-Fetch-Site checks and serves only read-only
cached pages, answering everything else 503. Rejections at the three
security gates are logged (method and query-stripped path) to a size-bounded
error log.

## Authentication model

A per-session random token is generated at bridge start, embedded once into
served pages, and required — via a custom request header — on every
state-changing route (build, create, open, delete, calendar switch, category
creation, quit), and on the read-only payload route (data). Page GETs are
intentionally token-free (the page is where the token comes from). The
token never touches disk in usable form: the encrypted page cache stores
the prior session's token only as a substitution
target, replaced at restore time; a restored page from a previous session
therefore carries the current session's token, and the on-disk copy is dead.

## Output encoding (injection audit conclusion)

Calendar-derived strings (subjects, category names) reach the page through
two channels, each encoded for its destination context. (1) The data
payload: a JSON document produced by a serialiser that escapes quotes,
backslashes and control characters, then has every "</" sequence neutralised
to "<\/" before insertion into the page's script block — the HTML parser
cannot terminate the block, and the JavaScript value is unchanged. (2) DOM
construction: event bars and labels are built with createElement and
textContent/property assignment (not HTML parsing); the small number of
innerHTML sites pass derived strings through an escaper covering & < > " and
use double-quoted attribute contexts; colours come from a fixed constant
palette. The audit enumerated every sink and found each correctly encoded.

**Correction (v3.4.9).** That audit covered the data payload and the DOM sinks
but missed a third channel: the saved category filter, which is written into
the page's script block as a string literal. It was escaped for JSON but not
for HTML, so a hidden category whose name contained "</script>" ended the block
and ran what followed, with the session token in reach. Since v3.4.9 every
seeded string (the filter and the per-category titles list) is written with "<"
and ";" as JSON escapes, so neither can end the block or the literal.

Two related faults in page assembly were fixed at the same time. The page's
placeholders were filled by a chain of replacements, each of which re-scanned
what the previous ones had inserted - calendar data included - so an
appointment titled "__HIDE__" corrupted the data and blanked the view; they are
now filled in a single pass that never re-reads an inserted value. And the code
that updates settings inside an already-built page searched for the first
occurrence of a declaration's text, which an appointment title could supply,
since the data comes before the code; it now matches only at the start of a
line, which calendar text (always inside a JSON string, where newlines are
escaped) can never produce. Both faults could be triggered by anyone able to
send the user a meeting invitation, and both could only blank the view, not run
code.

## Mechanical verification

The encoding claims above were checked with a test harness rather than by
inspection alone. That harness is not part of this repository, and it did not
include the seeded settings - which is how the filter fault above went
unnoticed. Its two instruments:
(1) an adversarial corpus (script-terminator payloads, attribute break-outs,
quote and backslash runs, control characters, Unicode line separators,
comment and CDATA sequences) is pushed through the serialiser semantics plus
the assembly-time "</" neutralisation, asserting three properties — decoded
JavaScript values are byte-identical to the originals, the assembled payload
contains no "</" sequence anywhere, and no raw control characters survive;
(2) an automated sink audit scans the extracted production client for every
dangerous sink (innerHTML, outerHTML, insertAdjacentHTML, document.write),
verifies each is either static markup or routes dynamic values through the
HTML escaper, and asserts the absence of event-handler attribute injection,
javascript: URLs, eval, and Function construction. Current result: all ten
corpus payloads pass all three properties; all ten sinks pass, including the
event-bar title (innerHTML wrapped in the escaper); no forbidden APIs
present.

For v3.4.9 the page-assembly and re-stamping code was run under PowerShell 7
against the real client template and the resulting pages rendered in a browser
DOM, with: appointment titles equal to each placeholder and to each seeded
declaration; category names containing "</script>" in both seeded lists; and a
hidden category name containing a newline followed by a declaration. In every
case the page rendered, no injected script ran, the data parsed back unchanged,
and each re-stamp altered only its real declaration.

## Data at rest and process hygiene

The saved category filter (v2.0.8) is persisted server-side in yv-filter.cfg beside the script, replacing an earlier browser-localStorage approach whose data lived in the Edge profile beyond the installer's reach; a genuine uninstall removes every runtime file (cache, calendar selection, filter, window geometry, verdict, log) and, since v3.4.9, the app window's own Edge profile, which earlier releases left behind, while version upgrades deliberately preserve them; the uninstall data-wipe custom action runs at an explicit, verified position in the installer sequence (bridge stopped, then data wiped, then files removed). The filter value is injected into the page as a complete JSON string literal into an unquoted template slot. (An earlier edition said script-terminator category names were safe here; they were not until v3.4.9 - see the correction above.) The on-disk page cache (rendered calendar content) and its metadata are
encrypted with Windows DPAPI (CurrentUser scope) and written atomically
(write-to-temp, rename), so no truncated file can ever be served. Since v3.4.9
the error log records no appointment titles or category names (a reader
self-check used to list differing items by title), so no calendar text rests
in plaintext there either. The calendar selection file stores
folder identifiers only. Native interop is limited to fourteen user32
window-management functions (find/show/move/focus/measure); no process, file-system,
or memory primitives are exposed. Command-line construction for the hidden
respawn uses explicit MSVCRT-correct argument quoting. Launch parameters
carry ValidateRange constraints. The installer is a per-user MSI requiring
no elevation; scripts run with a per-process execution-policy bypass (no
system policy change) — code signing can replace this where a signing
certificate is available.

## Outlook object-model surface

An audit (July 2026) enumerated every Outlook COM member the script calls, in
response to a programmatic-access prompt observed during development. The
complete set of item properties read from any calendar item is: Start, End,
Subject, AllDayEvent, Categories, IsRecurring, MessageClass, EntryID, and the
recurrence pattern; the only schema properties used are calendar start/end
dates and the recurring/all-day MAPI flags. The script does NOT access the
Contacts folder, the address book, the Global Address List, recipients,
attendees, organiser, sender, or any address-resolution member — verified by
exhaustive search, all absent.

Two calls can nonetheless trigger Outlook's programmatic-access guard, which
is raised by broad object-model areas rather than by contacts specifically:
opening an appointment for the user (Display / GetInspector), and expanding a
recurring series (GetRecurrencePattern / IncludeRecurrences). Neither reads
contact data; the guard's generic wording ("a program is trying to access…")
does not distinguish opening a meeting from reading an address book. A prompt
seen during use is therefore expected behaviour from these paths, not data
access. The supported way to suppress it for legitimate automation is to
ensure antivirus is registered with Windows Security Center; weakening the
Trust Center programmatic-access setting is neither required nor advised.

## Authorising an item against the selected calendar

An item identifier supplied by the page is resolved with a store-scoped
lookup, so it can name any item in the mailbox rather than only the calendar
in view. Since v2.1.28 every route that opens or deletes an item authorises it
first: the item must be an appointment AND its parent folder must be the
calendar YearView currently shows, matched on both folder and store
identifier. Anything else — a mail item, an appointment in another calendar of
the same mailbox, an item whose parent cannot be read — is refused.

This is best understood as a correctness and data-loss control rather than a
barrier against an attacker. Reaching these routes already requires the
session token, which exists only in YearView's own page and the process
serving it; anyone able to supply it can drive Outlook directly and delete
whatever they wish without involving YearView. What the check actually
prevents is a STALE identifier being acted on: identifiers are fixed into a
page when it is built, and a page can outlive the calendar selection it was
built under — a second window, a page restored from cache, or a deletion
already in flight when the calendar is switched. Deletion is silent and
unconfirmed by design, so without this check the first sign of the fault would
be a missing appointment in a calendar the user was not even viewing.

The delete route additionally accepts only its three known modes, and refuses
to treat a non-recurring item as an occurrence — so a malformed request can
never fall through to a branch that deletes more than was asked.

## Accepted risks and rejected mitigations (with reasons)

Fixed port rather than randomised: the port anchors single-instance
detection, stale-bridge takeover on upgrade, and the stop utility;
randomisation adds obscurity only, as all actions are token-gated. No rate
limiting: the client legitimately bursts (polling, in-build retries) and
throttling reintroduces resolved failure modes on an interface only local
software can reach. No COM call timeouts via background tasks: cross-
apartment calls into STA-bound Outlook objects deadlock; a watchdog bounds
the blast radius instead. Exposure to other signed-in users on shared
machines (full access, not view-only - see Architecture): accepted only for
the single-user deployment YearView is designed for, and stated as a
requirement in the README.
