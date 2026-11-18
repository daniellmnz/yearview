# YearView v3.7.3 — changes from v3.7.2

One character, with a measurement behind it.

## The fault

YearView decides whether an item falls inside the year **twice**, in two
different places. The bulk reader tests it itself:

    $start -lt $rangeEnd   -and   $end -gt $rangeStart

The reference reader hands the same question to Outlook:

    [Start] < '01/01/2027'   AND   [End] > '01/01/2026'

Those two disagreed about an item starting at **exactly** the first instant
after the year. Outlook includes it; `-lt` excludes it.

It had sat there from the beginning without ever being noticed, for a reason
worth recording: **only an item starting at exactly midnight can land on that
boundary.** In practice that means all-day items, and all-day items have
always been read by the other path. It surfaced in v3.7.0 only because that
release moved them onto this one.

## What it could have caused

Nothing visible, most of the time. But on a session where the reader
validation is skipped — which is the normal case — the bulk reader runs with
nothing to compare against, and an appointment starting at midnight on
1 January would simply be **absent from the previous year's view**, with
nothing in the log. A silent omission is the worst kind.

## The measurement

Outlook's two comparisons do **not** behave alike, so this was measured rather
than reasoned about, on the real calendar, at three dates each:

| filter | Outlook's actual behaviour | the reader's test | verdict |
|---|---|---|---|
| `[Start] < 'D'` | **inclusive** of an item starting exactly at D | `-lt` | wrong — now `-le` |
| `[End] > 'D'` | **strict** | `-gt` | already right — unchanged |

**Only the start comparison is changed.** Making the end symmetric "for
consistency" would pull in every item finishing exactly at midnight on the
first day of the year — a new fault created while fixing one.

The measuring tool itself had to be corrected twice before it could be
believed: it first ran for many minutes because its probes asked Outlook for
every appointment ever recorded, and it then returned contradictory verdicts
because it formatted probe dates as `MM/dd/yyyy` on a machine that reads
`dd/MM/yyyy`, so `02/06/2026` was being read as 2 June. The contradiction is
what exposed it. Probes now use only days after the 12th, where the literal
cannot be read two ways — which is precisely why YearView's own filter has
always used 1 January and nothing else.

## Verification (Linux container — no Windows/Outlook execution)

- **The fault was reproduced before being fixed.** With a stand-in Outlook
  using the measured semantics, v3.7.2 drops an item starting exactly on the
  boundary while the reference reader keeps it — 2 items against 3, and the
  comparison reports not-equal. v3.7.3 keeps it, 3 against 3, equal.
- **The mirror case stays excluded in both**, so the end comparison really is
  untouched in behaviour as well as in text.
- The reader, navigation, find, legend and render suites pass, with the two
  readers agreeing and the payload hash unchanged.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## On upgrading

The reader changed, so the dual-reader validation is re-earned once on this
upgrade and remembered after that.
