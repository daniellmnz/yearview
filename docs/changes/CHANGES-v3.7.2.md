# YearView v3.7.2 — changes from v3.7.1

**v3.7.0 and v3.7.1 were wrong and are withdrawn.** This release restores the
reader exactly as it stood in v3.6.13, and keeps every other feature.

## What went wrong

v3.7.0 moved non-recurring all-day items onto the bulk read. It was tested
against a stand-in for Outlook and in nine time zones, and it passed. On a real
calendar of 769 items it failed, in two separate ways — both of which YearView
caught, which is why no wrong date was ever shown.

**1. An all-day item that Outlook does not call an all-day item.** One item
carries the all-day flag in its stored properties, but Outlook's object model
reports `AllDayEvent = False` for it. It is the item that spans the April clock
change: its duration is six days *and an hour*, so it is no longer a whole
number of days. The bulk read trusted the stored flag and called it all-day;
the reference reader asked Outlook and was told it was not. The two readers
disagreed about what the item even is.

**2. A boundary that only all-day items can land on.** The bulk read's own
date test excludes an item starting at exactly the first moment after the year;
Outlook's filter, which the reference reader uses, includes it. An item on
1 January 2027 was therefore dropped by one reader and kept by the other.

That second fault had been harmless since the beginning for a simple reason:
**only an all-day item starts at exactly midnight**, so only an all-day item
can land exactly on that boundary. Moving them onto that path was what exposed
it.

## What YearView did about it

Exactly what it is built to do. The two readers are compared item for item on
the first build of a session; they disagreed, so the fast result was discarded
and the slow, reference one used instead. **No wrong date was ever displayed.**

The cost was speed, and it was worse than before: every session paid for both
reads and then used the slower one. That is the 20-second first build reported
from the field.

## This release

The reader is restored to v3.6.13. That is verifiable rather than asserted —
the hash of the reader's code is `3DED2A016FD9B3EB`, identical to v3.6.13 and
to every release from v3.6.0 onward:

```
v3.6.13   3DED2A016FD9B3EB
v3.7.0    09F49A314C2850A1   withdrawn
v3.7.1    852FA53BB0C0EB7F   withdrawn
v3.7.2    3DED2A016FD9B3EB   restored
```

Everything else is unchanged: AND/OR in Find, the outlined legend category,
"Current year starts at", and the rest.

You should see `reader: fast` again, and about a second for the calendar read
rather than three or seven.

## What is not fixed

The two items spanning 5 April 2026 are drawn a day too long again, as they
were before v3.7.0. That fault is real but small, and it is not worth the cure
that was attempted. Any future attempt has to begin with the two findings
above.

## Verification (Linux container — no Windows/Outlook execution)

- The reader's code hash matches v3.6.13 exactly, by two independent methods.
- The reader suite passes with the two readers agreeing and the payload hash
  unchanged; the bulk read is back to 4 items where v3.7.0 made it 5.
- The navigation suite passes.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## On upgrading

The saved validation verdict is keyed to the reader's code, and the reader has
changed back, so the dual read runs once more on this upgrade. After that it
should be remembered and stay quiet.
