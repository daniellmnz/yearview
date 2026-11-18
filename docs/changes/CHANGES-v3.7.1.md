# YearView v3.7.1 — changes from v3.7.0

A speed regression in v3.7.0, found from a field report of a slow first build.

## What happened

v3.7.0 tested `AllDayEvent` when deciding whether to correct an item's dates,
and then read the same property again a few lines later when recording it.
Each of those is a separate call across to Outlook.

That cost one extra call **per item** in both readers — including the classic
reader, which runs over every item in the calendar as the validation half of a
session's first build. On a slower machine, at roughly two and a half
milliseconds a call, that is about **two seconds** added to the first build of
a year.

The function already carried a note about exactly this, added in v3.2.1 about
a different property: *"one property access, not two — the classic reader runs
in full on the FIRST build of every session as the validation half."* The same
rule applies to every property, and v3.7.0 broke it two lines below the
comment saying so.

Both readers now read the property once into a local and use it twice.

## What this does not explain

A first build is inherently the expensive one: it reads the calendar twice,
once with each reader, to prove they agree. On a slow laptop that has always
been the slowest moment of a session, and this fix removes roughly two seconds
of it rather than all of it. If a first build still looks disproportionate,
the figures to look at are in the tooltip on the line beside the year —
"enumerate" and "first-build validation" in particular.

## Verification (Linux container — no Windows/Outlook execution)

- Property reads per item in the classic reader, counted directly from the
  source: 8 in v3.6.13, 9 in v3.7.0, 8 again in v3.7.1.
- The reader suite passes in every time zone tested, with the two readers
  still agreeing and the payload hash unchanged.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## Note on upgrading

This changes the reader, so the three-second validation is re-earned once more
on this upgrade. That is the mechanism working as intended — the verdict is
keyed to the reader's code, and the reader changed.
