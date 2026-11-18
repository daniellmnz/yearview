# YearView v3.7.0 — changes from v3.6.13

All-day items are read in bulk now instead of one at a time, and a date fault
nobody had noticed is fixed.

## Faster

All-day items used to be fetched from Outlook individually, which cost about
seventeen times as much per item as the bulk read everything else uses. That
was deliberate: an all-day item has more than one storage convention, and
reading its dates raw could not tell them apart.

The diagnostic you ran settled it. Across 2024, 2025 and 2026 — 140
non-recurring all-day items — every one is stored the same way, and the bulk
read reproduces Outlook's own answer exactly.

So they are read in bulk now, but **only when the value proves itself**: both
ends of the item must land on a midnight, or within two hours of one.
Anything that does not is fetched individually, exactly as before. The reader
verifies the outcome rather than trusting a convention, so an item it cannot
vouch for is slow rather than wrong.

Recurring all-day items still have to be expanded by Outlook and are
unchanged. On your calendar that is 38 of 99, which is why the saving is
smaller than the raw figures suggest:

| | before | after |
|---|---|---|
| Items fetched individually | 160 | 99 |
| Reading the calendar | ~1.0 s | **~0.7 s** |
| Whole refresh | ~1.2 s | ~0.9 s |

## Fixed: all-day items spanning a clock change were drawn a day too long

An all-day item that spans a daylight-saving change comes back from Outlook an
hour off midnight — an hour late where the clocks go back, an hour early where
they go forward. YearView counts only an exact midnight as belonging to the
previous day, so the late ones gained a whole extra day on the grid.

In your 2026 calendar two items were affected, both spanning 5 April:

```
31 March -> 5 April    was drawn as 7 days, should be 6
 4 April -> 19 April    was drawn as 17 days, should be 16
```

2024 and 2025 had no such items, which is why this had never shown up. Both
boundaries are now snapped to the nearest midnight.

The two-hour limit on that correction matters. Every daylight-saving shift in
the world is an hour or less — Lord Howe Island's is half an hour — so two
hours covers every real case while refusing to drag a value that is off for
some other reason onto a midnight it never had.

## Verification (Linux container — no Windows/Outlook execution)

Both readers were run against a stand-in for Outlook in which the bulk read
and the individual read deliberately disagree, which is the condition the
whole change turns on.

- **The two readers agree on every item**, and `Compare-YearReads` reports
  Equal. This is the check that matters most: had the change been made to one
  reader and not the other, every all-day item spanning a clock change would
  have differed, validation would have failed, and the session would have
  fallen back to the slow reader for good — silently. That is the exact
  failure that hid the batched-read fault from v3.2.4 to v3.6.0.
- **Nothing is drawn twice.** The list of items fetched individually was
  narrowed from "recurring or all-day" to "recurring", so an all-day item
  appearing in both halves was the obvious way to break this. The test mock
  honours the real filter, and each item appears exactly once.
- **The fallback fires exactly when it should** — once, for an item stored
  against another time zone's midnight — and not at all when every item
  verifies.
- **Run in nine time zones**, chosen for the awkward cases: offsets from
  UTC-5 to UTC+13, two half-hour offsets and one of three-quarters of an hour,
  a zone whose daylight-saving shift is half an hour rather than a whole one,
  a zone that changes its clocks at midnight (so local midnight does not exist
  on that date), northern and southern hemispheres, and a zone with no
  daylight saving at all. All passed, with the number of individual fetches
  correctly varying by zone.
- Separately, the guard and the snap were run over every day of a year in each
  of twelve zones — 4,380 days — producing no wrong dates anywhere.
- **The two day-count faults are corrected** in the test, 7 days to 6 and 4 to
  3, and the unaffected items are untouched.
- The existing reader suite passes with the payload hash unchanged, and the
  item that moved to the bulk read shows there: 5 items read in bulk where it
  was 4.
- The navigation, find, legend and render suites pass; the page draws 798 bars
  with no errors and hostile text still displays as plain text.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## What to expect on upgrading

This release changes the reader, so you will pay the three-second validation
once — which is correct, since the reader is exactly what that check attests
to. After that it should stay quiet again.

## Smoke test

1. Hover the line beside the year: "enumerate" should be noticeably lower, and
   "object model" should show roughly 99 items rather than 160.
2. Look at 31 March and 4 April 2026. Each of those two items should now end a
   day earlier than it did.
