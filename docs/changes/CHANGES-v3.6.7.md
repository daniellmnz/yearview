# YearView v3.6.7 — changes from v3.6.6

The line beside the year now says how long the work after the calendar read
took, which until now was never measured.

## Why

"Outlook read" counts only the time spent talking to Outlook. Everything
between the last call to Outlook and the finished page — mapping the colours,
building the legend, hashing the content, and turning the items into the
page's data — was never timed, so an update could sit on screen for seconds
longer than the figure explained.

## What you'll see

The stamp now reads, for example:

    … · 696 items · Outlook read 1.0 s · page data 1.1 s · YearView version 3.6.7

**page data** is the new figure. Hovering the line breaks both down:

    reader: fast · categories 0.1 s · restrict 0.0 s · enumerate 0.8 s (696 items)
    · baseline sort 0.0 s · mapping and hash 0.1 s · serialising 1.0 s · render 0.02 s

On a Linux container with 700 items, the serialising step took about as long
as the whole Outlook read. If your machines agree, that is where any further
speed-up should go. Please send me what your desktop and laptop report.

## Still not measured

- The 0.6 s pause YearView leaves before redrawing when an item is added,
  deleted or moved.
- Releasing Outlook's objects (about 0.06 s here).
- Encrypting and writing the page cache.
- Loading the redrawn page, except the part shown as "render".

Together those probably account for a few tenths of a second. They are not
shown because they happen after the page's data has been written, so the page
cannot carry its own figures for them.

## Verification (Linux container — no Windows/Outlook execution)

- The real timing code was run under PowerShell 7 with 700 items. The two new
  figures account for 1.35 s of a measured 1.40 s, so they cover the phase
  they claim to.
- The page was rendered in Chromium and the stamp and tooltip read correctly.
- A page built before this release shows no new figures and no gaps, so a
  cache from an older version still displays properly.
- The fallback serialiser's page says the step was not measured, rather than
  showing zero.
- A regression test caught a missing timing value making the page's data
  unreadable and blanking the year. The new figures are now coerced to whole
  numbers, so an absent one would cost a wrong figure rather than the page.
- The move, navigation, confirmation, tentative and v3.6.0 test suites pass.
- The MSI payload extracts byte-identical.
