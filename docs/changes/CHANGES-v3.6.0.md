# YearView v3.6.0 — changes from v3.5.1

Seven additions from the review of the page, plus a fix to the calendar reader
found while testing them.

## New

- **Today button and keyboard shortcuts.** **Today**, beside the year arrows,
  returns to the current year, scrolls to today and pulses its red box. Keys:
  - **←** and **→** change year;
  - **T** goes to today;
  - **/** jumps to Find;
  - **F** opens or closes the filter;
  - **P** prints.

  Keys are ignored while you are typing in a box, and Ctrl or Alt combinations
  are left to the browser.
- **Find.** A box in the header. Type part of a title or category name and the
  matching items keep their colour and gain an outline while everything else
  fades. A count appears beside the box. **Enter** scrolls to the first match;
  **Esc** clears it.
- **Hover card.** Rest the pointer on a bar and a card shows its full title,
  the date range, the times (or "all day" with the number of days), every
  category with its colour, and whether it repeats or is tentative. It
  replaces the browser's plain tooltip.
- **Tentative items are hatched.** An item Outlook marks as tentative, which is
  how an invitation you haven't accepted appears, is drawn with diagonal
  hatching over its colour.

## New options in Settings

All three take effect at once and are remembered.

- **Week starts on:** Sunday (the default, as before) or Monday.
- **Show item counts in legend:** off by default. Shows, beside each category,
  how many of its items fall in the twelve months on screen. Hidden categories
  are still counted.
- **Show second category as a stripe:** off by default. An item with two
  categories gets a thin band in the second category's colour along the foot
  of its bar. The main colour is still the first category.

## A reader fault, fixed

Testing the tentative marking meant running the calendar reader itself for the
first time, against a stand-in for Outlook. It failed on the first row of the
batched read added in v3.2.4. In PowerShell a comma binds more tightly than a
plus sign, so the index `[$r,$cLo+1]` is read as three numbers, not two.

YearView catches that error and falls back to the slower full reader for the
session, so the results were always correct. But it means the fast reader has
probably not been running on your machines since v3.2.4. The indices are now
bracketed.

This comes with two cautions:

- I confirmed the fault in PowerShell 7. Windows PowerShell 5.1, which YearView
  runs on, should parse the line the same way, but I can't run 5.1 here.
- With the fix, the fast reader should run for real. Hover over the line beside
  the year: the tooltip should begin "reader: fast" and show the table rows
  read in a few calls. If it says "classic", something else is stopping it.
  In that case `yv-error.log` will say why.

## What tentative marking costs

It needs one more field from Outlook (BusyStatus).

- **Ordinary timed items:** it is an extra column in the same batched read, so
  the cost is negligible.
- **All-day and recurring items:** it is one more call per item. On the slower
  laptop that is roughly a quarter of a second per hundred such items.
- **If Outlook refuses the extra column:** marking switches itself off for the
  session, and nothing else is affected.

## Verification (Linux container — no Windows/Outlook execution)

**Reader, run under PowerShell 7 against a mock Outlook**

- With the extra column accepted: the fast and full readers agree, and
  tentative flags are correct on table, all-day and recurring items.
- With the column refused: marking is off in both readers, and they still
  agree.
- The page's change detection notices when an item stops being tentative.

**Page, built by the real PowerShell and tested in Chromium**

- *Week start.* In both modes every one of the year's 365 day cells sits under
  the right weekday, and weekend shading falls on Saturdays and Sundays.
- *Creating an item.* Clicking a day in Monday mode creates the item on that
  date.
- *Find.* It matches on title and on category, shows the count, and Esc
  clears it.
- *Typing in Find.* Typing letters in the box does not trigger shortcuts.
- *Shortcuts.* Each one does its job.
- *Hover card.* The text is right for all-day, timed and multi-category items.
- *Options.* All three apply live and save.
- *Settings re-stamp.* Updating the settings inside an already-built page
  changes only the settings.
- *Hostile text.* A category named `</script>…` and a title made of the new
  setting's own marker text both display as plain text.

**Packaging**

- All PowerShell parses cleanly.
- The MSI payload extracts byte-identical.
- The install sequence was read back from the built package.

## Smoke test

1. Hover over the line beside the year and confirm it says "reader: fast".
2. Press **/**, type a word from one of your items, then press **Esc**.
3. Rest the pointer on a multi-day item and read the card.
4. In Settings, switch the week to Monday, and turn on the counts and the
   stripe.
5. Check that an unaccepted invitation shows hatched.
