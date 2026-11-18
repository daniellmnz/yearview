# YearView v3.6.2 — changes from v3.6.1

Fixes the page still showing a moved item on its old day after the move had
succeeded in Outlook.

## What was wrong

There were two possible causes, and both are fixed.

- **The page was redrawn too soon.** After saving the move, YearView checked
  it by asking Outlook for the item, which Outlook answers from the copy it
  has just saved. The redraw that follows reads the folder's listing of
  items, which can lag a moment behind. The page was therefore rebuilt from
  the old listing. A rebuild also resets change detection, so nothing noticed
  the difference afterwards.
- **The months after December used a stale copy.** When the view runs on into
  next year, those months are drawn from a kept copy of that year's items.
  Neither a move nor a delete was discarding the copy.

## What changed

- After a move, YearView waits until a fresh listing shows the item at its new
  date before the page is redrawn. It waits up to five seconds, and usually
  far less.
- If the listing still hasn't caught up, YearView re-reads the calendar once
  more a few seconds later, quietly, and corrects the page.
- A move or a delete now discards the kept copy of the following year. This
  also fixes a deleted item in the months after December being redrawn.

## Verification (Linux container — no Windows/Outlook execution)

The move code was run under PowerShell 7 against a mock Outlook whose listing
can lag, in two time zones:

- Listing already up to date: the reply is immediate.
- Listing updates on the fourth look: the reply waits for it.
- Listing never updates, or can't be read: the move still succeeds and the
  later re-read is scheduled.
- An all-day item stored the way imported items are: recognised at its new
  date.
- The earlier refusal cases still save nothing.

The page itself is unchanged apart from the version number.

This fix is based on the symptom you reported and on how the delete code
already handles the same delay. I couldn't reproduce the delay itself here.

## Smoke test

1. Move a test appointment. The bar should appear on the new day within a few
   seconds.
2. If you use "Year starts at", also move one in the months after December.
