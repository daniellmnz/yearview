# YearView v3.6.3 — changes from v3.6.2

Fixes the occasional full "loading" screen after moving an item.

## What was wrong

After a move, the page waited 0.7 seconds before redrawing. In that gap
YearView's regular check for calendar changes could notice the move and start
its own background update. The redraw then arrived while that update was
running, was told YearView was busy, and reloaded the page. The move had just
emptied the store of ready-built pages, so the reload showed the loading
screen and rebuilt from scratch.

It happened only sometimes because it depended on whether a check landed in
that 0.7 seconds.

## What changed

The redraw now starts the moment the move is confirmed, so the regular check
finds one already under way and leaves it alone. The delay was no longer
needed, because since v3.6.2 YearView waits for the move to show in Outlook
before confirming it. If a background update happens to be running already
when the move completes, the redraw waits for it to finish.

## Verification (Linux container — no Windows/Outlook execution)

- A Chromium test that models YearView handling one request at a time
  reproduced the collision on v3.6.2 in three runs out of three. On v3.6.3
  it did not occur in three runs out of three, and exactly one rebuild
  followed each move.
- The drag-to-move tests from v3.6.1 pass.
- The MSI payload extracts byte-identical.

## Smoke test

Move several items one after another. Each should update in place without the
loading screen.
