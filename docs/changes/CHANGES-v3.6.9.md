# YearView v3.6.9 — changes from v3.6.8

Two faults, both reported, and they turned out to be one chain.

## Deleting no longer triggers a full page rebuild

**No, it was not necessary.** It was the same fault the move path had in
v3.6.3, which was fixed there and left in place here.

After a delete the page waited 700 ms before refreshing itself. During that
pause the change poll could notice *your own* deletion, decide the calendar
had changed and start a background update. The refresh then arrived while
that update was running, was told "busy", navigated, and found no cached page
to show — so you got the "Building YearView page" screen and a full rebuild
instead of a quick refresh.

The pause was pointless either way: the delete request already waits for the
item to disappear from Outlook's own listing before it replies, so there is
nothing left to wait for when it does. The refresh now starts immediately,
which also claims the rebuild before any poll reply can be acted on. If a
background update is genuinely already running, it waits for that one instead
of colliding with it.

## The rebuild no longer loses your view

When a rebuild did fall through to the "Building YearView page" screen, that
screen sent you back to `/?year=…` and nothing else — so it dropped the rest
of the view state. That is why a deletion made on a **whole year** came back
in the split view: the whole-year flag was simply discarded on the way.

It was also silently losing the highlighted category. The hidden-category
filter survived only because it is saved on the server independently.

The loading screen now carries the whole query string through, so you come
back to the view you left. Anything unexpected in the address falls back to
exactly the old behaviour.

Because the two faults chained, you would normally only have seen the second
one when the first had happened — which matches the report.

## Verification (Linux container — no Windows/Outlook execution)

- **The delete fault was reproduced before being fixed.** With the change
  poll timed to fall during the delete, v3.6.8 hit the busy-and-navigate path
  3 times out of 3; v3.6.9, 0 times out of 3. The move path, tested the same
  way, stays at 0 out of 3.
- **The loading screen was driven directly** from a URL carrying the
  whole-year flag, a highlighted category and a filter. v3.6.8 sent the
  browser to `/?year=2026`. v3.6.9 sends it to
  `/?year=2026&whole=1&arm=Work&hide=Travel`.
- **The year-navigation suite from v3.6.6 passes unchanged**, including the
  whole-year view surviving a reload.
- **The page renders as before:** 799 bars, no page errors, and the hostile
  titles and categories still display as plain text with nothing injected.
- The script differs from v3.6.8 by 40 lines, all of them the two fixes, the
  comments explaining them and the version string — checked line by line.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## Smoke test

1. Set a split year, then press the left arrow to see the whole first year.
2. Delete an item from that whole-year view.
3. It should refresh quickly, and still be showing the whole year.
