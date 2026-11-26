# YearView v3.7.4 — changes from v3.7.3

**A confidentiality fault.** Titles hidden by the Filter were still present in
a printed PDF, invisible but selectable and copyable.

## What was wrong

Turning off titles for a category did not remove the title from the page. It
wrote the title into the bar exactly as before and then set its ink to
transparent:

    .bar.notitle{color:transparent}

Transparent text is still text. It prints into a PDF, where anyone who selects
the bar and copies it gets the title back. It also answers the browser's own
Find, shows in the page source, and would be read out by a screen reader.

The comment in the code explained the choice — the text was kept in the bar
because opening and deleting an item both read the title back out of the
element — and that reasoning held right up until a year was printed and handed
to somebody. A view printed with titles turned off was meant to be safe to
share, and it was not.

## The fix

The title is no longer written into a hidden bar at all. The two places that
used to read it back now take it from the item's own record, which the page
has always held alongside each bar.

Nothing else changes:

- **You can still identify a hidden item.** Hovering still shows its card, and
  the delete prompt still names it, so you can see what you are about to
  remove.
- **Find still matches hidden titles**, because Find searches the data rather
  than the page text.
- Turning titles back on restores them immediately, as before.

The transparent-ink rule is kept, but only as a second line of defence if some
later change ever puts text in such a bar again. It is no longer what does the
hiding.

## One thing it does not do

The hover card still shows the full title of a hidden item — deliberately, so
you can tell your own items apart. If you are hiding titles because somebody
else is looking at the screen rather than at a printout, hovering will still
reveal them.

## Verification (Linux container — no Windows/Outlook execution)

Tested by doing the thing that was reported: printing a real PDF from real
Chromium and trying to extract the text.

| | v3.7.3 | v3.7.4 |
|---|---|---|
| Hidden titles copyable from the page | **leaked** | none |
| Hidden titles extractable from the PDF | **leaked** | none |
| Ordinary titles still in the PDF | 2 of 2 | 2 of 2 |
| Bars marked hidden that still carry text | 2 of 2 | 0 of 2 |

- The delete prompt still names a hidden item, and its hover card still shows
  the title — both identical to v3.7.3.
- The reader, boundary, navigation, find, legend and render suites pass, with
  the payload hash unchanged.
- The reader is untouched, so this release costs no validation time.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

A note on the testing, because the first attempt was misleading: printed at a
small page size the bars are a few pixels wide, Chromium drops their clipped
text, and the test reported no leak from either version — a false pass. It only
became a real test once the sample items were long enough for their titles to
be visible on paper, which is the condition under which the fault was reported.

## Smoke test

Turn off titles for a category in the Filter, print to PDF, then select a bar
in the PDF and copy it. Nothing should come out.
