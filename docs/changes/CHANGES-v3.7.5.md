# YearView v3.7.5 — changes from v3.7.4

**The status line is now optional, and off by default.** The line under the
year naming the mailbox, the item count and the read times can be turned on
and off in Settings.

## What changed

A new checkbox, **Show status line**, sits with the other display options in
**⚙ Settings**. It is off by default, so an upgraded installation will not
show the line until it is switched on.

The line is worth having while something is being looked into — it says when
the page was generated, which mailbox it came from, how many items it holds
and how long Outlook took — and is worth not having the rest of the time,
because a page with a mailbox name on it is often the page being shown to
someone else or printed.

## How "off" is implemented

Off means the line is not written, not that it is hidden:

    stampEl.textContent = "";
    stampEl.removeAttribute("title");
    stampEl.style.display = "none";

This follows v3.7.4 directly. There, titles hidden by the Filter were drawn as
usual and made transparent, and transparent text is still text — it printed
into a PDF where selecting a bar revealed it. A hidden element has the same
defect: the text is still in the document. A line that names a mailbox has to
be absent when it is off, not merely invisible.

## Where the setting takes effect

Settings are substituted into each page when it is built, and pages are
cached, so a change also has to be applied to pages already cached. The new
literal joins the existing re-stamp list, which rewrites `const STATUS_LINE`
in place. Without that, turning the line off would have left it visible until
the cache was next rebuilt.

## Testing

* The page rendered in Chromium with the line off: the element carries no
  text and no tooltip, is out of the layout, and the mailbox name cannot be
  found in the page text, in the stamp element's markup, or in the text
  extracted from a PDF printed from that page. With the line on, all of it is
  present, as it should be.
* The checkbox: the panel opens showing the setting as it actually is, and
  toggling it changes the line immediately, with no reload.
* The bridge call and the settings file, both directions, including that the
  value is parsed correctly alongside other settings in one request, that an
  absent key leaves the default alone, and that only an exact `1` turns it on.
* The cached-page re-stamp, both directions, with a control confirming that
  an unchanged setting rewrites the page identically.
* The source gates: 14 seeded literals unique, line-anchored and declared
  before the first render; no unreplaced or orphaned placeholders.

## Not changed

The calendar reader is untouched — its signature is identical to v3.7.4, so an
upgraded installation does not re-run first-build validation.
