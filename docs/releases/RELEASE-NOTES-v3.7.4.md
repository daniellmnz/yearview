Install with the MSI; it upgrades in place over any earlier version and keeps
your settings.

**Fixed — a confidentiality fault. Please read if you have shared a printout.**

Turning titles off for a category in the Filter did not remove those titles
from the page. It drew them exactly as before and simply made the text
transparent. Transparent text is still text, so:

- **a PDF printed with titles turned off still contained them** — invisible on
  the page, but revealed by selecting a bar and copying it;
- the browser's own Find matched them;
- they were plain in the page source, and a screen reader would announce them.

A year printed with titles turned off was meant to be safe to pass on, and it
was not. **If you have shared or filed such a PDF, assume the hidden titles
went with it.** Printing again with this version produces a file that does not
contain them.

The title is no longer written into a hidden bar at all. Nothing else changes:
hovering a hidden item still shows its card, the delete prompt still names it,
and Find still matches hidden titles, because Find searches the calendar data
rather than the text on the page.

**One limit worth knowing.** The hover card still shows the full title of a
hidden item, deliberately, so you can tell your own items apart. The setting
conceals a printout, not a screen — if someone is watching while you work,
hovering will still reveal a title.

**On upgrading**

The calendar reader is unchanged in this release, so there is no startup
validation to re-earn.
