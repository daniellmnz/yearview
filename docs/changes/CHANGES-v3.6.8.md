# YearView v3.6.8 — changes from v3.6.7

The step v3.6.7 exposed has been fixed. Building the page's data was costing
about a second on a full year; it now costs a few hundredths.

## Why it was slow

Your desktop reported this for 793 items:

    … · enumerate 1.0 s (793 items) · mapping and hash 0.1 s · serialising 1.0 s

Serialising — turning the items into the page's data — cost as much as the
entire Outlook read. The cause turned out to be unrelated to the amount of
data.

Each item's title, category and identifier is passed through an escaper that
makes it safe to embed in the page. That escaper works by handing a small
block of code to the regular-expression engine to be called for each character
needing escape. Ordinary calendar text contains no such characters, so the
block never ran — but PowerShell still had to convert it into something the
engine could call, on every single call, and that conversion was the cost. At
three calls per item it came to roughly a second a year.

The comment in the code asserted this case was already cheap, on the reasoning
that the callback never fires. It never does fire. The expense was in
preparing it.

## The fix

The check for "does this text need escaping at all?" is now done directly, and
the escaper is called only for the rare string that does need it. The escaper
itself is unchanged, and still handles every character it did before.

Measured on the real code under PowerShell 7, 793 items:

| | v3.6.7 | v3.6.8 |
|---|---|---|
| Serialising, text containing hostile characters throughout | 1271 ms | 73 ms |
| Serialising, ordinary calendar text | 764 ms | 6 ms |

On your desktop the stamp should change from roughly

    · Outlook read 1.0 s · page data 1.1 s

to

    · Outlook read 1.0 s · page data 0.1 s

Nothing else about the page changes.

## What this does not fix

Your figures showed two larger costs still in place, both bigger pieces of
work than this one:

- **The object-model reader.** 160 of your 793 items took 0.9 s, while the
  other 633 took 0.2 s — about seventeen times the cost per item. These are
  the all-day and repeating items, which are read one at a time rather than in
  bulk. Bringing the non-repeating all-day items into the bulk read would be
  the single biggest remaining saving.
- **First-build validation, 3.0 s once per session.** A deferral has been
  discussed before and is still open.

Your screenshot did confirm one thing that had been in doubt: "bulk 788 in 2
calls" means the batched read repaired in v3.6.0 is now genuinely running on
Windows PowerShell 5.1. It had been falling back silently since v3.2.4.

## Verification (Linux container — no Windows/Outlook execution)

The whole point of this change is that it alters speed and nothing else, so
the testing was aimed at proving the output did not move.

- **The built page is identical.** A full page was built by the real
  PowerShell from both versions over the same 799 items, and the two pages
  compare byte for byte once the version string and the timing figures are
  set aside.
- **The payload is identical under hostile input.** The real serialiser was
  run from both versions over a corpus seeded with `</script>` tags, embedded
  quotes and backslashes, control characters, newlines and tabs, the page's
  own marker text, null titles and null categories. Byte for byte identical,
  and every value still decodes back to exactly what went in.
- **Both escapers fuzzed.** `ConvertTo-JsonStr` and `ConvertTo-SeedStr` were
  run from both versions over every code point from 0 to 0x2FF — alone and
  embedded — plus 4,000 random strings drawn from an alphabet of the
  characters most likely to break escaping, plus empty, null and 5,000
  character inputs. Identical on every one. This covers the other callers of
  the escaper as well as the page build.
- **Rendered in Chromium.** 798 bars drawn, no page errors. The hostile titles
  and categories all display as plain text: no injected element appears, and
  the injected script does not run.
- **The stamp reads correctly**, showing the new, much smaller figure.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## Smoke test

1. Hover the line beside the year. "serialising" should now read 0.0 s or
   close to it, and "page data" in the stamp should be a fraction of what it
   was.
2. Check that an item with a quotation mark or an unusual character in its
   title still shows correctly.
