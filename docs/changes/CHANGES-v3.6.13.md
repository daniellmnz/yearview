# YearView v3.6.13 — changes from v3.6.12

A category matched by the search is now outlined in the legend, the same way
matching bars are.

## Why

Find has always looked at the category as well as the title — searching
`family` finds an item called "Delta planning" because that is its category.
But nothing said so. The right bars lit up and you were left to guess whether
the search had matched the category or just happened to hit those words in
some titles.

Now the category chip itself is outlined, so the answer is on screen.

## The rule

A chip is marked when the query would match an item whose only searchable text
is that category name — the same rule the bars use, applied to the name alone.

- `work` marks **Work**.
- `work OR family` marks both.
- `travel AND work` marks **neither**, even though bars still light, because no
  single category satisfies it. That is the honest answer rather than the
  flattering one.
- `planning` marks nothing: it matches a title, not a category.

Chips are deliberately **not** faded the way non-matching bars are. The legend
is the control surface, and dimming it during a search would make filtering
harder exactly when you want it. The outline also sits on top of the red ring
on an armed category and on a filtered-out category's dimming, without
disturbing either.

"(No category)" is never marked — those items carry no category text, so they
can never match a term.

## Verification (Linux container — no Windows/Outlook execution)

A page was built by the real PowerShell over items covering a single category,
two categories on one item, and no category at all, then driven in Chromium.
Seven cases, all passing:

| Query | Chips outlined | Bars |
|---|---|---|
| `work` | Work | 3 |
| `travel` | Travel | 3 — including an item where Travel is the *second* category |
| `family` | Family | 1 |
| `work OR family` | Family, Work | 4 |
| `travel AND work` | none | 2 |
| `planning` | none | 1 — a title match, not a category |
| *(empty)* | none | 0 |

- Arming a category while it matches keeps both marks — the chip ends up
  `item hit sel`, the outline and the ring together.
- The v3.6.12 AND/OR suite still passes, all eleven cases.
- The reader, year-navigation and render suites pass; the full page draws 798
  bars with no errors and hostile text still displays as plain text.
- The reader is untouched, so this release costs no validation time.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## Smoke test

Press **/** and type one of your category names. The bars in that category
should light, and that category's chip in the legend should gain an outline.
