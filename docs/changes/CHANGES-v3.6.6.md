# YearView v3.6.6 — changes from v3.6.5

When the year starts at a month other than January, the back arrow now goes
to this year from January. Before, it skipped straight to last year.

## The arrows now step through the years in order

With "Year starts at" set to, say, November, the arrows walk through four
views:

    last year  ‹ ›  this year, Jan–Dec  ‹ ›  Nov this year – Oct next year  ‹ ›  next year

- **‹ from your chosen view** shows this year from January to December.
  Before, it went to last year, so this year's earlier months couldn't be
  reached at all.
- **› from your chosen view** shows next year from January, as before.
- The opposite arrow always takes you back to where you were. From next year,
  **‹** returns to your chosen view. From last year, **›** shows this year
  from January, and **›** again shows your chosen view.
- **Today** goes straight to your chosen view.
- The ← and → keys follow the same order.

Switching between this year from January and your chosen view redraws in
place, with no reload. If the page updates itself while you are on this year
from January, it stays there.

With "Year starts at" set to January, nothing changes.

## Also fixed

Changing "Year starts at" back to January while the split view was showing
could draw two full years, one above the other. It now draws one.

## Verification (Linux container — no Windows/Outlook execution)

Tested in Chromium with a July start:

- Each arrow step in both directions shows the right heading and the right
  twelve months, including the round trips through last year and next year.
- The two in-place steps cause no navigation.
- Items from next year appear only in the split view.
- The ← and → keys and **Today** behave as described.
- Reloading while on this year from January keeps that view.
- Changing the setting to another month, and to January, each draws the right
  single view.
- With a January setting, back still goes to last year.
- The move, confirmation, tentative and v3.6.0 tests pass.
- The MSI payload extracts byte-identical.
