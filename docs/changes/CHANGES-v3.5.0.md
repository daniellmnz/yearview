# YearView v3.5.0 — changes from v3.4.9

Settings has a new option, **Page shading**, second in the list after "Year
starts at". It tints the page's white areas:

- **White**, the default and the same as before;
- **Light grey**;
- **Medium grey**;
- **Darker grey**.

## What changes and what doesn't

The page background, the header and legend strips and the empty day cells take
the chosen tint.

The other parts of the page are shaded one step darker along with it, so they
stay distinct at every level:

- grid lines;
- weekend columns;
- month labels;
- days outside the month.

At the darkest level the original weekend tint and grid lines all but vanished
into the grey, which is why each level has its own set.

The coloured bars keep their own colours, and buttons and pop-up panels stay
white.

The change is instant, with no reload. It also applies to printing, so the
printout matches the screen. Choose White before printing if you'd rather not
spend ink on a grey background. Like the other settings, it is remembered
across restarts and re-applied to pages already built.

## Verification (Linux container — no Windows/Outlook execution)

- **Edits:** 14 asserted edits.
- **Gates:**
  - The client passes `node --check`.
  - Here-strings balance 4/4.
  - The placeholder check finds no placeholder missing or left over, including
    the new `__SHADE__`.
  - Ordering checks confirm the new value is set up before the first render.
  - The new setting's marker in the page is unique and line-anchored.
- **Rendered in Chromium at each level:**
  - The real page was rendered and the colours read from it confirm that each
    level steps the page, labels, out-of-month days and lines down together.
  - Screenshots confirm weekends, out-of-month days, lines and bars stay
    distinct at every level.
  - Changing the setting in the panel re-tints the page at once and saves
    `shade=2`.
  - The print view keeps the tint.
- **Real PowerShell 7:**
  - Page assembly and settings re-stamping were run, and all 13 adversarial
    page cases pass.
  - That includes an appointment titled `const SHADE = `, which leaves the
    data intact while the real setting is updated.
  - Loading and saving the setting with values 0 to 3 and with invalid ones
    (7, -1, text, empty): valid values stick, invalid ones fall back to White,
    and a saved value round-trips.
- **Regression:** the per-category titles and print-legend simulations were
  re-run and pass.
- **Packaging:** the MSI payload extracts byte-identical, and the sequence gate
  and wipe condition were read back from the built package.

## Smoke test

Open **⚙** and step through the four shading levels. At each one, check the
page tints at once and weekends, lines and bars stay easy to tell apart. Then
restart YearView and confirm the choice held.
