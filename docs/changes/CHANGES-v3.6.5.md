# YearView v3.6.5 — changes from v3.6.4

The hatching on tentative items is now a setting.

## What changed

- Settings has a new checkbox, **Show tentative items hatched**, below "Show
  second category as a stripe". It is on by default, so nothing changes
  unless you untick it.
- With it off, tentative items are drawn like any other item.
- Resting the pointer on an item still shows "Tentative" in the card, whatever
  the setting.
- It takes effect at once and is remembered.

## Verification (Linux container — no Windows/Outlook execution)

Tested in Chromium on pages built by the real PowerShell:

- With the setting on, the three tentative test items are hatched.
- Unticking it removes the hatching at once and saves the setting.
- Ticking it again restores the hatching.
- A page that starts with it off shows no hatching, and the second-category
  stripes are unaffected.
- The hover card still says "Tentative" with the setting off.
- The setting is updated correctly inside already-built pages.
- The move and confirmation tests pass.
- The MSI payload extracts byte-identical.
