# YearView v3.6.10 — changes from v3.6.9

One label.

**Year starts at** in Settings is now **Current year starts at**.

The setting only ever governed the current year — every other year shows
January to December regardless — and the old wording did not say so. Nothing
about the behaviour has changed.

The help panel and the function summary now use the new wording too. The
release history keeps the old one where it describes what earlier versions
did.

## Verification (Linux container — no Windows/Outlook execution)

- The panel was opened in Chromium and the label read back: "Year starts at"
  before, "Current year starts at" after. The panel widens to fit it; nothing
  is clipped.
- The setting itself is unchanged — it is keyed on `startmonth`, not on the
  label — and the v3.6.6 year-navigation suite passes unchanged, including
  the April and January start-month cases.
- No page errors.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.
