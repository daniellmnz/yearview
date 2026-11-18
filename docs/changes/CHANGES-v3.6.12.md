# YearView v3.6.12 — changes from v3.6.11

**AND** and **OR** in the Find box.

    abc AND ghi      items containing both
    abc OR ghi       items containing either

As before, a term matches against the item's title *and* its category, so
`board AND plain` finds an item called "plain item" in the "abc board"
category.

## The details worth knowing

- **AND binds first.** `a AND b OR c` means `(a AND b) OR c`, which is the
  usual convention. Any number of terms works: `a AND b AND c`.
- **The operators must be in capitals.** Lower-case `and` is far too common in
  a subject line — "Bob and Alice" — to be taken as an operator.
- **That is also the escape hatch.** To search for the word itself, type it in
  lower case: `abc and ghi` looks for that exact text.
- **Nothing changes for a plain search.** One word behaves exactly as it did.
- **A half-typed query does not misbehave.** `abc AND` with nothing after it
  yet is not an operator — it is simply the text "abc and" — so the view does
  not flash to empty while you are still typing.

The count beside the box and **Enter** to jump to the first match work as
before, and **Esc** still clears.

## Verification (Linux container — no Windows/Outlook execution)

A page was built by the real PowerShell over a controlled set of items and
driven in Chromium. Eleven cases, all passing:

| Query | Found |
|---|---|
| `abc` | 5 — unchanged plain search, matching categories too |
| `abc AND ghi` | 2 |
| `abc OR ghi` | 6 |
| `abc AND xyz` | none |
| `xyz OR nothing` | 2 |
| `abc AND ghi OR xyz` | 3 — AND binding before OR |
| `abc and ghi` | 1 — lower case treated as literal text |
| `abc AND` | 1 — literal "abc and", not an operator |
| `abc AND zz` | none |
| `ABC AND GHI` | 2 — terms case-insensitive, operators not |
| `board AND plain` | 1 — one term matching a title, the other a category |

- **Esc** clears the query, the highlight and the count.
- No page errors.
- The reader, year-navigation and render suites pass; the full page still
  draws 798 bars and hostile text still displays as plain text.
- The reader's code is untouched, so this release does **not** cost the
  three-second validation — the first visible benefit of v3.6.11.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## Smoke test

Press **/**, type two words from different items with `AND` between them in
capitals, and check the count. Then change `AND` to `OR` and watch it grow.
