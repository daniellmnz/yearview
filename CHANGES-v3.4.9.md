# YearView v3.4.9 — changes from v3.4.8

This release is the pre-publication review. It contains no new features: it
fixes what the privacy, security and robustness check found before the code
goes on GitHub.

## Privacy

- **Publisher name.** The MSI's Publisher, Author and support-contact fields
  now read **daniellmnz** instead of a personal name.
- **Build script example.** The example signing-certificate path in
  `build-msi.ps1` no longer carries initials.
- **Error log.** `yv-error.log` no longer records appointment titles or
  category names. YearView compares its two calendar readers on first use, and
  when they disagreed it used to list the differing items by title, in plain
  text, while the page cache beside it is encrypted. It now logs dates, flags
  and the last eight characters of each item's id, which is enough to find
  the item and safe to attach to a bug report.

## Security

- **A hidden category could run script.** The saved category filter is
  written into the page's script block as a string. It was escaped for JSON
  but not for HTML, so a category whose name contained `</script>` ended the
  block and ran whatever followed, once you hid it ("Hide all" was enough).
  The filter now uses the same escaping the per-category titles list has had
  since v3.4.7: `<` and `;` are written as JSON escapes, so the name is
  unchanged but can no longer end the block.
- **An appointment title could blank the view.** Two faults had the same
  effect, and anyone able to send you a meeting invitation could trigger
  either:
  - The page's placeholders were filled by a chain of replacements, each of
    which re-read everything already inserted, calendar data included. A title
    of `__HIDE__` was itself replaced, corrupting the data. The page is now
    filled in a single pass that never re-reads an inserted value.
  - Settings are updated inside already-built pages by finding their
    declaration. The search took the first match anywhere, and the calendar
    data comes before the code, so a title such as `const START_MONTH = ` was
    "updated" instead and the page broke. Declarations are now matched only at
    the start of a line. Calendar text can never supply that, because in the
    page it always sits inside a JSON string, where line breaks are escaped.
    The same applies to the filter seed when cached pages are restored or
    re-stamped.
- **Shared machines.** The Security Design note said that on a PC shared by
  several signed-in accounts, another account could only *view* the
  calendar. That was wrong: the page it can read carries the session token,
  so it could also create, open and delete items. And with the fixed port, a
  second user starting YearView while yours is running would see your
  calendar. The note is corrected, and the README states the requirement:
  YearView is for single-user machines.

## Robustness

- **Uninstall now removes the Edge profile.** Since v3.3.2 the app window has
  run from its own Edge profile inside the install folder, and the MSI
  uninstall left it behind, despite a code comment saying otherwise. The
  uninstall now closes the YearView window, then removes the profile, which
  also holds that window's browsing history. The script installer's
  `-Uninstall` closes the window first too, so its folder removal can't fail
  part-way on files Edge still holds open.
- **"Launch YearView now?" in the script installer.** It handed the launcher
  to `powershell.exe`, which treated the path as a command. That fails when
  the Windows user name contains a space. It now runs the launcher with
  `wscript.exe`, as the shortcuts do.

## Verification

PowerShell 7 was available in the build container for the first time, so the
new server code was run as written, not transcribed:

- The real page-assembly block, settings re-stamper, filter-seed restorer and
  escaping helpers were extracted from the script and run against the real
  client template. The resulting pages were then rendered in a browser DOM.
- Twelve cases were covered:
  - appointment titles equal to `__HIDE__` and `__DATA2__` (with a real spill
    year);
  - titles equal to each seeded declaration, followed by a re-stamp;
  - `</script>` category names in both the filter and the titles list;
  - a hidden category name containing a line break and a declaration.
- In every case:
  - the page rendered;
  - no injected script ran;
  - the data parsed back unchanged;
  - each re-stamp changed only the real declaration.
- Against v3.4.8's logic (checked during the review from a transcription of
  it), the same inputs gave a blank page or ran the injected script.
- All PowerShell parses cleanly: the script, the installer, the build script,
  and the inline commands in each `.cmd` helper.
- The client is unchanged apart from the version stamp. The v3.4.7 and v3.4.8
  simulations (per-category titles, print legend) were re-run and pass.
- The installer round-trips byte-identical. The MSI payload, now including the
  updated uninstall helper, extracts byte-identical. The sequence gate, the
  wipe condition and the new publisher fields were read back from the built
  package.

Not testable here: the Edge-window shutdown during uninstall needs Windows.

## Smoke test

1. Install the MSI over v3.4.8. Settings Apps should list YearView with
   publisher daniellmnz, and the view should behave exactly as before.
2. Optionally, create an appointment titled `__HIDE__` and confirm the view
   still shows. Then delete the appointment.
3. To test the uninstall, open YearView, then uninstall from Settings. The
   window should close, and `%LOCALAPPDATA%\YearView` should be gone.
