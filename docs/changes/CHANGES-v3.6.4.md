# YearView v3.6.4 — changes from v3.6.3

The confirmation when you move an item is now a setting.

## What changed

- Settings has a new checkbox, **Confirm before moving**, directly below
  "Confirm before deleting". It is on by default, so nothing changes unless
  you untick it.
- With it off, dropping a bar on another day moves the item straight away.
  **Esc** while dragging still cancels.
- The "Moved" message now names the new date. With the prompt off, that
  message is the only place the new date is stated.
- The refusals are unchanged: repeating items, meetings and items changed
  since the page was drawn are still not moved, whatever the setting.
- Like the other settings, it takes effect at once and is remembered.

## Verification (Linux container — no Windows/Outlook execution)

Tested in Chromium on pages built by the real PowerShell:

- With the setting on, the dialog appears and nothing is sent before you
  answer.
- Unticking it saves the setting, and the next drag moves the item with no
  dialog.
- A page that starts with it off moves the item with no dialog.
- Delete still asks.
- The setting is updated correctly inside already-built pages.
- The earlier drag tests pass.
- The MSI payload extracts byte-identical.
