# YearView v3.6.1 — changes from v3.6.0

You can now move an item by dragging its bar to another day.

## How it works

- Press on a bar and drag it. The days the item would land on are highlighted
  in green, and the original bar fades.
- Drop it on a day. A dialog names the item with its old and new dates. Nothing
  changes in Outlook until you click **Move**.
- The item moves by a whole number of days. Its time of day and its length
  stay the same. For a multi-day item it doesn't matter where along the bar
  you take hold.
- You can drop onto a different month row.
- **Esc** cancels mid-drag. Dropping back on the same day, or outside the
  calendar, does nothing.
- A plain click still opens the item.

## What it won't move

In each of these cases you get a message saying why, and nothing changes:

- **Repeating items.** Open the item in Outlook instead.
- **Meetings.** Moving one here wouldn't tell the other attendees, so it has
  to be done in Outlook.
- **Items changed since the page was drawn.** Refresh and try again.

There is a confirmation every time. It isn't tied to the "Confirm before
deleting" setting.

## A fix found in testing

After a double-click on a bar, its text was selected. On the next press the
browser then started its own text drag and swallowed the gesture. Bar text can
no longer be selected.

## Verification (Linux container — no Windows/Outlook execution)

**Move code, run under PowerShell 7 against a mock item**

- A timed item, a four-day all-day item and a move across the end of daylight
  saving all land on the right date with time and length intact.
- Invalid day counts are refused, as are repeating items, meetings, stale
  pages, items in another calendar and unknown items. None of them saves
  anything.
- If Outlook doesn't show the new date afterwards, that is reported as a
  failure.

**Page, tested in Chromium**

- A click, and a press with a slight wobble, still open the item.
- A drag highlights the right days, and nothing is sent before confirming.
- **Move** sends the right item and day count, then refreshes. **Cancel**,
  **Esc**, same-day and off-calendar drops send nothing.
- A multi-day item held mid-bar and dropped in another month moves by the
  right amount.
- A repeating item shows the message without a dialog.
- A refusal from the server is shown and the page isn't refreshed.
- The delete dialog and click-to-create are unchanged, and the v3.6.0 tests
  still pass.

**Not tested: the save in real Outlook.** The new date goes in by setting the
item's start and end and saving, which I could only run against a mock. Try
the first move on an unimportant item.

## Smoke test

1. Create a test appointment, drag it two days later and click **Move**.
   Check it in Outlook: same time, new date.
2. Drag it and press **Esc**. Nothing should change.
3. Try dragging a repeating item and a meeting. Each should be refused with a
   message.
