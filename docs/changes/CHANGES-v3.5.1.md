# YearView v3.5.1 — changes from v3.5.0

With **Page shading** set to grey, the grid lines and outlines now darken
steadily with each level, so they stay crisp against the darker pages.

## What changed

In v3.5.0 the lines did get darker as the page got darker, but only by about
the same amount. Measured against the page, their contrast held roughly level
across the four settings, and it was slightly lower at Light grey than on
White. On a grey page, lines with that contrast look fainter than the same
contrast on white.

Each level's colours are now worked out so the contrast rises steadily:

| Level       | Month and label lines | Day lines |
|-------------|-----------------------|-----------|
| White       | 1.40 : 1 (unchanged)  | 1.19 : 1  |
| Light grey  | 1.52 : 1              | 1.26 : 1  |
| Medium grey | 1.67 : 1              | 1.32 : 1  |
| Darker grey | 1.82 : 1              | 1.41 : 1  |

The outline grey used by the header buttons, the legend chips and hollow
(uncategorised) bars now steps down with the lines as well. Lines and outlines
keep their slight blue-grey tint. White looks exactly as before.

## Verification (Linux container — no Windows/Outlook execution)

- **Rendering:** each level was rendered in Chromium. The line colours were
  read back from the page and match the solved values, and changing the
  setting live still re-tints at once and saves.
- **Regression:** the per-category titles and print-legend simulations pass.
- **Checks:** the build gates pass. The MSI payload extracts byte-identical,
  and the sequence gate was read back from the built package.

## Smoke test

Step through the four shading levels. The lines between months and days
should look as clear on Darker grey as they do on White.
