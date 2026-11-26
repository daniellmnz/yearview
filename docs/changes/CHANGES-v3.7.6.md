# YearView v3.7.6 — changes from v3.7.5

**The running version is shown at the foot of the Settings panel.**

## Why

v3.7.5 turned the status line off by default, and that line was the only place
on screen that said which version was running. The version is the first thing
wanted when something is being reported or checked, so it now sits at the foot
of **⚙ Settings**, where it is always reachable and never in the way.

## One version, one place

The number had been written in two places: `$script:yvVersion`, and the page's
stamp string, whose only safeguard was a comment on the first saying to bump
the second. Adding the panel would have made three.

Instead the version is substituted into the page from the single declaration,
as the settings already are:

    $verLit  = ConvertTo-SeedStr ([string]$script:yvVersion)

and the client reads it from one literal:

    const YV_VERSION = __VERSION__;

Both the status line and the panel use that. There is now one place to change
the version, and the comment asking a human to keep two copies in step has
gone with the second copy.

## Testing

* The panel rendered in Chromium: the version is present, reads
  `YearView 3.7.6`, and is the last element in the panel.
* The status line, when on, carries the same version from the same literal.
* The source gates: no unreplaced or orphaned placeholders, so a page can
  never ship with the version substitution missed.
* The v3.7.5 tests re-run unchanged against this build — the status line off
  leaves no text, no tooltip and nothing extractable from a printed PDF; the
  settings file and bridge round trips; the cached-page re-stamp both ways.

## Not changed

The calendar reader is untouched — its signature is identical to v3.7.4, so an
upgraded installation does not re-run first-build validation.
