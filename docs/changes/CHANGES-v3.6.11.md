# YearView v3.6.11 — changes from v3.6.10

The three-second startup validation is no longer repeated after every release.

## What was happening

On its first build of a session YearView reads your calendar **twice** — once
with the fast reader and once with the slow one — and compares them, to prove
the fast reader is telling the truth on your machine. That costs about three
seconds and shows in the stamp as "first-build validation".

The result was already saved to disk so it would not have to be repeated. But
it was filed under the **script version**, so installing any new version threw
it away.

The reader itself has not changed since v3.6.0. Its code is byte-for-byte
identical in 3.6.7, 3.6.8, 3.6.9, 3.6.10 and 3.6.11 — so each of those
installs made you spend three seconds re-proving a result that was already
proven, because the version number beside it had moved.

## What changed

The saved result is now filed under a hash of the **reader's own code** rather
than the version number.

- Change the reader and the result is correctly discarded.
- Change anything else — a label, a setting, the serialiser, the page — and it
  stands.

Everything else about the check is unchanged: it still expires after seven
days so environmental drift is caught, it is still thrown away if the fast
reader ever fails, and the first run on a machine still earns it.

If the signature cannot be worked out for any reason, YearView validates
exactly as it did before. The fallback is always the old behaviour, never a
skipped check.

**On upgrading to this version you will pay the three seconds once more** — the
saved result from v3.6.10 is in the old format and cannot be matched. From then
on it should persist across releases.

The tooltip said "once per session", which was never quite true. It now says
"once, then remembered until the reader changes".

## Verification (Linux container — no Windows/Outlook execution)

The real `Get-ReaderSignature` was extracted from the built script and run:

- **Identical across five releases.** 3.6.7 through 3.6.11 all produce
  `3DED2A016FD9B3EB`, so a result saved now survives all of them.
- **A change to the reader voids it.** An edit inside `Read-YearFast` moved
  the signature to `2F8815D203D1B7EF`.
- **A change outside it does not.** Bumping the version to 9.9.9 left the
  signature untouched — which is the whole point.
- **A v3.6.10 result does not match**, so it re-validates once and is then
  rewritten in the new form.
- **Every failure mode returns nothing**, and so validates as before: missing
  file, unbalanced braces, and an implausibly short extraction were each
  tried.
- **Cost:** 18 ms, once, and only when there is a saved result to check — a
  session with none pays nothing for it. Against three seconds of validation.
- **A build gate was added.** The signature is taken by matching braces in the
  script's own text, which is about five times quicker than parsing the file
  properly. The build now computes it both ways and refuses to package unless
  they agree, so an unbalanced brace inside a reader function is caught before
  release rather than at your startup.
- The reader, year-navigation and render suites pass; the page draws 798 bars
  with no errors and hostile text still displays as plain text.
- All PowerShell parses cleanly; the MSI payload extracts byte-identical; the
  install sequence was read back from the built package.

## Smoke test

Start YearView, let it settle, then close and start it again. The second start
should not show "first-build validation" in the tooltip — and nor should the
first start after the next release.
