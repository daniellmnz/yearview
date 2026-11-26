# YearView

A year-at-a-glance view of your Outlook calendar, in its own window.

YearView shows all twelve months at once, with each appointment as a coloured
bar in its category's colour. It reads your calendar directly from Outlook
(classic) on your own PC and keeps itself up to date as the calendar changes.
You can also create, open and delete items from the year view.

## Features

- Twelve months on one screen, with multi-day items drawn as continuous bars.
- Colours taken from your Outlook categories, plus a category legend and
  filter. For each category, you can choose whether to show it at all and
  whether its items show their titles. A title turned off is not written into
  the page at all, so it does not travel in a printed PDF.
- To create an item, click a category and then a day, or drag across several
  days. Click a bar to open it in Outlook, drag it to another day to move
  it, or right-click to delete it.
- Updates itself a few seconds after calendar changes, and switches instantly
  between years you have already viewed.
- The year can start at the current month, or at any month you choose, so a
  view can run across into next year.
- Find items by title or category, combining terms with AND and OR. A matched
  category is outlined in the legend. Plus a Today button and keyboard
  shortcuts.
- Rest the pointer on an item for its full details. Tentative items (such as
  invitations you haven't accepted) are shown hatched, unless you turn that
  off.
- Week can start on Sunday or Monday. Optional item counts in the legend, and
  an optional stripe for an item's second category.
- Page shading: keep the page white, or tint it light, medium or darker grey.
- Prints to a chosen number of pages, with or without the category legend.
- Works with any calendar in your default Outlook mailbox.

## Requirements

- Windows 10 or 11.
- **Outlook (classic)**, the desktop app, running. The new Outlook and Outlook
  on the web are not supported.
- Microsoft Edge, which is installed with Windows.
- Windows PowerShell 5.1, also built into Windows.
- **A single-user PC.** See [Security](#privacy-and-security).

## Installing

Download `YearView-<version>.msi` from the [Releases](../../releases) page
and run it. It installs for your Windows account only, needs no administrator
rights, and puts everything in `%LOCALAPPDATA%\YearView`. Choose the optional
"Start at sign-in" feature if you want YearView ready in the background. To
remove it, use Settings → Apps.

The MSI is not code-signed, so Windows SmartScreen may warn before running it.
Each release includes `SHA256SUMS.txt`; you can check the download against it
before installing:

```powershell
Get-FileHash .\YearView-3.7.4.msi
```

## Using it

Launch **YearView** from the Start Menu or Desktop, with Outlook already
running. Click **?** in the window for quick help, and **⚙** for the settings:

- year start month;
- week start (Sunday or Monday);
- item counts in the legend, and a stripe for a second category;
- hatching for tentative items;
- page shading (white, or light, medium or darker grey);
- confirm before deleting, and before moving;
- print page count;
- print legend.

## Privacy and security

- **Local only.** YearView runs a small web server that listens only on
  `127.0.0.1` (port 8377), so it can't be reached from the network. It
  doesn't send anything anywhere.
- **What it reads.** It reads each calendar item's start, end, subject,
  all-day flag, categories and recurrence. It never touches contacts, the
  address book, attendees or mail.
- **What it stores.** Its files live in `%LOCALAPPDATA%\YearView`: settings,
  window position, an error log, and a cache of the rendered calendar. The
  cache is encrypted with Windows DPAPI, so only your Windows account can read
  it. Uninstalling removes all of it.
- **Single-user PCs only.** Any program on the PC can talk to YearView's local
  server. On a PC shared by several signed-in accounts, such as a Remote
  Desktop server or a family PC using fast user switching, another account
  could read and change the calendar YearView is showing. Don't run it on a
  shared PC.

The details, including the request checks and how calendar text is escaped,
are in [docs/YearView-Security-Design.md](docs/YearView-Security-Design.md).

## Reporting a problem

Please open an issue. If it's about an error, `%LOCALAPPDATA%\YearView\yv-error.log`
helps. The log doesn't record appointment titles or category names, but do
read it before attaching it.

## Building the MSI

On Windows, with the [WiX Toolset v3](https://github.com/wixtoolset/wix3/releases):

```powershell
.\build-msi.ps1                 # builds YearView-<version>.msi from YearView.wxs and src\
```

On Linux, with `msitools` installed:

```sh
./build/build-msi.sh            # builds dist/YearView-<version>.msi and SHA256SUMS.txt
python3 build/verify-msi.py     # checks the package against the sources it came from
```

The version is read from `src/New-OutlookYearView.ps1`, never passed in, so
the package cannot disagree with the script it carries. `verify-msi.py` checks
that it does not, along with the product code, the per-user install, the
payload against `src/` byte for byte, the install sequence, and the shortcuts.

Two builds of the same sources are not byte-identical, and should not be:
Windows Installer requires a fresh package code for each build, and the
package stores a build timestamp. So `verify-msi.py` also prints a canonical
digest, taken over every table and stream with the package code, the build
time and the name of the building tool left out. Two builders that print the
same digest built the same installer. `wixl` 0.103 and 0.106 produce packages
1024 bytes apart from these sources and the same digest; `--dump` prints the
material the digest covers, to find out where two that disagree differ.

### Releases

[`.github/workflows/release.yml`](.github/workflows/release.yml) builds and
verifies on every push to `main`, and keeps the package it built. When the
version in `src/New-OutlookYearView.ps1` has no release yet, it also tags the
commit `V<version>` and publishes the MSI and `SHA256SUMS.txt`. Bumping the
version is all it takes to cut a release; pushing anything else is safe.

Release notes come from `docs/releases/RELEASE-NOTES-v<version>.md` — written
for whoever is installing it. `docs/changes/CHANGES-v<version>.md` is the
technical record of what changed, and is used only as a fallback so that a
release is never published without notes. The checksum block at the end of the
notes is appended by the workflow from the package it built, so it is not
written by hand.

## Documentation

- [Function summary](docs/YearView-Function-Summary.md): what YearView does
  and how.
- [Security design](docs/YearView-Security-Design.md): the threat model and
  protections.
- [Release history](docs/YearView-Release-History.md): how it got here.

## Licence

[MIT](LICENSE)
