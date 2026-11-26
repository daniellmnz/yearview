#!/usr/bin/env python3
"""Checks a built YearView MSI against the sources it was built from.

Run this after build/build-msi.sh. Every check here is a fact that must hold
for the package to be safe to publish, and any failure is fatal. The release
workflow will not create a tag or upload an asset unless this script passes.

Two builds of the same sources are NOT byte-identical, and should not be:
Windows Installer requires a fresh package code for every build, so wixl
generates one, and it stores a build timestamp. Those two fields aside, every
table and the whole payload are reproducible. So this script verifies the
package by what is in it rather than by its hash, which also means a msitools
upgrade cannot quietly change the installer without failing a check.

Usage: build/verify-msi.py [output-directory]      (default: dist)
       build/verify-msi.py --dump [output-directory]

--dump prints the material the canonical digest is taken over, so two builds
whose digests disagree can be diffed to find out where.
"""

import hashlib
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import uuid

ROOT = pathlib.Path(__file__).resolve().parent.parent
ARGS = [a for a in sys.argv[1:] if not a.startswith("--")]
DUMP = "--dump" in sys.argv[1:]
OUT = ROOT / (ARGS[0] if ARGS else "dist")

# The namespace the product code is derived from, as its UUID rather than the
# name it came from. uuid5 is a one-way hash, so this is the identical
# namespace with nothing identifying in it, and every code it derives is
# unchanged. The upgrade code is fixed for the life of the product: changing it
# would stop new versions replacing installed ones.
PRODUCT_NS = uuid.UUID("dc919b73-2afe-5bba-94a2-a4efba27e8fa")
UPGRADE_CODE = "{D18151D9-0EF3-5CA2-B555-2069861B2555}"

# Payload: every file the package installs, and the source it must match.
PAYLOAD = [
    "New-OutlookYearView.ps1",
    "YearView.cmd",
    "Stop-YearView.cmd",
    "YearView.ico",
    "YearView-launch.vbs",
    "Fix-YearViewShortcuts.cmd",
    "Remove-YearViewData.cmd",
]

# Custom actions, and what each one has to sit between. The bridge must be shut
# down before its files are removed or replaced, the data wipe must run before
# the files it reads are gone, and the shortcut repair must run after the
# install has finished. Sequence numbers come from YearView.wxs, so only the
# ORDER is asserted - that is what would actually break.
ORDERING = [
    ("InstallValidate", "RemoveExistingProducts"),
    ("RemoveExistingProducts", "InstallInitialize"),
    ("CostFinalize", "SetCmdExe"),
    ("SetCmdExe", "StopBridgeOnRemove"),
    ("StopBridgeOnRemove", "WipeDataOnRemove"),
    ("WipeDataOnRemove", "RemoveFiles"),
    ("InstallFiles", "StopBridgeOnInstall"),
    ("InstallFinalize", "FixShortcutsOnInstall"),
]

failures = []
checks = 0


def check(ok, label, detail=""):
    global checks
    checks += 1
    if ok:
        print(f"  ok    {label}")
    else:
        print(f"  FAIL  {label}" + (f"\n          {detail}" if detail else ""))
        failures.append(label)


def msi_table(msi, table):
    """Rows of an MSI table as dicts. msiinfo export emits three header lines:
    column names, column types, then the table name."""
    txt = subprocess.run(
        ["msiinfo", "export", str(msi), table],
        capture_output=True, text=True, check=True,
    ).stdout.replace("\r\n", "\n")
    lines = [ln for ln in txt.split("\n") if ln]
    cols = lines[0].split("\t")
    return [dict(zip(cols, ln.split("\t"))) for ln in lines[3:]]


def main():
    # ---- version, from the one place that defines it --------------------
    src_ps1 = ROOT / "src" / "New-OutlookYearView.ps1"
    m = re.search(r"^\$script:yvVersion = '([0-9][0-9.]*)'", src_ps1.read_text(encoding="utf-8"), re.M)
    if not m:
        sys.exit("verify: no $script:yvVersion in src/New-OutlookYearView.ps1")
    ver = m.group(1)
    print(f"version {ver} (from src/New-OutlookYearView.ps1)\n")

    msi = OUT / f"YearView-{ver}.msi"
    if not msi.is_file():
        sys.exit(f"verify: {msi} not found - run build/build-msi.sh first")

    wxs = (ROOT / "YearView.wxs").read_text(encoding="utf-8")
    prop = {r["Property"]: r["Value"] for r in msi_table(msi, "Property")}

    print("identity")
    check(f'Version="{ver}"' in wxs, "YearView.wxs version matches the script",
          "the wxs and the script disagree - a version bump was left half done")
    check(prop.get("ProductVersion") == ver, "MSI ProductVersion matches",
          f'MSI says {prop.get("ProductVersion")!r}')

    want_pc = "{" + str(uuid.uuid5(PRODUCT_NS, "Product-" + ver)).upper() + "}"
    check(prop.get("ProductCode") == want_pc, "ProductCode is the one derived for this version",
          f'expected {want_pc}, MSI has {prop.get("ProductCode")}')
    check(prop.get("UpgradeCode") == UPGRADE_CODE, "UpgradeCode unchanged",
          f'expected {UPGRADE_CODE}, MSI has {prop.get("UpgradeCode")}')

    # ---- per-user, no elevation ----------------------------------------
    print("\nper-user install")
    check("ALLUSERS" not in prop, "no ALLUSERS property",
          "ALLUSERS turns this into a machine install, which needs elevation "
          "and would put the files in AppData for one account only")
    dirs = {r["Directory"]: r.get("Directory_Parent", "") for r in msi_table(msi, "Directory")}
    check(dirs.get("INSTALLFOLDER") == "LocalAppDataFolder",
          "INSTALLFOLDER sits under LocalAppDataFolder",
          f'parent is {dirs.get("INSTALLFOLDER")!r}')
    suminfo = subprocess.run(["msiinfo", "suminfo", str(msi)],
                             capture_output=True, text=True, check=True).stdout
    wc = re.search(r"^Source:\s*(\d+)", suminfo, re.M)
    wc = int(wc.group(1)) if wc else -1
    check(wc & 2 == 2, "package is compressed", f"word count {wc}")
    check(wc & 8 == 8, "package declares elevation not required", f"word count {wc}")

    # ---- the payload is the committed source, byte for byte -------------
    print("\npayload")
    tmp = pathlib.Path(tempfile.mkdtemp())
    try:
        subprocess.run(["msiextract", "-C", str(tmp), str(msi)],
                       capture_output=True, check=True)
        extracted = {p.name: p for p in tmp.rglob("*") if p.is_file()}
        check(sorted(extracted) == sorted(PAYLOAD), "installs exactly the expected files",
              f"extracted {sorted(extracted)}")
        for name in PAYLOAD:
            got, want = extracted.get(name), ROOT / "src" / name
            check(got is not None and got.read_bytes() == want.read_bytes(),
                  f"{name} is the committed source",
                  "the packaged copy differs from src/ - the cab is stale")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # ---- install sequence ----------------------------------------------
    print("\ninstall sequence")
    seq = {r["Action"]: int(r["Sequence"]) for r in msi_table(msi, "InstallExecuteSequence")
           if r.get("Sequence", "").lstrip("-").isdigit()}
    for first, then in ORDERING:
        check(first in seq and then in seq and seq[first] < seq[then],
              f"{first} before {then}",
              f"{first}={seq.get(first)} {then}={seq.get(then)}")

    # ---- shortcuts and registry ----------------------------------------
    # wixl does not implement every shortcut attribute, so the shortcuts are
    # worth reading back rather than assuming. All three launch through the
    # .vbs under wscript, which is what stops a console window appearing, and
    # the sign-in one must carry the background arguments or it would open the
    # calendar at every sign-in.
    print("\nshortcuts and registry")
    sc = {r["Shortcut"]: r for r in msi_table(msi, "Shortcut")}
    check(sorted(sc) == ["ShortcutDesktop", "ShortcutMenu", "ShortcutStartup"],
          "the three shortcuts are present", f"found {sorted(sc)}")
    for name, row in sorted(sc.items()):
        check(row.get("Target", "").lower().endswith("wscript.exe")
              and "YearView-launch.vbs" in row.get("Arguments", ""),
              f"{name} launches through the .vbs under wscript",
              f'target={row.get("Target")!r} args={row.get("Arguments")!r}')
    startup = sc.get("ShortcutStartup", {}).get("Arguments", "")
    for flag in ("-Serve", "-App", "-NoOpen"):
        check(flag in startup, f"sign-in shortcut passes {flag}",
              f"arguments are {startup!r} - it would open a window at sign-in")

    feat = {r["Feature"]: r for r in msi_table(msi, "Feature")}
    check(feat.get("FeatureAutostart", {}).get("Level") == "1000",
          "Start at sign-in is optional, not installed by default",
          f'level is {feat.get("FeatureAutostart", {}).get("Level")!r}')

    # Root 1 is HKCU. A 2 anywhere here is HKLM, which a per-user install
    # cannot write without elevation.
    reg = msi_table(msi, "Registry")
    bad = [r["Key"] for r in reg if r.get("Root") != "1"]
    check(reg and not bad, "every registry write is under HKCU",
          f"not HKCU: {bad}")

    # ---- the checksum file describes this package -----------------------
    print("\nchecksums")
    sums = OUT / "SHA256SUMS.txt"
    if sums.is_file():
        digest = hashlib.sha256(msi.read_bytes()).hexdigest()
        text = sums.read_text()
        check(digest in text and msi.name in text, "SHA256SUMS.txt matches the built package",
              f"{sums.name} does not name {msi.name} with {digest[:16]}...")
    else:
        check(False, "SHA256SUMS.txt exists", f"{sums} not found")

    # ---- canonical digest ----------------------------------------------
    # One number covering everything in the package except the two fields
    # that cannot be reproducible: the package code, which Windows Installer
    # requires to be fresh for every build, and the build timestamp. Two
    # builders who print the same digest built the same installer, even on
    # different wixl versions and with different file sizes.
    print("\ncanonical digest")
    tables = sorted(subprocess.run(["msiinfo", "tables", str(msi)],
                                   capture_output=True, text=True, check=True
                                   ).stdout.replace("\r\n", "\n").split())
    tmp = pathlib.Path(tempfile.mkdtemp())
    try:
        h = hashlib.sha256()
        for table in tables + ["_SummaryInformation"]:
            # Run in a scratch directory: msiinfo writes binary stream tables
            # (Icon) to disk as <Table>/<Table>.<key> beside the caller.
            rows = subprocess.run(["msiinfo", "export", str(msi), table],
                                  capture_output=True, text=True, cwd=tmp
                                  ).stdout.replace("\r\n", "\n").split("\n")
            if table == "_SummaryInformation":
                # 9 is the package code, 12 and 13 the created and saved times.
                rows = [r for r in rows if r.split("\t")[0] not in ("9", "12", "13")]
            h.update(f"[{table}]\n".encode())
            if DUMP:
                print(f"[{table}]")
            for row in sorted(rows):
                h.update(row.encode() + b"\n")
                if DUMP and row:
                    print(f"  {row}")
        # binary streams msiinfo wrote out, by content
        for p in sorted(tmp.rglob("*")):
            if p.is_file():
                rel = p.relative_to(tmp).as_posix()
                d = hashlib.sha256(p.read_bytes()).digest()
                h.update(rel.encode())
                h.update(d)
                if DUMP:
                    print(f"[stream] {rel} {d.hex()[:32]} {p.stat().st_size} bytes")
        print(f"  {h.hexdigest()}")
        print("  (same digest = same installer; it excludes the package code and build time)")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print(f"\n{checks - len(failures)}/{checks} checks passed")
    if failures:
        print("FAILED: " + "; ".join(failures))
        return 1
    print("PASS - package matches its sources and is safe to publish")
    return 0


if __name__ == "__main__":
    sys.exit(main())
