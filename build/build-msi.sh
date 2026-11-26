#!/bin/sh
# Builds YearView-<version>.msi from YearView.wxs and src/, using wixl from
# msitools. This is the build the release workflow runs; build-msi.ps1 in the
# repository root is the Windows equivalent and produces the same package.
#
# The version is READ from src/New-OutlookYearView.ps1 and is never passed in,
# so the package can never disagree with the script it carries. Run
# build/verify-msi.py afterwards to check that it does not.
#
# Usage: build/build-msi.sh [output-directory]      (default: dist)
set -eu

cd "$(dirname "$0")/.."

VER=$(sed -n "s/^\$script:yvVersion = '\([0-9][0-9.]*\)'.*/\1/p" src/New-OutlookYearView.ps1 | head -1)
if [ -z "$VER" ]; then
    echo "build-msi: no \$script:yvVersion in src/New-OutlookYearView.ps1" >&2
    exit 1
fi

OUT=${1:-dist}
MSI="$OUT/YearView-$VER.msi"
LOG="$OUT/.wixl.log"
mkdir -p "$OUT"
rm -f "$MSI"

if ! wixl -o "$MSI" YearView.wxs 2>"$LOG"; then
    echo "build-msi: wixl failed" >&2
    cat "$LOG" >&2
    rm -f "$LOG"
    exit 1
fi

# wixl does not implement the Show attribute on a shortcut and says so on
# every run. It needs no action and the shortcuts are correct without it.
# Anything else wixl printed is worth seeing.
grep -v "has no property named 'Show'" "$LOG" | grep -v '^[[:space:]]*$' >&2 || true
rm -f "$LOG"

if [ ! -f "$MSI" ]; then
    echo "build-msi: wixl reported success but produced no package" >&2
    exit 1
fi

( cd "$OUT" && sha256sum "YearView-$VER.msi" > SHA256SUMS.txt )

echo "built  $MSI  ($(stat -c%s "$MSI") bytes)"
cat "$OUT/SHA256SUMS.txt"
