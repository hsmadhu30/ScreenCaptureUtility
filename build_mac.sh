#!/bin/bash
# ============================================================
#  Build "Screen Capture Utility.app" for macOS
#  Run on a Mac (macOS 11 or later), in the folder that holds
#  main.py, icon.ico and this script:
#
#      chmod +x build_mac.sh
#      ./build_mac.sh
#
#  To sign for company-wide use with your Developer ID:
#      SIGN_ID="Developer ID Application: Your Company (TEAMID)" ./build_mac.sh
#  (then notarise - see MAC_README.md)
# ============================================================
set -euo pipefail
cd "$(dirname "$0")"

APP_NAME="Screen Capture Utility"
BUNDLE_ID="${BUNDLE_ID:-com.yourcompany.screencaptureutility}"

# ---- Which Python ----------------------------------------------------
# Apple's own /usr/bin/python3 (from the Command Line Tools) comes with an
# old Tk 8.5 that no longer works on current macOS: an app built with it
# sits in the Dock without ever showing its window. So use the python.org
# Python (preferred) or Homebrew's, never Apple's.
pick_python() {
    if [ -n "${PYTHON:-}" ]; then echo "$PYTHON"; return; fi
    for v in 3.13 3.12 3.11 3.14; do
        p="/Library/Frameworks/Python.framework/Versions/$v/bin/python$v"
        if [ -x "$p" ]; then echo "$p"; return; fi
    done
    for p in /opt/homebrew/bin/python3 /usr/local/bin/python3; do
        if [ -x "$p" ]; then echo "$p"; return; fi
    done
    command -v python3
}
PYTHON="$(pick_python)"

echo "=== Checking Python and Tk ==="
echo "Using: $PYTHON"
TK_OK=$("$PYTHON" - <<'PY' 2>/dev/null || echo "none"
import sys, tkinter
print("Python", sys.version.split()[0], "- Tk", tkinter.TkVersion, "-", sys.executable)
print("ok" if tkinter.TkVersion >= 8.6 and not sys.executable.startswith(("/usr/bin/", "/Library/Developer/CommandLineTools/")) else "old")
PY
)
echo "$TK_OK" | head -1
if [ "$(echo "$TK_OK" | tail -1)" != "ok" ]; then
    echo
    echo "This Python can't build a working Mac app (it's Apple's built-in Python"
    echo "with the old Tk 8.5, or it has no Tk at all)."
    echo "  1. Install Python 3.13 from https://www.python.org/downloads/macos/"
    echo "  2. Run this script again (it finds that Python by itself)."
    exit 1
fi

echo "=== Creating the build environment (venv-mac) ==="
# rebuild the environment if it was made with a different Python
if [ -d venv-mac ] && [ "$(cat venv-mac/.made-with 2>/dev/null)" != "$PYTHON" ]; then
    echo "Removing the old build environment (made with another Python)"
    rm -rf venv-mac
fi
if [ ! -d venv-mac ]; then
    "$PYTHON" -m venv venv-mac
    echo "$PYTHON" > venv-mac/.made-with
fi
source venv-mac/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements-mac.txt

echo "=== Cleaning previous build ==="
rm -rf build dist "$APP_NAME.spec"

ICON_ARGS=()
if [ -f icon.ico ]; then
    ICON_ARGS=(--icon icon.ico)          # PyInstaller converts it to .icns
fi

echo "=== Building $APP_NAME.app ==="
python -m PyInstaller \
    --noconfirm --clean \
    --windowed \
    --name "$APP_NAME" \
    --osx-bundle-identifier "$BUNDLE_ID" \
    --hidden-import docx --collect-data docx \
    --hidden-import openpyxl \
    --hidden-import pptx --collect-data pptx \
    --hidden-import imageio_ffmpeg --collect-all imageio_ffmpeg \
    --hidden-import objc --hidden-import Foundation --hidden-import AppKit \
    --hidden-import Quartz --hidden-import Vision --hidden-import ApplicationServices \
    --exclude-module cv2 --exclude-module numpy \
    --exclude-module pyaudiowpatch --exclude-module winrt \
    ${ICON_ARGS[@]+"${ICON_ARGS[@]}"} \
    main.py

APP="dist/$APP_NAME.app"
PLIST="$APP/Contents/Info.plist"

echo "=== Adding the permission texts macOS shows to users ==="
set_plist() {   # key, type (-string / -bool), value - plutil adds or replaces it
    plutil -replace "$1" "$2" "$3" "$PLIST"
}
set_plist NSMicrophoneUsageDescription -string "Records your voice with screen recordings when Microphone is switched on."
set_plist NSAppleEventsUsageDescription -string "Copies screenshots to the clipboard."
set_plist LSMinimumSystemVersion -string "11.0"
set_plist NSHighResolutionCapable -bool YES

echo "=== Signing ==="
if [ -n "${SIGN_ID:-}" ]; then
    codesign --force --deep --options runtime --timestamp \
        --entitlements entitlements.plist --sign "$SIGN_ID" "$APP"
    echo "Signed with: $SIGN_ID  (next: notarise it, see MAC_README.md)"
else
    codesign --force --deep --sign - "$APP"
    echo "Signed ad hoc (fine for testing on this Mac; for other Macs use SIGN_ID)."
fi
codesign --verify --deep --strict "$APP" && echo "Signature OK"

echo
echo "============================================================"
echo " Done!  Your app is here:"
echo "   $(pwd)/$APP"
echo " If it doesn't open, its log is in:"
echo "   ~/Library/Logs/Screen Capture Utility/app.log"
echo " Zip it for sharing with:"
echo "   ditto -c -k --keepParent \"$APP\" \"dist/$APP_NAME.zip\""
echo "============================================================"
