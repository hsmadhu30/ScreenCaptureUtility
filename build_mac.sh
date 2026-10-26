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
PYTHON="${PYTHON:-python3}"

echo "=== Checking Python and Tk ==="
"$PYTHON" -c "import sys, tkinter; print('Python', sys.version.split()[0], '- Tk', tkinter.TkVersion)" || {
    echo "Python with Tk is needed. Install Python from python.org (includes Tk),"
    echo "or with Homebrew: brew install python python-tk"
    exit 1
}

echo "=== Creating the build environment (venv-mac) ==="
if [ ! -d venv-mac ]; then
    "$PYTHON" -m venv venv-mac
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
    --noconfirm \
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
echo " Zip it for sharing with:"
echo "   ditto -c -k --keepParent \"$APP\" \"dist/$APP_NAME.zip\""
echo "============================================================"
