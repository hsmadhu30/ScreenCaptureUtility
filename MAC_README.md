# Screen Capture Utility for Mac

The Mac version is the **same `main.py`** as the Windows version. When it runs on macOS it uses the Mac's own tools instead of the Windows ones. You build it once on a Mac into **`Screen Capture Utility.app`**, then share that app with other Mac users.

> **Important:** this version was built and tested without a Mac, using stand-ins for the macOS tools. Every Mac-specific part ran in that simulation, but it has **not yet run on a real Mac**. Work through the checklist at the end the first time you build it.

---

## 1. What works on a Mac

| Feature | On a Mac | Difference from Windows |
|---|---|---|
| Capture Screen | Yes | Captures the main display (Windows: all monitors) |
| Capture Region | Yes | Retina screens give full-resolution images (twice the size) |
| Full Page (scrolling) | Yes | Needs the Accessibility permission (see section 3) |
| Review window (annotate, blur, crop, text) | Yes | Same |
| Word, Excel, PowerPoint output | Yes | Same |
| Extract Text (OCR) | Yes | Uses Apple's built-in text recognition, on the Mac, with automatic language detection |
| Compare Images | Yes | Same |
| Record Screen / Record Region | Yes | Uses macOS's own screen recorder; saved as .mp4 |
| Microphone in recordings | Yes | Same |
| Computer sound in recordings | **No** | macOS doesn't let apps record system sound without an extra audio driver |
| Keyboard shortcuts | Yes | **Cmd**+Shift+F, R, S, T, C, V, P (Ctrl also works) |
| iPad | **No** | iPadOS doesn't let any app capture other apps' screens |

---

## 2. Build the app (once, on a Mac)

**You need:** a Mac with macOS 11 or later, and **Python 3.13 from [python.org](https://www.python.org/downloads/macos/)** (Homebrew's `python` + `python-tk` also work).
- **Don't use Apple's built-in `python3`** (the one that comes with Xcode or the Command Line Tools). It has an old Tk 8.5 that no longer works on current macOS. An app built with it sits in the Dock without ever opening, and can use enormous amounts of memory. The build script now checks for this and stops with a message.

**Steps:**
1. Put these files in one folder: `main.py`, `icon.ico`, `build_mac.sh`, `requirements-mac.txt`, `entitlements.plist`.
2. Open **Terminal** in that folder and run:
   ```
   bash build_mac.sh
   ```
   The script finds the python.org Python by itself. If an earlier build used another Python, the script rebuilds its environment.
3. The script creates its own Python environment (`venv-mac`), installs the packages, and builds the app.
4. The app is at **`dist/Screen Capture Utility.app`**. To share it, zip it with:
   ```
   ditto -c -k --keepParent "dist/Screen Capture Utility.app" "dist/Screen Capture Utility.zip"
   ```

Build on an Apple-silicon Mac (M1 or later) for Apple-silicon Macs. Build on an Intel Mac for Intel Macs, or build on each.

---

**Before rebuilding,** quit any copy that is stuck. Press Option + Command + Esc, choose **Screen Capture Utility** and click **Force Quit**, or run `killall -9 "Screen Capture Utility"` in Terminal. Then delete the old app from Applications.

---

## 3. First run on each Mac

macOS asks for three permissions. Each is a one-time step per Mac.

| Permission | Needed for | Where to allow it |
|---|---|---|
| **Screen Recording** | All screenshots and recordings | System Settings > Privacy & Security > Screen Recording (called Screen & System Audio Recording on macOS 15) |
| **Accessibility** | Full Page (the app scrolls the page for you) | System Settings > Privacy & Security > Accessibility |
| **Microphone** | Recording your voice | macOS asks the first time you record with Microphone on |

After switching on Screen Recording, **quit and reopen** the app. Until then, screenshots show only the desktop background.

**If macOS says the app "can't be opened"** (it isn't signed with a Developer ID yet), do one of the following:
- **Right-click** the app, choose **Open**, then click **Open** again.
- Or go to System Settings > Privacy & Security, scroll down, and click **Open Anyway**.

---

## 4. Sign and notarise for company-wide use

Apps shared across a company should be signed with the company's **Apple Developer ID**, so macOS opens them without warnings. This needs an Apple Developer Program membership, normally held by IT.

1. Build and sign:
   ```
   SIGN_ID="Developer ID Application: Your Company (TEAMID)" ./build_mac.sh
   ```
2. Zip it and send it to Apple for notarisation:
   ```
   ditto -c -k --keepParent "dist/Screen Capture Utility.app" "dist/Screen Capture Utility.zip"
   xcrun notarytool submit "dist/Screen Capture Utility.zip" \
       --apple-id you@company.com --team-id TEAMID --password APP-SPECIFIC-PASSWORD --wait
   ```
3. Attach the approval to the app, then zip it again for sharing:
   ```
   xcrun stapler staple "dist/Screen Capture Utility.app"
   ditto -c -k --keepParent "dist/Screen Capture Utility.app" "dist/Screen Capture Utility.zip"
   ```

Device management tools (Jamf, Intune, Kandji) can also grant the Screen Recording and Accessibility permissions centrally with a PPPC profile, so users don't have to.

---

## 5. Check on a real Mac (first build)

Because the Mac version hasn't run on a real Mac yet, try each of these once and note anything that misbehaves:

- [ ] The app opens and the window looks right: buttons coloured, text readable.
- [ ] Capture Screen and Capture Region: the screenshot shows your windows, not just the desktop background.
- [ ] Review window: draw, blur, crop, then **Copy** and paste into Mail, Word or Teams.
- [ ] Saving into a Word, Excel and PowerPoint file.
- [ ] Extract Text reads the text correctly.
- [ ] Full Page scrolls a long web page and stitches it. Esc or Stop & Save ends it.
- [ ] Record Region with Microphone on, including **Pause** and **Resume**. The .mp4 plays with sound.
- [ ] Compare Images: Capture region, Open and Paste for both images.
- [ ] Cmd+Shift+R and the other shortcuts.

The parts most likely to need a small fix on a real Mac:
- **Recording:** uses macOS's `screencapture -v`. It's stopped by typing a character to it, with Ctrl+C as a backup.
- **Full Page scrolling:** sends scroll-wheel events.
- **Esc during Full Page:** macOS may only report the Esc key if the app also has the Input Monitoring permission. The Stop button always works.

---

## 6. If the app doesn't open

The app writes a log file every time it starts:

```
~/Library/Logs/Screen Capture Utility/app.log
```

To open it, go to Finder, choose **Go > Go to Folder...**, and paste `~/Library/Logs/Screen Capture Utility`.

- **If the window hasn't appeared after 20 seconds**, the log records exactly where the app is stuck, every 20 seconds.
- **If the app ever uses more than 3 GB of memory**, it records where it was in the log and closes itself, instead of filling up the Mac.
- **Only one copy runs at a time.** Opening the app again brings the running copy to the front.
- **Cmd+Q, the app menu's Quit and the Dock's Quit always work.**

Send the log file to whoever maintains the app. It shows the cause straight away.

