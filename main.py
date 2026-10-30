import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageGrab, ImageOps, ImageFilter, ImageDraw, ImageChops, ImageStat
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass
from io import BytesIO
import threading
import re
import math
import time
import ctypes
import sys
import os
import subprocess
import importlib
import importlib.util
import wave
import asyncio
import concurrent.futures
import tempfile
import signal
from types import SimpleNamespace
import faulthandler


# ============================================================
# Start-up log and safety guard
# ============================================================
# An app opened from the Finder has no Terminal, so on a Mac everything the
# app prints goes to ~/Library/Logs/Screen Capture Utility.log. If start-up
# stalls, the log gets a stack trace of what the app was doing; if start-up
# memory runs away, the app stops itself (instead of filling the Mac's
# memory) and says where the log is.
# "--selftest" starts the app, shows the window for a moment and quits; the
# Mac build script uses it to check the finished app.

APP_NAME = "Screen Capture Utility"
FROZEN = getattr(sys, "frozen", False)
SELFTEST = "--selftest" in sys.argv
STARTUP_MEMORY_LIMIT_MB = 1500        # the window normally starts with ~100 MB
_T0 = time.time()
_started = threading.Event()


def peak_memory_mb():
    """Highest memory use of this process so far (MB)."""
    try:
        import resource
        peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        return peak / (1024 * 1024) if os.uname().sysname == "Darwin" else peak / 1024   # bytes on macOS, KB on Linux
    except Exception:
        try:                                   # Windows
            class PMC(ctypes.Structure):
                _fields_ = [("cb", ctypes.c_ulong), ("PageFaultCount", ctypes.c_ulong),
                            ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t)] + \
                           [(f"f{i}", ctypes.c_size_t) for i in range(6)]
            pmc = PMC(); pmc.cb = ctypes.sizeof(PMC)
            ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.windll.kernel32.GetCurrentProcess(),
                                                     ctypes.byref(pmc), pmc.cb)
            return pmc.PeakWorkingSetSize / (1024 * 1024)
        except Exception:
            return 0.0


def _open_log():
    if sys.platform != "darwin" and not SELFTEST:
        return None, None
    try:
        if sys.platform == "darwin":
            folder = Path.home() / "Library" / "Logs"
        else:
            folder = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / APP_NAME
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / f"{APP_NAME}.log"
        return open(path, "w", encoding="utf-8", buffering=1), path
    except OSError:
        return None, None


LOG_FILE, LOG_PATH = _open_log()


class _Tee:
    """Writes to the log and, when there is one, the Terminal."""

    def __init__(self, *streams):
        self.streams = [s for s in streams if s is not None]

    def write(self, text):
        for s in self.streams:
            try:
                s.write(text)
            except Exception:
                pass
        return len(text)

    def flush(self):
        for s in self.streams:
            try:
                s.flush()
            except Exception:
                pass


if LOG_FILE is not None:
    sys.stdout = _Tee(sys.stdout, LOG_FILE)
    sys.stderr = _Tee(sys.stderr, LOG_FILE)
    faulthandler.enable(LOG_FILE)
    faulthandler.dump_traceback_later(30, exit=False, file=LOG_FILE)   # stalled start-up -> stack trace


def startup_step(name):
    print(f"[start-up {time.time() - _T0:5.1f} s, {peak_memory_mb():5.0f} MB] {name}")


def _tell_user_and_quit(message):
    print(message)
    if LOG_FILE is not None:
        faulthandler.dump_traceback(LOG_FILE, all_threads=True)
        LOG_FILE.flush()
    if sys.platform == "darwin" and not SELFTEST:
        text = message.replace('"', "'") + f"\\n\\nDetails: {LOG_PATH}"
        try:
            subprocess.Popen(["/usr/bin/osascript", "-e",
                              f'display alert "{APP_NAME} could not start" message "{text}" as critical'])
        except Exception:
            pass
    os._exit(70)


def _startup_guard():
    """Stops a start-up whose memory runs away (seen on some Macs)."""
    while not _started.is_set() and time.time() - _T0 < 120:
        used = peak_memory_mb()
        if used > STARTUP_MEMORY_LIMIT_MB:
            _tell_user_and_quit(f"Start-up stopped: it was using {used:.0f} MB of memory "
                                f"(normally about 100 MB).")
        time.sleep(0.25)


threading.Thread(target=_startup_guard, name="startup-guard", daemon=True).start()
startup_step(f"Python {sys.version.split()[0]} on {sys.platform}, Tk {tk.TkVersion}, "
             f"{'app' if FROZEN else 'script'}{', self-test' if SELFTEST else ''}")


# ============================================================
# Optional packages (installed automatically if missing)
# ============================================================
# Screen recording uses ffmpeg (imageio-ffmpeg) and, for sound,
# PyAudioWPatch. OpenCV + NumPy are a fallback for video without sound.
# Adding screenshots to a Word document needs python-docx, to an Excel
# workbook openpyxl, and to a PowerPoint presentation python-pptx.
# If they are not installed in the Python environment that runs
# this script, we install them with pip on first start.

def ensure_package(import_name, pip_name):
    try:
        return importlib.import_module(import_name)
    except ImportError:
        # Inside a packaged .exe, packages are bundled at build time and
        # pip isn't available (sys.executable is the .exe itself).
        if getattr(sys, "frozen", False):
            return None

        print(f"Installing missing package '{pip_name}' (one-time setup)...")
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", pip_name]
            )
            importlib.invalidate_caches()
            return importlib.import_module(import_name)
        except Exception as e:
            print(f"Could not install '{pip_name}': {e}")
            return None


def optional_import(import_name):
    """Import a package if it happens to be installed (never installs it)."""
    try:
        return importlib.import_module(import_name)
    except ImportError:
        return None


# Only used as a fallback for recording when ffmpeg is missing - optional.
np = optional_import("numpy")
cv2 = optional_import("cv2")
docx = ensure_package("docx", "python-docx")
openpyxl = ensure_package("openpyxl", "openpyxl")
pptx = ensure_package("pptx", "python-pptx")


# ============================================================
# Platform (Windows and macOS)
# ============================================================
# The same file runs on Windows and on macOS. The Windows-only parts
# (Win32 calls, Windows OCR, WASAPI audio) have macOS counterparts:
#   screenshots      -> the built-in `screencapture` command
#   clipboard image  -> NSPasteboard (pyobjc), AppleScript as a fallback
#   text extraction  -> Apple's Vision framework (pyobjc)
#   Full Page        -> Quartz scroll-wheel events + Esc key state (pyobjc)
#   recording        -> `screencapture -v` (+ microphone), joined into an
#                       .mp4 by ffmpeg
# macOS asks the user once to allow Screen Recording (all captures) and
# Accessibility (Full Page scrolling) in System Settings.

IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
KEY = "Cmd" if IS_MAC else "Ctrl"          # the shortcut key named in the window


def _mac_tool(name, folder):
    """Full path of a built-in macOS tool (apps opened from Finder get a
    minimal PATH)."""
    path = os.path.join(folder, name)
    return path if os.path.exists(path) else name


SCREENCAPTURE = _mac_tool("screencapture", "/usr/sbin")
OSASCRIPT = _mac_tool("osascript", "/usr/bin")

MAC_PACKAGES = [("objc", "pyobjc-core"),
                ("AppKit", "pyobjc-framework-Cocoa"),
                ("Quartz", "pyobjc-framework-Quartz"),
                ("Vision", "pyobjc-framework-Vision"),
                ("ApplicationServices", "pyobjc-framework-ApplicationServices")]
mac = {}
if IS_MAC:
    for _module, _pip_name in MAC_PACKAGES:
        mac[_module] = ensure_package(_module, _pip_name)


def mac_screen_capture_allowed(ask=False):
    """macOS 10.15+: has the user allowed Screen Recording for this app?
    ask=True shows the system prompt (macOS shows it only once)."""
    quartz = mac.get("Quartz")
    try:
        if quartz is None or quartz.CGPreflightScreenCaptureAccess():
            return True
        if ask:
            quartz.CGRequestScreenCaptureAccess()
        return False
    except Exception:
        return True                              # older macOS: no such check


def mac_accessibility_allowed(ask=False):
    """Needed to send scroll-wheel events (Full Page)."""
    services = mac.get("ApplicationServices")
    try:
        if services is None:
            return False
        if ask:
            return bool(services.AXIsProcessTrustedWithOptions(
                {services.kAXTrustedCheckOptionPrompt: True}))
        return bool(services.AXIsProcessTrusted())
    except Exception:
        return False


def mac_grab(bbox=None):
    """Screenshot with macOS's own `screencapture`. bbox in screen points;
    on a Retina display the image has the full pixel resolution."""
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        command = [SCREENCAPTURE, "-x"]                # -x: no camera sound
        if bbox:
            left, top, right, bottom = (int(round(v)) for v in bbox)
            command += ["-R", f"{left},{top},{max(1, right - left)},{max(1, bottom - top)}"]
        else:
            command += ["-m"]                          # the main display
        subprocess.run(command + [path], check=True, timeout=30, capture_output=True)
        with Image.open(path) as image:
            image.load()
            return image.convert("RGB")
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass


def grab_screen(bbox=None):
    """Screenshot of bbox=(left, top, right, bottom) in screen coordinates,
    or of the whole screen (all monitors on Windows, the main display on a Mac)."""
    if IS_MAC:
        return mac_grab(bbox)
    return ImageGrab.grab(bbox=bbox, all_screens=True)


def mac_copy_image(image):
    """Put an image on the macOS clipboard (PNG + TIFF). True on success."""
    rgb = image.convert("RGB")
    png = BytesIO()
    rgb.save(png, format="PNG")
    data = png.getvalue()
    appkit = mac.get("AppKit")
    if appkit is not None:
        try:
            import Foundation
            tiff = BytesIO()
            rgb.save(tiff, format="TIFF")
            board = appkit.NSPasteboard.generalPasteboard()
            board.clearContents()
            ok = board.setData_forType_(Foundation.NSData.dataWithBytes_length_(data, len(data)),
                                        appkit.NSPasteboardTypePNG)
            board.setData_forType_(Foundation.NSData.dataWithBytes_length_(
                tiff.getvalue(), len(tiff.getvalue())), appkit.NSPasteboardTypeTIFF)
            if ok:
                return True
        except Exception as e:
            print(f"Clipboard (AppKit) failed: {e}")
    # Fallback without pyobjc: AppleScript reads the PNG file
    fd, path = tempfile.mkstemp(suffix=".png")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        script = f'set the clipboard to (read (POSIX file "{path}") as «class PNGf»)'
        return subprocess.run([OSASCRIPT, "-e", script], capture_output=True).returncode == 0
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass


def mac_work_area():
    """(left, top, width, height) of the usable part (without menu bar and
    Dock) of the screen that has the focus, in top-left screen points."""
    appkit = mac["AppKit"]
    screen = appkit.NSScreen.mainScreen() or appkit.NSScreen.screens()[0]
    visible = screen.visibleFrame()
    primary_height = appkit.NSScreen.screens()[0].frame().size.height
    top = primary_height - (visible.origin.y + visible.size.height)
    return (int(visible.origin.x), int(top), int(visible.size.width), int(visible.size.height))


def mac_set_windows_shared(shared):
    """Show/hide this app's windows in screen recordings (macOS)."""
    appkit = mac["AppKit"]
    for window in appkit.NSApplication.sharedApplication().windows():
        window.setSharingType_(1 if shared else 0)     # ReadOnly / None
    return True


def mac_scroll(x, y, notches, move_pointer):
    """Turn the scroll wheel at screen point (x, y). Positive = down."""
    quartz = mac["Quartz"]
    if move_pointer:
        quartz.CGWarpMouseCursorPosition((x, y))
    lines = int(round(-3 * notches))
    while lines:
        step = max(-30, min(30, lines))
        lines -= step
        event = quartz.CGEventCreateScrollWheelEvent(None, quartz.kCGScrollEventUnitLine, 1, step)
        quartz.CGEventSetLocation(event, (x, y))
        quartz.CGEventPost(quartz.kCGHIDEventTap, event)


def mac_pointer():
    quartz = mac["Quartz"]
    point = quartz.CGEventGetLocation(quartz.CGEventCreate(None))
    return (point.x, point.y)


def mac_escape_down():
    quartz = mac["Quartz"]
    return bool(quartz.CGEventSourceKeyState(quartz.kCGEventSourceStateCombinedSessionState, 53))


def run_mac_ocr(image, language_tag=None):
    """Runs on the OCR thread: Apple's Vision text recognition, on the Mac.
    Returns the same dict as run_windows_ocr."""
    import objc
    import Vision
    import Foundation

    def language_name(tag):
        try:
            name = Foundation.NSLocale.currentLocale().localizedStringForLanguageCode_(tag)
            return str(name) if name else tag
        except Exception:
            return tag

    with objc.autorelease_pool():
        prepared = _prepare_for_ocr(image, 4096).convert("RGB")

        request = Vision.VNRecognizeTextRequest.alloc().init()
        request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
        request.setUsesLanguageCorrection_(True)
        try:
            supported, _error = request.supportedRecognitionLanguagesAndReturnError_(None)
            supported = [str(t) for t in (supported or [])]
        except Exception:
            supported = []
        if language_tag:
            request.setRecognitionLanguages_([language_tag])
        elif hasattr(request, "setAutomaticallyDetectsLanguage_"):
            request.setAutomaticallyDetectsLanguage_(True)       # macOS 13+

        def recognize(picture):
            """Vision on a PIL image -> lines in that image's pixels."""
            buffer = BytesIO()
            picture.convert("RGB").save(buffer, format="PNG")
            data = buffer.getvalue()
            handler = Vision.VNImageRequestHandler.alloc().initWithData_options_(
                Foundation.NSData.dataWithBytes_length_(data, len(data)), None)
            ok, error = handler.performRequests_error_([request], None)
            if not ok:
                raise RuntimeError(f"Text recognition failed: {error}")
            width, height = picture.size
            found = []
            for observation in request.results() or []:
                candidates = observation.topCandidates_(1)
                if not candidates:
                    continue
                box = observation.boundingBox()           # 0..1, origin bottom-left
                rect = SimpleNamespace(x=box.origin.x * width,
                                       y=(1 - box.origin.y - box.size.height) * height,
                                       width=box.size.width * width,
                                       height=box.size.height * height)
                text = str(candidates[0].string())
                found.append(SimpleNamespace(text=text,
                                             words=[SimpleNamespace(text=text, bounding_rect=rect)]))
            found.sort(key=lambda line: (round(line.words[0].bounding_rect.y / 8),
                                         line.words[0].bounding_rect.x))
            return found

        lines = recognize(prepared)
        try:
            _reread_lines(lines, prepared, recognize)
        except Exception as e:
            print(f"Re-reading lines skipped: {e}")

    languages = [("Automatic", "")] + [(language_name(t), t) for t in supported]
    try:
        _repair_symbols(lines, prepared)       # bullets, *, ", ) and 0 checked against the image
    except Exception as e:
        print(f"Symbol check skipped: {e}")

    return {
        "text": _layout_text(lines),
        "lines": lines,
        "language": language_name(language_tag) if language_tag else "Automatic",
        "language_tag": language_tag or "",
        "languages": languages,
    }


def run_ocr(image, language_tag=None):
    if IS_MAC:
        return run_mac_ocr(image, language_tag)
    return run_windows_ocr(image, language_tag)


def mirror_control_keys(widget):
    """On a Mac, make every Ctrl shortcut of a window work with Cmd too."""
    if not IS_MAC:
        return
    for sequence in widget.bind():
        if "Control-" in sequence:
            widget.bind(sequence.replace("Control-", "Command-"), widget.bind(sequence))


class MacScreenRecorder:
    """macOS screen recording with the built-in `screencapture -v` (it uses
    macOS's own screen capture, so it needs the Screen Recording permission;
    with the microphone switched on it records the default input too).
    Pause ends the current piece and Resume starts a new one; on Stop the
    pieces are joined and converted to an H.264/AAC .mp4 with ffmpeg.
    Same interface as ScreenRecorder."""

    def __init__(self, filename, area, fps=None, record_mic=False, record_system=False):
        self.filename = filename
        self.area = dict(area)
        self.record_mic = record_mic
        self.notes = ["computer sound isn't available on a Mac"] if record_system else []
        self.error = None
        self.warning = None
        self.parts = []
        self.proc = None
        self.lock = threading.Lock()
        self.start_time = None
        self.paused = False
        self.pause_started = None
        self.paused_total = 0.0
        self.recorded_seconds = 0.0
        self.running = False

    # ---- clock (same as ScreenRecorder) ------------------------------

    def elapsed(self):
        with self.lock:
            if self.start_time is None:
                return 0
            now = self.pause_started if self.paused else time.perf_counter()
            return now - self.start_time - self.paused_total

    def duration(self):
        return self.recorded_seconds

    def is_running(self):
        """False if screencapture stopped by itself (e.g. no permission)."""
        if not self.running:
            return False
        return self.paused or (self.proc is not None and self.proc.poll() is None)

    # ---- pieces ------------------------------------------------------

    def _start_part(self):
        path = hidden_temp_path(self.filename, f".part{len(self.parts) + 1}.mov")
        a = self.area
        command = [SCREENCAPTURE, "-v", "-x", "-C",
                   "-R", f"{int(a['left'])},{int(a['top'])},{int(a['width'])},{int(a['height'])}"]
        if self.record_mic:
            command.append("-g")                      # default microphone
        self.proc = subprocess.Popen(command + [str(path)], stdin=subprocess.PIPE,
                                     stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        self.parts.append(path)

    def _stop_part(self):
        proc, self.proc = self.proc, None
        if proc is None:
            return
        if proc.poll() is not None:                   # it stopped by itself: an error
            message = proc.stderr.read().decode(errors="replace").strip() if proc.stderr else ""
            if message and not self.error:
                self.error = f"screencapture stopped: {message}"
            return
        try:                                          # "type any character to stop"
            proc.stdin.write(b"q\n")
            proc.stdin.flush()
            proc.wait(timeout=3)
        except Exception:
            try:
                proc.send_signal(signal.SIGINT)
                proc.wait(timeout=5)
            except Exception:
                proc.terminate()

    # ---- control -----------------------------------------------------

    def start(self):
        with self.lock:
            self.start_time = time.perf_counter()
        self._start_part()
        self.running = True

    def pause(self):
        with self.lock:
            if self.paused or self.start_time is None:
                return
            self.paused = True
            self.pause_started = time.perf_counter()
        self._stop_part()

    def resume(self):
        with self.lock:
            if not self.paused:
                return
            self.paused_total += time.perf_counter() - self.pause_started
            self.paused = False
            self.pause_started = None
        self._start_part()

    def stop(self):
        self.recorded_seconds = self.elapsed()
        self._stop_part()
        self.running = False

    def finalize(self):
        parts = [p for p in self.parts if p.exists() and p.stat().st_size > 0]
        try:
            if not parts:
                if not self.error:
                    self.error = ("No video was recorded. Allow Screen Recording for Screen Capture "
                                  "Utility in System Settings > Privacy & Security, then try again.")
                return
            if FFMPEG_EXE is None:                     # keep macOS's own .mov
                target = self.filename.with_suffix(".mov")
                os.replace(parts[0], target)
                self.filename = target
                if len(parts) > 1:
                    self.warning = "only the part before the first pause was kept (ffmpeg is missing)"
                return
            listing = hidden_temp_path(self.filename, ".parts.txt")
            listing.write_text("".join(f"file '{p.as_posix()}'\n" for p in parts), encoding="utf-8")
            command = [FFMPEG_EXE, "-y", "-loglevel", "error", "-f", "concat", "-safe", "0",
                       "-i", str(listing),
                       "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
                       "-c:v", "libx264", "-preset", "veryfast", "-crf", "23", "-pix_fmt", "yuv420p"]
            command += (["-c:a", "aac", "-b:a", "192k"] if self.record_mic else ["-an"])
            command += ["-movflags", "+faststart", str(self.filename)]
            result = subprocess.run(command, capture_output=True)
            try:
                listing.unlink()
            except OSError:
                pass
            if result.returncode != 0:
                details = result.stderr.decode(errors="replace").strip()
                print(f"Converting the recording failed: {details}")
                target = self.filename.with_suffix(".mov")
                os.replace(parts[0], target)
                self.filename = target
                self.warning = "saved as .mov (converting to .mp4 failed, see console)"
        except Exception as e:
            self.error = f"Could not save the recording: {e}"
        finally:
            for path in self.parts:
                if path != self.filename:
                    try:
                        path.unlink()
                    except OSError:
                        pass

    def audio_summary(self):
        parts = ["with microphone"] if self.record_mic else []
        parts += self.notes
        if self.warning:
            parts.append(self.warning)
        return "; ".join(parts)



# ffmpeg (bundled with imageio-ffmpeg) encodes the video as H.264 and
# combines it with the recorded audio into one .mp4 file.
imageio_ffmpeg = ensure_package("imageio_ffmpeg", "imageio-ffmpeg")

# PyAudioWPatch records the microphone and the computer's own sound
# ("what you hear", via WASAPI loopback). It is Windows-only.
pyaudio = (ensure_package("pyaudiowpatch", "PyAudioWPatch")
           if sys.platform == "win32" else None)

FFMPEG_EXE = None
if imageio_ffmpeg is not None:
    try:
        FFMPEG_EXE = imageio_ffmpeg.get_ffmpeg_exe()
    except Exception as e:
        print(f"ffmpeg not available: {e}")

# On a Mac, recording uses macOS's own `screencapture` (ffmpeg converts it
# to .mp4) and records the microphone; the computer's own sound needs
# Windows (WASAPI loopback).
RECORDING_AVAILABLE = IS_MAC or FFMPEG_EXE is not None or (np is not None and cv2 is not None)
AUDIO_AVAILABLE = FFMPEG_EXE is not None and (pyaudio is not None or IS_MAC)
SYSTEM_AUDIO_AVAILABLE = AUDIO_AVAILABLE and not IS_MAC
DOCX_AVAILABLE = docx is not None
XLSX_AVAILABLE = openpyxl is not None
PPTX_AVAILABLE = pptx is not None
OFFICE_AVAILABLE = DOCX_AVAILABLE or XLSX_AVAILABLE or PPTX_AVAILABLE

if DOCX_AVAILABLE:
    from docx import Document
    from docx.shared import Emu, Inches, Pt

if not RECORDING_AVAILABLE:
    print("Screen recording is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install imageio-ffmpeg')

if not AUDIO_AVAILABLE:
    print("Audio recording is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install imageio-ffmpeg' + ("" if IS_MAC else " PyAudioWPatch"))

if not DOCX_AVAILABLE:
    print("Saving to Word documents is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install python-docx')

if not XLSX_AVAILABLE:
    print("Saving to Excel workbooks is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install openpyxl')

if not PPTX_AVAILABLE:
    print("Saving to PowerPoint presentations is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install python-pptx')

# Text extraction (OCR) uses the OCR engine built into Windows 10/11 via
# PyWinRT. We only check that it's installed here; it is imported later
# on its own thread (see "Text Extraction").
OCR_PACKAGES = [
    "winrt-runtime",
    "winrt-Windows.Foundation",
    "winrt-Windows.Foundation.Collections",
    "winrt-Windows.Globalization",
    "winrt-Windows.Graphics.Imaging",
    "winrt-Windows.Media.Ocr",
    "winrt-Windows.Storage.Streams",
]


def ocr_installed():
    try:
        return importlib.util.find_spec("winrt.windows.media.ocr") is not None
    except Exception:
        return False


OCR_AVAILABLE = False
if sys.platform == "win32":
    if not ocr_installed() and not getattr(sys, "frozen", False):
        print("Installing text extraction (OCR) support (one-time setup)...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + OCR_PACKAGES)
            importlib.invalidate_caches()
        except Exception as e:
            print(f"Could not install OCR support: {e}")
    OCR_AVAILABLE = ocr_installed()
elif IS_MAC:
    OCR_AVAILABLE = mac.get("Vision") is not None and mac.get("objc") is not None

if not OCR_AVAILABLE:
    print("Text extraction (OCR) is disabled (needs Windows 10/11 and PyWinRT, "
          "or macOS 10.15+ and pyobjc-framework-Vision).")

# Run ffmpeg without flashing a console window (matters for the .exe)
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


# ============================================================
# Windows DPI awareness
# ============================================================
# Without this, on displays with scaling (125%, 150% ...) the
# selected region and the recorded area don't line up with
# what you actually see on screen.

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# ============================================================
# Settings
# ============================================================

RECORDING_FPS = 15

# Audio is kept aligned with the video clock; if it drifts further
# than this (seconds), silence is inserted or samples are dropped.
AUDIO_MAX_DRIFT = 0.08

# How often (ms) to retry writing to the Word document while it is
# open in Word and therefore locked.
WORD_RETRY_MS = 3000

# Screen pixels -> Word size (EMU) at 96 DPI. Used so small region
# captures are not blown up larger than their real size.
EMU_PER_PIXEL = 9525


# ============================================================
# Save Folder (chosen by the user)
# ============================================================

save_folder = None  # Path object, set when the user picks a folder


def browse_folder():
    """Open a folder picker and put the chosen path in the text box."""
    folder = filedialog.askdirectory(
        parent=root,
        title="Choose a folder to save screenshots and recordings",
        mustexist=True
    )

    if folder:
        folder_var.set(folder)
        set_save_folder(folder)


def set_save_folder(folder_text):
    """Validate the folder path and remember it. Returns True if usable."""
    global save_folder

    folder_text = folder_text.strip().strip('"')

    if not folder_text:
        return False

    folder = Path(folder_text).expanduser()

    if not folder.exists():
        create = messagebox.askyesno(
            "Folder not found",
            f"The folder does not exist:\n\n{folder}\n\nCreate it?",
            parent=root
        )
        if not create:
            return False
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Could not create folder:\n{e}",
                parent=root
            )
            return False

    if not folder.is_dir():
        messagebox.showerror(
            "Invalid folder",
            f"This path is not a folder:\n\n{folder}",
            parent=root
        )
        return False

    save_folder = folder
    folder_var.set(str(folder))
    status_var.set(f"Saving to: {folder}")
    print(f"Files will be saved to: {folder}")
    return True


def ensure_save_folder():
    """
    Make sure a valid save folder has been provided before capturing.
    Uses the typed path if present, otherwise prompts with a folder picker.
    This runs while the main window is still visible, so the dialog
    is never hidden behind a withdrawn window.
    """
    typed = folder_var.get()

    if typed.strip():
        return set_save_folder(typed)

    messagebox.showinfo(
        "Choose a folder",
        "Please choose the folder where files should be saved.",
        parent=root
    )
    browse_folder()
    return save_folder is not None


# ============================================================
# Generate Filenames
# ============================================================

def timestamp(when=None):
    return (when or datetime.now()).strftime("%Y-%m-%d_%H-%M-%S")


def unique_path(folder, stem, suffix):
    """folder/stem.suffix, adding _2, _3 ... if that name is already taken
    (e.g. two captures within the same second)."""
    path = folder / f"{stem}{suffix}"
    n = 2
    while path.exists():
        path = folder / f"{stem}_{n}{suffix}"
        n += 1
    return path


def get_screenshot_filename(taken_at=None):
    folder = save_folder or (word_doc_path.parent if word_doc_path else Path.cwd())
    return unique_path(folder, f"Screenshot_{timestamp(taken_at)}", ".png")


def get_recording_filename():
    return unique_path(save_folder, f"Recording_{timestamp()}", ".mp4")


# ============================================================
# Save Screenshot as a standalone .png file
# ============================================================

def save_screenshot(image, taken_at=None, update_status=True):
    filename = get_screenshot_filename(taken_at)

    try:
        image.save(filename, format="PNG")
        print("Screenshot saved successfully:")
        print(filename)
        if update_status:
            status_var.set(f"Saved: {filename.name}")

    except Exception as e:
        print(f"Error saving screenshot: {e}")
        status_var.set("Error saving screenshot (see console)")


# ============================================================
# Document Output (Word / Excel / PowerPoint)
# ============================================================
# When a document is selected, every screenshot is added to it and the
# file is saved straight away, always in the order the screenshots were
# taken (earliest first):
#   Word (.docx)        - images one after another, 3 blank lines apart
#   Excel (.xlsx)       - images stacked down a "Screenshots" sheet,
#                         3 empty rows apart
#   PowerPoint (.pptx)  - one slide per screenshot, fitted to the slide
#
# If the file is open in Word / Excel / PowerPoint, Windows locks it and
# it can't be written. In that case captures wait in a queue (in capture
# order) and are written automatically as soon as the file is closed.

@dataclass
class Capture:
    taken_at: datetime
    kind: str          # "Full screen" or "Region"
    png: bytes         # the image, PNG-encoded
    size: tuple        # (width, height) in pixels


word_doc_path = None     # Path of the selected .docx / .xlsx / .pptx, or None
pending_captures = []    # captures not yet written to the document
retry_job = None         # id of the scheduled auto-retry, if any


def make_capture(image, kind, taken_at):
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return Capture(taken_at=taken_at, kind=kind,
                   png=buffer.getvalue(), size=image.size)


def _length(value, default):
    return value if value is not None else default


BLANK_LINES_BETWEEN_IMAGES = 3


def document_has_content(doc):
    """True if the document already has text, images or tables."""
    if doc.tables:
        return True
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            return True
        if paragraph._p.xpath(".//w:drawing") or paragraph._p.xpath(".//w:pict"):
            return True
    return False


def append_captures_to_docx(path, captures):
    """
    Open the .docx, add each capture (oldest first) at the end of the
    document with a caption, and save it. The file is written to a
    temporary file first and then swapped in, so the document can't be
    left half-written.
    """
    path = Path(path)
    doc = Document(str(path)) if path.exists() else Document()

    # Usable area of the page (inside the margins)
    section = doc.sections[-1]
    page_w = _length(section.page_width, Inches(8.5))
    page_h = _length(section.page_height, Inches(11))
    usable_w = (page_w
                - _length(section.left_margin, Inches(1))
                - _length(section.right_margin, Inches(1)))
    usable_h = (page_h
                - _length(section.top_margin, Inches(1))
                - _length(section.bottom_margin, Inches(1)))
    max_h = int(usable_h * 0.95)

    # Images only (no captions), oldest capture first
    for capture in sorted(captures, key=lambda c: c.taken_at):
        # 3 blank lines between images (none before the very first one)
        if document_has_content(doc):
            for _ in range(BLANK_LINES_BETWEEN_IMAGES):
                doc.add_paragraph()

        # Fit the image to the page width (and height), keeping its
        # proportions, but never enlarge it beyond its real size.
        w_px, h_px = capture.size
        width = min(usable_w, w_px * EMU_PER_PIXEL)
        fitted_h = width * h_px / w_px

        if fitted_h > max_h and max_h / fitted_h < 0.6:
            # Very tall (e.g. a full-page capture): keep it readable by
            # splitting it into page-sized slices at full width
            slice_px = max(1, int(max_h * w_px / width))
            with Image.open(BytesIO(capture.png)) as tall:
                for top_px in range(0, h_px, slice_px):
                    part = tall.crop((0, top_px, w_px, min(h_px, top_px + slice_px)))
                    buffer = BytesIO()
                    part.save(buffer, format="PNG")
                    buffer.seek(0)
                    doc.add_picture(buffer, width=Emu(int(width)))
            continue

        if fitted_h > max_h:
            width = max_h * w_px / h_px

        doc.add_picture(BytesIO(capture.png), width=Emu(int(width)))

    temp_path = path.with_name(f".{path.stem}.saving.tmp")
    doc.save(str(temp_path))
    try:
        os.replace(temp_path, path)
    except Exception:
        try:
            temp_path.unlink()
        except Exception:
            pass
        raise


# ---- Excel ----------------------------------------------------------

EXCEL_SHEET = "Screenshots"
EXCEL_MAX_WIDTH_PX = 1600         # wider screenshots are scaled down to this
EXCEL_GAP_ROWS = 3                # empty rows between screenshots
EXCEL_DEFAULT_ROW_PT = 15         # Excel's default row height (points)


def _excel_row_px(ws, row):
    """Height of a worksheet row (1-based) in screen pixels."""
    dim = ws.row_dimensions.get(row) if hasattr(ws.row_dimensions, "get") else None
    height = getattr(dim, "height", None) if dim is not None else None
    if not height:
        height = ws.sheet_format.defaultRowHeight or EXCEL_DEFAULT_ROW_PT
    return float(height) * 96 / 72


def _excel_rows_for(ws, first_row, height_px):
    """How many rows (starting at first_row, 1-based) an image covers."""
    rows, covered = 0, 0.0
    while covered < height_px - 0.5:
        covered += _excel_row_px(ws, first_row + rows)
        rows += 1
    return max(1, rows)


def _excel_next_free_row(ws):
    """First row (1-based) below everything already on the sheet."""
    last = 0
    if ws.max_row > 1 or ws.cell(1, 1).value is not None:
        last = max(last, max((c.row for row in ws.iter_rows() for c in row
                              if c.value is not None), default=0))
    for img in getattr(ws, "_images", []):
        anchor = img.anchor
        start = getattr(getattr(anchor, "_from", None), "row", None)
        if start is None:
            continue
        end = getattr(getattr(anchor, "to", None), "row", None)   # two-cell anchor
        if end is None:
            # displayed height (the anchor's size; img.height is the raw image size)
            ext = getattr(anchor, "ext", None)
            height = ext.height / EMU_PER_PIXEL if ext is not None and ext.height else img.height
            end = start + _excel_rows_for(ws, start + 1, height) - 1
        last = max(last, end + 1)
    return last + 1


def _excel_sheet(wb):
    """The sheet screenshots go on: "Screenshots", or a single empty sheet."""
    if EXCEL_SHEET in wb.sheetnames:
        return wb[EXCEL_SHEET]
    if len(wb.worksheets) == 1:
        ws = wb.worksheets[0]
        empty = ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None
        if empty and not getattr(ws, "_images", []):
            ws.title = EXCEL_SHEET
            return ws
    return wb.create_sheet(EXCEL_SHEET)


def new_excel_workbook():
    wb = openpyxl.Workbook()
    wb.active.title = EXCEL_SHEET
    return wb


def append_captures_to_xlsx(path, captures):
    """Stack the screenshots (oldest first) down the Screenshots sheet."""
    from openpyxl.drawing.image import Image as XLImage
    path = Path(path)
    wb = openpyxl.load_workbook(str(path)) if path.exists() else new_excel_workbook()
    ws = _excel_sheet(wb)
    row = _excel_next_free_row(ws)
    if row > 1:
        row += EXCEL_GAP_ROWS
    keep = []                                   # buffers must live until saved
    for capture in sorted(captures, key=lambda c: c.taken_at):
        w_px, h_px = capture.size
        buffer = BytesIO(capture.png)
        keep.append(buffer)
        img = XLImage(buffer)
        if w_px > EXCEL_MAX_WIDTH_PX:
            img.width = EXCEL_MAX_WIDTH_PX
            img.height = round(h_px * EXCEL_MAX_WIDTH_PX / w_px)
        else:
            img.width, img.height = w_px, h_px
        ws.add_image(img, f"A{row}")
        row += _excel_rows_for(ws, row, img.height) + EXCEL_GAP_ROWS
    wb.active = wb.worksheets.index(ws)
    _save_atomically(path, lambda temp: wb.save(str(temp)))


# ---- PowerPoint -----------------------------------------------------

PPT_MARGIN_IN = 0.2               # space around the screenshot on each slide


def new_presentation():
    from pptx.util import Inches as PInches
    prs = pptx.Presentation()
    prs.slide_width = PInches(13.333)        # 16:9 widescreen
    prs.slide_height = PInches(7.5)
    return prs


def _blank_layout(prs):
    """The 'Blank' slide layout (or the one with the fewest placeholders)."""
    layouts = list(prs.slide_layouts)
    for layout in layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(layouts, key=lambda l: len(l.placeholders))


def _add_picture_slide(prs, png_bytes, w_px, h_px, max_w, max_h):
    slide = prs.slides.add_slide(_blank_layout(prs))
    for ph in list(slide.placeholders):        # keep the slide clean
        ph._element.getparent().remove(ph._element)
    scale = min(max_w / w_px, max_h / h_px, EMU_PER_PIXEL)   # never enlarge
    width, height = int(w_px * scale), int(h_px * scale)
    left = (prs.slide_width - width) // 2
    top = (prs.slide_height - height) // 2
    slide.shapes.add_picture(BytesIO(png_bytes), left, top, width, height)


def append_captures_to_pptx(path, captures):
    """Add one slide per screenshot (oldest first). Very tall screenshots
    (full-page captures) are split over several slides so they stay readable."""
    from pptx.util import Inches as PInches
    path = Path(path)
    prs = pptx.Presentation(str(path)) if path.exists() else new_presentation()
    margin = PInches(PPT_MARGIN_IN)
    max_w = prs.slide_width - 2 * margin
    max_h = prs.slide_height - 2 * margin
    for capture in sorted(captures, key=lambda c: c.taken_at):
        w_px, h_px = capture.size
        fit_width = min(max_w / w_px, EMU_PER_PIXEL)
        if h_px * fit_width > max_h and (max_h / h_px) < fit_width * 0.6:
            slice_px = max(1, int(max_h / fit_width))
            with Image.open(BytesIO(capture.png)) as tall:
                for top_px in range(0, h_px, slice_px):
                    part = tall.crop((0, top_px, w_px, min(h_px, top_px + slice_px)))
                    buffer = BytesIO()
                    part.save(buffer, format="PNG")
                    _add_picture_slide(prs, buffer.getvalue(), part.width, part.height, max_w, max_h)
            continue
        _add_picture_slide(prs, capture.png, w_px, h_px, max_w, max_h)
    _save_atomically(path, lambda temp: prs.save(str(temp)))


# ---- shared -----------------------------------------------------------

DOC_TYPES = {   # extension -> (app name, file type label)
    ".docx": ("Word", "Word document"),
    ".xlsx": ("Excel", "Excel workbook"),
    ".pptx": ("PowerPoint", "PowerPoint presentation"),
}


def doc_app(path=None):
    """'Word', 'Excel' or 'PowerPoint' for the selected (or given) file."""
    path = path if path is not None else word_doc_path
    if path is None:
        return "Word"
    return DOC_TYPES.get(Path(path).suffix.lower(), ("Word",))[0]


def _save_atomically(path, write):
    """Write to a temporary file first, then swap it in, so the file can't
    be left half-written."""
    temp_path = path.with_name(f".{path.stem}.saving.tmp")
    write(temp_path)
    try:
        os.replace(temp_path, path)
    except Exception:
        try:
            temp_path.unlink()
        except Exception:
            pass
        raise


def append_captures(path, captures):
    """Add captures to the selected Word / Excel / PowerPoint file."""
    suffix = Path(path).suffix.lower()
    if suffix == ".xlsx":
        append_captures_to_xlsx(path, captures)
    elif suffix == ".pptx":
        append_captures_to_pptx(path, captures)
    else:
        append_captures_to_docx(path, captures)


def is_file_locked(path):
    """True if another program (Word, Excel, PowerPoint...) has the file open."""
    if not path.exists():
        return False
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def flush_pending(interactive=False):
    """
    Write any waiting captures to the selected document.
    Returns True if nothing is left waiting.
    """
    if not pending_captures or word_doc_path is None:
        update_word_controls()
        return True

    name = word_doc_path.name

    if is_file_locked(word_doc_path):
        report_document_locked()
        return False

    count = len(pending_captures)

    try:
        append_captures(word_doc_path, pending_captures)

    except PermissionError:
        report_document_locked()
        return False

    except Exception as e:
        print(f"Error updating {doc_app()} file: {e}")
        status_var.set(f"Could not update {name} (see console)")
        update_word_controls()
        if interactive:
            messagebox.showerror(
                f"Could not update {doc_app()} file",
                f"{word_doc_path}\n\n{e}",
                parent=root
            )
        return False

    pending_captures.clear()
    noun = "image" if count == 1 else "images"
    print(f"Added {count} {noun} to {word_doc_path}")
    status_var.set(f"Added {count} {noun} to {name} and saved it")
    update_word_controls()
    return True


def report_document_locked():
    n = len(pending_captures)
    noun = "capture" if n == 1 else "captures"
    status_var.set(
        f"{word_doc_path.name} is open in {doc_app()} - {n} {noun} waiting. "
        f"Close it there and they will be added automatically."
    )
    update_word_controls()
    schedule_retry()


def schedule_retry():
    global retry_job
    if retry_job is None:
        retry_job = root.after(WORD_RETRY_MS, auto_retry)


def auto_retry():
    global retry_job
    retry_job = None
    if pending_captures:
        flush_pending()


def save_pending_as_png():
    """Fallback: write waiting captures out as separate .png files."""
    saved = []
    for capture in sorted(pending_captures, key=lambda c: c.taken_at):
        filename = get_screenshot_filename(capture.taken_at)
        try:
            filename.write_bytes(capture.png)
            saved.append(filename)
        except Exception as e:
            print(f"Error saving {filename}: {e}")
    pending_captures.clear()
    return saved


def handle_capture(image, kind, taken_at):
    """Route a new screenshot to the selected document or to a .png file."""
    if word_doc_path is not None:
        pending_captures.append(make_capture(image, kind, taken_at))

        if keep_png_var.get():
            save_screenshot(image, taken_at, update_status=False)

        flush_pending()
    else:
        save_screenshot(image, taken_at)


def can_change_document():
    """Don't switch documents while captures are still waiting."""
    if not pending_captures or flush_pending(interactive=True):
        return True

    messagebox.showwarning(
        "Captures still waiting",
        f"{len(pending_captures)} capture(s) are still waiting to be added to:\n\n"
        f"{word_doc_path}\n\n"
        f"Close that file in {doc_app()} so they can be saved, then try again.",
        parent=root
    )
    return False


def document_start_folder():
    if word_doc_path is not None:
        return str(word_doc_path.parent)
    typed = folder_var.get().strip().strip('"')
    if typed and Path(typed).is_dir():
        return typed
    return str(save_folder) if save_folder else str(Path.home())


def _available_types():
    """File-dialog types for the document formats that can be used."""
    types = []
    if DOCX_AVAILABLE:
        types.append(("Word documents", "*.docx"))
    if XLSX_AVAILABLE:
        types.append(("Excel workbooks", "*.xlsx"))
    if PPTX_AVAILABLE:
        types.append(("PowerPoint presentations", "*.pptx"))
    return types


def _check_document(path):
    """Open the file with the matching library to make sure it can be used.
    Returns an error message, or None if it's fine."""
    suffix = path.suffix.lower()
    old = {".doc": ("Word", ".docx", "Word Document (*.docx)"),
           ".xls": ("Excel", ".xlsx", "Excel Workbook (*.xlsx)"),
           ".xlsm": ("Excel", ".xlsx", "Excel Workbook (*.xlsx)"),
           ".ppt": ("PowerPoint", ".pptx", "PowerPoint Presentation (*.pptx)")}
    if suffix in old:
        app, new, label = old[suffix]
        return (f"Only {new} files are supported, not {suffix}. "
                f"Open it in {app} and use Save As > {label}.")
    if suffix not in DOC_TYPES:
        return "Please choose a Word (.docx), Excel (.xlsx) or PowerPoint (.pptx) file."
    try:
        if suffix == ".docx":
            Document(str(path))
        elif suffix == ".xlsx":
            openpyxl.load_workbook(str(path))
        else:
            pptx.Presentation(str(path))
    except Exception as e:
        return f"{e}\n\nThe file could not be opened as a {DOC_TYPES[suffix][1]}."
    return None


def select_word_document():
    """Select an existing Word, Excel or PowerPoint file."""
    types = _available_types()
    if not types or not can_change_document():
        return

    patterns = " ".join(t[1] for t in types)
    path = filedialog.askopenfilename(
        parent=root,
        title="Select the Word, Excel or PowerPoint file to add screenshots to",
        initialdir=document_start_folder(),
        filetypes=[("Word, Excel or PowerPoint", patterns)] + types
    )
    if not path:
        return

    path = Path(path)

    if path.name.startswith("~$"):
        messagebox.showerror(
            "Not a document",
            "That is a temporary lock file Office creates while a file "
            "is open. Please select the real file.",
            parent=root
        )
        return

    problem = _check_document(path)
    if problem:
        messagebox.showerror("Can't use this file", f"{path}\n\n{problem}", parent=root)
        return

    set_word_document(path)


def new_document_menu():
    """'New' button: choose Word, Excel or PowerPoint."""
    menu = tk.Menu(root, tearoff=0, font=F_BODY)
    menu.add_command(label="Word document (.docx)", command=lambda: new_word_document(".docx"),
                     state="normal" if DOCX_AVAILABLE else "disabled")
    menu.add_command(label="Excel workbook (.xlsx)", command=lambda: new_word_document(".xlsx"),
                     state="normal" if XLSX_AVAILABLE else "disabled")
    menu.add_command(label="PowerPoint presentation (.pptx)", command=lambda: new_word_document(".pptx"),
                     state="normal" if PPTX_AVAILABLE else "disabled")
    x = new_doc_button.winfo_rootx()
    y = new_doc_button.winfo_rooty() + new_doc_button.winfo_height()
    try:
        menu.tk_popup(x, y)
    finally:
        menu.grab_release()


def new_word_document(suffix=".docx"):
    """Create a new, empty Word / Excel / PowerPoint file for screenshots."""
    if not can_change_document():
        return
    app, label = DOC_TYPES[suffix]

    path = filedialog.asksaveasfilename(
        parent=root,
        title=f"Create a new {label} for screenshots",
        initialdir=document_start_folder(),
        initialfile=f"Screen Captures {datetime.now():%Y-%m-%d}{suffix}",
        defaultextension=suffix,
        filetypes=[(f"{label}s", f"*{suffix}")]
    )
    if not path:
        return

    path = Path(path)
    if path.suffix.lower() != suffix:
        path = path.with_name(path.name + suffix)

    try:
        if suffix == ".docx":
            Document().save(str(path))           # empty - images only, no heading
        elif suffix == ".xlsx":
            new_excel_workbook().save(str(path))
        else:
            new_presentation().save(str(path))   # no slides yet
    except Exception as e:
        messagebox.showerror(
            "Could not create file",
            f"{path}\n\n{e}",
            parent=root
        )
        return

    set_word_document(path)


def set_word_document(path):
    global word_doc_path

    word_doc_path = path
    doc_var.set(str(path))

    # Use the document's folder as the save folder if none is set yet
    if not folder_var.get().strip():
        set_save_folder(str(path.parent))

    status_var.set(f"Screenshots will be added to: {path.name}")
    print(f"Screenshots will be added to {DOC_TYPES[path.suffix.lower()][1]}: {path}")
    update_word_controls()


def clear_word_document():
    global word_doc_path

    if word_doc_path is None or not can_change_document():
        return

    word_doc_path = None
    doc_var.set("")
    status_var.set("No document selected - screenshots will be saved as .png files")
    update_word_controls()


def open_word_document():
    if word_doc_path is None:
        return
    try:
        os.startfile(str(word_doc_path))  # Windows: open in Word / Excel / PowerPoint
    except Exception as e:
        messagebox.showerror("Could not open document", str(e), parent=root)


def update_word_controls():
    """Enable/disable the document buttons to match the current state."""
    has_doc = word_doc_path is not None
    state = "normal" if has_doc else "disabled"

    clear_doc_button.config(state=state)
    open_doc_button.config(state=state)
    keep_png_check.config(state=state)

    n = len(pending_captures)
    app = doc_app()
    if n:
        save_pending_button.config(text=f"Save to {app} now ({n} waiting)",
                                   state="normal")
    else:
        save_pending_button.config(text=f"Save to {app} now", state="disabled")


# ============================================================
# Text Extraction (OCR)
# ============================================================
# Uses the OCR engine built into Windows 10/11 (Windows.Media.Ocr):
# offline, fast, nothing extra to install, and it reads the languages
# installed in Windows. Python talks to it through PyWinRT.
#
# All WinRT calls run on one dedicated background thread. (Importing
# WinRT on the Tk thread could interfere with the Windows file dialogs.)

OCR_WORKER = concurrent.futures.ThreadPoolExecutor(max_workers=1,
                                                   thread_name_prefix="ocr")


OCR_TARGET_HEIGHT = 40      # px per line of text: the size OCR engines read most accurately


def _ink_row_runs(gray):
    """Heights of the bands of rows that contain text (ink differs from the
    most common shade, the background)."""
    hist = gray.histogram()
    background = max(range(256), key=hist.__getitem__)
    ink = gray.point(lambda v: 255 if abs(v - background) > 60 else 0)
    w, h = ink.size
    m = ink
    while w > 1:                                   # fold columns: any ink in the row?
        half = (w + 1) // 2
        m = ImageChops.lighter(m.crop((0, 0, half, h)), m.crop((w - half, 0, w, h)))
        w = half
    rows = [v > 0 for v in m.tobytes()]
    runs, start, gap = [], None, 0
    for y, busy in enumerate(rows + [False, False]):
        if busy:
            if start is None:
                start = y
            gap = 0
        elif start is not None:
            gap += 1
            if gap > 1:                            # allow 1 blank row inside a line (e.g. "i" dots)
                runs.append(y - gap + 1 - start)
                start, gap = None, 0
    return [r for r in runs if r >= 4]


def _text_height(gray):
    """Typical height (px) of a line of text, or None if there's no text."""
    runs = sorted(_ink_row_runs(gray))
    if not runs:
        return None
    return runs[len(runs) // 2]


def _dark_bands_to_light(gray):
    """Turn light-on-dark text (dark mode, dark title bars and headers) into
    dark-on-light, band by band, so a dark header above a light page is
    read as well as the page."""
    w, h = gray.size
    if w < 2 or h < 2:
        return gray
    row_means = list(gray.resize((1, h), Image.BOX).tobytes())
    dark = [m < 110 for m in row_means]
    # smooth: a band must be at least 3 rows tall
    bands, start = [], 0
    for y in range(1, h + 1):
        if y == h or dark[y] != dark[start]:
            bands.append((start, y, dark[start]))
            start = y
    if not any(d for _, _, d in bands):
        return gray
    if all(d for _, _, d in bands):
        return ImageOps.invert(gray)
    out = gray.copy()
    for y0, y1, is_dark in bands:
        if is_dark and y1 - y0 >= 3:
            band = ImageOps.invert(gray.crop((0, y0, w, y1)))
            out.paste(ImageOps.autocontrast(band, cutoff=1), (0, y0))   # white background, like the page
    return out


def _prepare_for_ocr(image, max_dimension):
    """Image prepared for the most accurate OCR:
    - grey, with light-on-dark text (dark mode) turned into dark-on-light,
    - full contrast,
    - enlarged so a line of text is about OCR_TARGET_HEIGHT pixels tall
      (small screen text is what makes symbols like _ # @ { } come out
      wrong), sharpened slightly after enlarging,
    - with a white margin around the text.
    Returned as RGBA (what Windows OCR expects)."""
    gray = _dark_bands_to_light(image.convert("L"))
    gray = ImageOps.autocontrast(gray, cutoff=1)

    height = _text_height(gray)
    scale = OCR_TARGET_HEIGHT / height if height else 2.0
    scale = max(1.0, min(4.0, scale))
    margin = 16
    limit = max_dimension or 4000
    scale = min(scale, (limit - 2 * margin) / max(gray.size))
    if abs(scale - 1.0) > 0.02:
        gray = gray.resize((max(1, round(gray.width * scale)), max(1, round(gray.height * scale))),
                           Image.LANCZOS)
        gray = gray.filter(ImageFilter.UnsharpMask(radius=1.2, percent=70, threshold=2))
    gray = ImageOps.expand(gray, border=margin, fill=255)
    return gray.convert("RGBA")


def _is_cjk(text):
    return any("぀" <= ch <= "ヿ" or "㐀" <= ch <= "鿿" or "가" <= ch <= "힯"
               for ch in text)


BULLET_LOOKALIKES = set("•·∙●◦○▪■*¢«»©®°eoc")


def _layout_text(lines, keep_layout=True):
    """Rebuild the text from the OCR's words and their positions.

    - Words on the same visual row are joined even when the OCR returned
      them as separate lines (table cells, columns), in left-to-right order.
    - keep_layout: indentation is kept and wide gaps (table columns) become
      several spaces, so code and tables keep their shape; otherwise words
      are joined with single spaces.
    - A blank line marks a clear vertical gap (a new paragraph).
    - A small round mark at the start of a row is a bullet: written as •."""
    words = []
    for index, line in enumerate(lines):
        for word in getattr(line, "words", None) or []:
            r = word.bounding_rect
            if word.text.strip():
                words.append(SimpleNamespace(text=word.text, x=r.x, y=r.y, w=r.width, h=r.height,
                                             line=index))
    if not words:
        return "\n".join(line.text for line in lines).strip()

    heights = sorted(w.h for w in words)
    line_h = heights[len(heights) // 2] or 1
    # average character width, from all words of 3+ characters together
    # (steadier than any one word, so indentation lands on the right column)
    long_words = [w for w in words if len(w.text) >= 3] or words
    char_w = max(1.0, sum(w.w for w in long_words) / max(1, sum(len(w.text) for w in long_words)))

    # rows: the OCR's own lines, merged when they sit on the same visual row
    groups = {}
    for w in words:
        groups.setdefault(w.line, []).append(w)
    boxes = []
    for ws in groups.values():
        y0 = min(w.y for w in ws)
        y1 = max(w.y + w.h for w in ws)
        boxes.append(SimpleNamespace(words=ws, y0=y0, y1=y1, mid=(y0 + y1) / 2))
    boxes.sort(key=lambda b: (b.mid, min(w.x for w in b.words)))
    rows = []
    for b in boxes:
        row = rows[-1] if rows else None
        if row is not None and abs(b.mid - row.mid) < 0.5 * line_h:
            overlap = any(not (w.x + w.w <= o.x or o.x + o.w <= w.x) for w in b.words for o in row.words)
            if not overlap:
                row.words += b.words
                row.y0, row.y1 = min(row.y0, b.y0), max(row.y1, b.y1)
                row.mid = (row.y0 + row.y1) / 2
                continue
        rows.append(SimpleNamespace(words=list(b.words), y0=b.y0, y1=b.y1, mid=b.mid))

    left = min(w.x for w in words)
    # indentation: rows starting within ~half a character of each other share
    # one indent (glyph edges differ slightly, e.g. "r" vs "{")
    starts = sorted({(min(w.x for w in row.words) - left) / char_w for row in rows})
    indent_of, cluster = {}, []
    for value in starts + [None]:
        if cluster and (value is None or value - cluster[0] > 0.6):
            for v in cluster:
                indent_of[v] = round(cluster[0])
            cluster = []
        if value is not None:
            cluster.append(value)
    # paragraph gaps: clearly more space than between the lines of a paragraph
    pitches = sorted(b.mid - a.mid for a, b in zip(rows, rows[1:]))
    pitch = pitches[len(pitches) // 2] if len(pitches) >= 3 else None
    out = []
    previous = None
    for row in rows:
        ws = sorted(row.words, key=lambda w: w.x)
        if previous is not None and row.y0 - previous.y1 > line_h * 1.2 and (
                pitch is None or row.mid - previous.mid > 1.4 * pitch):
            out.append("")                                    # paragraph gap
        # a bullet: a small round mark, centred on the row, before the first word
        first = ws[0]
        if (len(ws) > 1 and len(first.text) == 1 and first.text in BULLET_LOOKALIKES
                and first.h < 0.6 * line_h and 0.5 < first.w / max(1, first.h) < 1.8
                and abs((first.y + first.h / 2) - row.mid) < 0.2 * line_h
                and ws[1].x - (first.x + first.w) > 0.6 * char_w):
            first.text = "•"
        row_chars = sum(len(w.text) for w in ws)
        row_char_w = (sum(w.w for w in ws) / row_chars) if row_chars >= 8 else char_w
        parts = []
        if keep_layout:
            parts.append(" " * min(80, indent_of[(min(w.x for w in row.words) - left) / char_w]))
        end = None
        for w in ws:
            if end is not None:
                gap = w.x - end
                if keep_layout and gap >= 2.2 * row_char_w:
                    # a column: start it where it sits on screen, so columns line up
                    column = round((w.x - left) / char_w)
                    parts.append(" " * max(2, column - len("".join(parts))))
                elif gap < 0.35 * row_char_w and (_is_cjk(w.text) or _is_cjk(parts[-1])):
                    pass                                      # CJK text has no spaces
                elif gap < 0.25 * row_char_w or (gap < 1.0 * row_char_w and (
                        w.text[0] in ".,;:)]}" or parts[-1][-1:] in ("(", "[", "{"))):
                    pass                                      # one word split in two: root.mainloop(), print("x")
                else:
                    parts.append(" ")
            parts.append(w.text)
            end = w.x + w.w
        out.append("".join(parts).rstrip())
        previous = row
    return _clean_ocr_text("\n".join(out).strip("\n"))


OCR_REPLACEMENTS = {"“": '"', "”": '"', "„": '"', "″": '"',
                    "‘": "'", "’": "'", "‚": "'", "′": "'",
                    "ﬁ": "fi", "ﬂ": "fl", "ﬀ": "ff", "ﬃ": "ffi", "ﬄ": "ffl",
                    " ": " ", "‐": "-", "‑": "-", "−": "-"}


def _clean_ocr_text(text):
    """Undo typographic look-alikes OCR engines invent: curly quotes become
    the plain " and ' that screens (code, URLs, data) almost always show,
    ligatures become separate letters, odd hyphens and spaces become plain ones."""
    for bad, good in OCR_REPLACEMENTS.items():
        text = text.replace(bad, good)
    # two apostrophes against a word are a double quote read as two marks: ''Name'' -> "Name"
    text = re.sub(r"(?<![\w'])''(?=\w)|(?<=\w)''(?![\w'])", '"', text)
    return text


OCR_INK = 140          # darker than this (in the prepared, dark-on-light image) is ink
_INK_LUT = [1 if v < OCR_INK else 0 for v in range(256)]


def _components(gray, box):
    """8-connected ink blobs inside box (x0, y0, x1, y1) of the prepared image.
    Each: x0, y0, x1, y1 (exclusive), n (ink pixels), touches (set of crop sides)."""
    x0, y0, x1, y1 = (int(round(v)) for v in box)
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(gray.width, x1), min(gray.height, y1)
    w, h = x1 - x0, y1 - y0
    if w <= 0 or h <= 0:
        return []
    ink = bytearray(gray.crop((x0, y0, x1, y1)).point(_INK_LUT).tobytes())   # 1 = ink
    blobs = []
    start = ink.find(1)
    while start != -1:
        next_start = start
        ink[start] = 0
        stack = [start]
        bx0 = bx1 = start % w
        by0 = by1 = start // w
        n = 0
        while stack:
            i = stack.pop()
            n += 1
            x, y = i % w, i // w
            if x < bx0: bx0 = x
            if x > bx1: bx1 = x
            if y < by0: by0 = y
            if y > by1: by1 = y
            for dy in (-1, 0, 1):
                yy = y + dy
                if yy < 0 or yy >= h:
                    continue
                for dx in (-1, 0, 1):
                    xx = x + dx
                    if 0 <= xx < w:
                        j = yy * w + xx
                        if ink[j]:
                            ink[j] = 0
                            stack.append(j)
        touches = set()
        if bx0 == 0: touches.add("left")
        if bx1 == w - 1: touches.add("right")
        if by0 == 0: touches.add("top")
        if by1 == h - 1: touches.add("bottom")
        blobs.append(SimpleNamespace(x0=x0 + bx0, y0=y0 + by0, x1=x0 + bx1 + 1, y1=y0 + by1 + 1,
                                     n=n, touches=touches))
        start = ink.find(1, next_start + 1)
    return blobs


def _glyph_groups(blobs, min_size=2):
    """Blobs that overlap horizontally belong to one glyph (the dot of an i,
    an accent), left to right."""
    groups = []
    for b in sorted((b for b in blobs if b.n >= min_size), key=lambda b: b.x0):
        g = groups[-1] if groups else None
        if g is not None and b.x0 < g.x1 - 1:
            g.parts.append(b)
            g.x0, g.y0, g.x1, g.y1 = min(g.x0, b.x0), min(g.y0, b.y0), max(g.x1, b.x1), max(g.y1, b.y1)
            g.n += b.n
            g.touches |= b.touches
        else:
            groups.append(SimpleNamespace(parts=[b], x0=b.x0, y0=b.y0, x1=b.x1, y1=b.y1, n=b.n,
                                          touches=set(b.touches)))
    return groups


def _line_glyphs(gray, region):
    """Glyphs of one line inside region. Ink cut by the region's top or
    bottom edge belongs to the lines above or below (tightly spaced code),
    so it's left out before glyphs are put together."""
    blobs = [b for b in _components(gray, region) if not (b.touches & {"top", "bottom"})]
    return _glyph_groups(blobs)


def _has_hole(gray, g):
    """True if the glyph encloses background (o, °, ◦), False for * • ▪."""
    w, h = g.x1 - g.x0, g.y1 - g.y0
    if w < 3 or h < 3:
        return False
    data = gray.crop((g.x0, g.y0, g.x1, g.y1)).tobytes()
    bg = bytearray(1 if v >= OCR_INK else 0 for v in data)
    stack = [i for i in range(w * h) if bg[i] and (i % w in (0, w - 1) or i // w in (0, h - 1))]
    for i in stack:
        bg[i] = 0
    while stack:
        i = stack.pop()
        x, y = i % w, i // w
        for j in (i - 1 if x > 0 else -1, i + 1 if x < w - 1 else -1, i - w if y > 0 else -1, i + w if y < h - 1 else -1):
            if j >= 0 and bg[j]:
                bg[j] = 0
                stack.append(j)
    return sum(bg) >= 2


def _classify_symbol(gray, groups, top, base):
    """What one glyph is, from its shape and where it sits on the line
    (top = top of capital letters, base = baseline): "•", "◦", "▪", "*",
    '"', "'", or None. Word-shaped glyphs (letters, digits) give None.
    groups: one glyph group, or two for a double quote."""
    H = max(4.0, base - top)
    if len(groups) == 2:
        a, b = groups
        if all(_is_tick(g, top, H) for g in groups) and b.x0 - a.x1 < 0.7 * H \
                and abs(a.y0 - b.y0) < 0.15 * H:
            return '"'
        return None
    g = groups[0]
    w, h = g.x1 - g.x0, g.y1 - g.y0
    fill = g.n / float(max(1, w * h))
    rel_top, rel_bottom = (g.y0 - top) / H, (g.y1 - top) / H
    rel_mid = (rel_top + rel_bottom) / 2
    aspect = w / float(max(1, h))
    if _is_tick(g, top, H):
        return "'"
    if 0.18 * H <= h <= 0.7 * H and 0.7 <= aspect <= 1.45 and 0.3 <= rel_mid <= 0.85:
        hole = _has_hole(gray, g)
        corners = gray.crop((g.x0, g.y0, g.x1, g.y1))
        cw, ch = corners.size
        inked = sum(corners.getpixel(p) < OCR_INK for p in ((0, 0), (cw - 1, 0), (0, ch - 1), (cw - 1, ch - 1)))
        if fill >= 0.8 and inked >= 3 and not hole:
            return "▪"
        if fill >= 0.6 and not hole:
            return "•"
        if hole and 0.3 <= fill < 0.75:
            return "◦"
    if (0.25 * H <= h <= 0.7 * H and 0.5 <= aspect <= 1.5 and rel_top < 0.2 and rel_bottom < 0.66
            and 0.2 <= fill <= 0.65 and not _has_hole(gray, g)):
        return "*"
    return None


def _ink_spans(gray, g):
    """Leftmost and rightmost ink of each row of a glyph (None for a blank row)."""
    w, h = g.x1 - g.x0, g.y1 - g.y0
    data = gray.crop((g.x0, g.y0, g.x1, g.y1)).tobytes()
    spans = []
    for y in range(h):
        row = data[y * w:(y + 1) * w]
        xs = [x for x in range(w) if row[x] < OCR_INK]
        spans.append((xs[0], xs[-1]) if xs else None)
    return spans


def _classify_punct(gray, g, top, base):
    """Small punctuation OCR engines drop at the end of a line or a word:
    ) ( ] [ . , : ; - from the glyph's shape and position. None if unsure."""
    H = max(4.0, base - top)
    w, h = g.x1 - g.x0, g.y1 - g.y0
    rel_top, rel_bottom = (g.y0 - top) / H, (g.y1 - top) / H
    fill = g.n / float(max(1, w * h))

    def dot_like(p):
        pw, ph = p.x1 - p.x0, p.y1 - p.y0
        return 0.07 * H <= ph <= 0.26 * H and 0.55 <= pw / float(max(1, ph)) <= 1.8

    parts = getattr(g, "parts", [g])
    if len(parts) == 2:
        upper, lower = sorted(parts, key=lambda p: p.y0)
        if dot_like(upper) and 0.25 <= ((upper.y0 + upper.y1) / 2 - top) / H <= 0.7:
            if dot_like(lower) and 0.82 <= (lower.y1 - top) / H <= 1.15:
                return ":"
            if (lower.y1 - top) / H >= 1.05 and (lower.y1 - lower.y0) <= 0.5 * H:
                return ";"
        return None
    if len(parts) != 1:
        return None
    if dot_like(g) and 0.82 <= rel_bottom <= 1.15 and fill >= 0.5:
        return "."
    if 0.15 * H <= h <= 0.45 * H and w <= 0.3 * H and rel_top >= 0.62 and rel_bottom >= 1.05:
        return ","
    if h <= 0.16 * H and 1.6 <= w / float(max(1, h)) and w <= 1.0 * H and 0.4 <= (rel_top + rel_bottom) / 2 <= 0.8:
        return "-"
    if 0.85 * H <= h <= 1.7 * H and w <= 0.5 * H and w <= 0.45 * h and rel_top <= 0.15 and rel_bottom >= 0.98:
        spans = _ink_spans(gray, g)
        n = len(spans)
        band = max(1, n // 7)

        def centre(rows):
            rows = [r for r in rows if r]
            return sum((a + b) / 2 for a, b in rows) / len(rows) if rows else None

        def width(rows):
            rows = [r for r in rows if r]
            return max(b - a + 1 for a, b in rows) if rows else 0

        top_rows, mid_rows, bot_rows = spans[:band], spans[n // 2 - band // 2:n // 2 + band // 2 + 1], spans[-band:]
        edge_rows = max(1, n // 16)                     # a bracket's flat top and bottom
        c_top, c_mid, c_bot = centre(top_rows), centre(mid_rows), centre(bot_rows)
        if None in (c_top, c_mid, c_bot):
            return None
        # a square bracket has one straight side; a parenthesis is curved
        body = [r for r in spans[n // 5:n - n // 5] if r]
        lefts, rights = [a for a, _ in body], [b for _, b in body]
        straight_left = body and max(lefts) - min(lefts) <= max(1, 0.12 * w)
        straight_right = body and max(rights) - min(rights) <= max(1, 0.12 * w)
        if (straight_left or straight_right) and width(spans[:edge_rows]) >= 0.6 * w \
                and width(spans[-edge_rows:]) >= 0.6 * w and width(mid_rows) <= 0.6 * w and w >= 3:
            return "[" if straight_left else "]"
        bulge = c_mid - (c_top + c_bot) / 2
        if bulge >= max(1.0, 0.18 * w):
            return ")"
        if bulge <= -max(1.0, 0.18 * w):
            return "("
    return None


def _is_tick(g, top, H):
    """A quote mark: a short upright stroke at the top of the line."""
    w, h = g.x1 - g.x0, g.y1 - g.y0
    return (0.15 * H <= h <= 0.5 * H and h >= 1.3 * w and w <= 0.2 * H
            and (g.y0 - top) / H < 0.2 and (g.y1 - top) / H < 0.6)


BULLETS = ("•", "◦", "▪")
PUNCTUATION = (")", "(", "]", "[", ".", ",", ":", ";", "-")
QUOTE_LOOKALIKES = {"''", "``", "“", "”", "„", "″", "'", "`", "‘", "’", "*", "°", "^", "-"}
STAR_LOOKALIKES = {"x", "X", "×", "+", "^", "°", "'", '"', "-", "#", "%", "¥"}
SYMBOL_LOOKALIKES = set("•·∙●◦○▪■*¢«»©®°eoaOcCx×+'\"`^-—~#%")


_TALL_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZbdfhkl0123456789")
_LOW_CHARS = set("gjpqy()[]{}|,;_$@Q")


def _ink_baseline(gray, words, top):
    """Baseline from the ink: the lowest row that still has a good share of
    the line's ink (descenders - g p y, brackets - have much less)."""
    x0 = int(min(w.x for w in words)); x1 = int(max(w.x + w.w for w in words))
    y1 = int(max(w.y + w.h for w in words))
    band = gray.crop((x0, int(top), x1, y1))
    w, h = band.size
    data = band.tobytes()
    counts = [sum(1 for v in data[y * w:(y + 1) * w] if v < OCR_INK) for y in range(h)]
    if not counts or max(counts) == 0:
        return y1
    peak = max(counts)
    last = max(y for y, c in enumerate(counts) if c >= 0.3 * peak)
    return int(top) + last + 1


def _line_metrics(words, gray=None):
    """Top of capitals and baseline of a line of words. The words' own text
    says which boxes to trust: the top from words with capitals, digits or
    tall letters (b d f h k l), the baseline from words with nothing below
    the line (no g j p q y, brackets or commas). Medians, so one odd box
    doesn't skew them."""
    usable = [w for w in words if len(w.text) >= 2] or words
    tall = [w for w in usable if any(c in _TALL_CHARS for c in w.text)] or usable
    flat = [w for w in usable if not any(c in _LOW_CHARS for c in w.text)]
    tops = sorted(w.y for w in tall)
    top = tops[len(tops) // 2]
    if gray is not None:                 # the ink is the most reliable guide
        base = _ink_baseline(gray, words, top)
    elif flat:
        bottoms = sorted(w.y + w.h for w in flat)
        base = bottoms[len(bottoms) // 2]
    else:
        bottoms = sorted(w.y + w.h for w in usable)
        base = bottoms[len(bottoms) // 2]
    if base - top < 4:
        return min(w.y for w in words), max(w.y + w.h for w in words)
    return top, base


def _ink_lines(gray):
    """Lines of ink in the prepared image: (top, bottom, left, right)."""
    hist = gray.histogram()
    background = max(range(256), key=hist.__getitem__)
    mask = gray.point(lambda v: 255 if abs(v - background) > 60 else 0)
    w, h = mask.size
    folded, fw = mask, w
    while fw > 1:                                  # fold columns: any ink in the row?
        half = (fw + 1) // 2
        folded = ImageChops.lighter(folded.crop((0, 0, half, h)), folded.crop((fw - half, 0, fw, h)))
        fw = half
    rows = folded.tobytes()
    spans, start, gap = [], None, 0
    for y in range(h + 2):
        busy = y < h and rows[y] > 0
        if busy:
            if start is None:
                start = y
            gap = 0
        elif start is not None:
            gap += 1
            if gap > 2:
                y1 = y - gap + 1
                box = mask.crop((0, start, w, y1)).getbbox()
                if box and y1 - start >= 4:
                    spans.append((start, y1, box[0], box[2]))
                start, gap = None, 0
    return spans


def _reread_lines(lines, gray, recognize, limit=15):
    """OCR engines (Windows OCR in particular) sometimes skip a whole line,
    or stop before the end of one (a closing ") or ..." left off). Each line
    of ink in the image that no word covers, or that has ink well past its
    first or last word, is read again on its own - a single line is read
    much more reliably - and the better reading is kept.
    recognize(image) -> lines in that image's coordinates."""
    import difflib
    gray = gray.convert("L")
    spans = _ink_lines(gray)
    if not spans:
        return lines
    heights = sorted(y1 - y0 for y0, y1, _, _ in spans)
    typical = heights[len(heights) // 2]
    words = [(line, word) for line in lines for word in (getattr(line, "words", None) or []) if word.text.strip()]
    done = 0
    for y0, y1, ink_left, ink_right in spans:
        bh = y1 - y0
        if bh < 0.5 * typical or bh > 2.5 * typical or done >= limit:
            continue
        mine = [(line, word) for line, word in words
                if y0 - 0.25 * bh <= word.bounding_rect.y + word.bounding_rect.height / 2 <= y1 + 0.25 * bh]
        if mine:
            left = min(word.bounding_rect.x for _, word in mine)
            right = max(word.bounding_rect.x + word.bounding_rect.width for _, word in mine)
            if right >= ink_right - 0.6 * bh and left <= ink_left + 0.6 * bh:
                continue                                       # the words cover the line
        done += 1
        # just this line (the lines above and below would only confuse the engine)
        top = max(0, y0 - 2)
        left_x = max(0, int(ink_left - bh))
        strip = gray.crop((left_x, top, min(gray.width, int(ink_right + bh)), min(gray.height, y1 + 2)))
        margin = max(16, int(0.6 * bh))
        canvas = Image.new("L", (strip.width + 2 * margin, strip.height + 2 * margin), 255)
        canvas.paste(strip, (margin, margin))
        try:
            found = recognize(canvas)
        except Exception:
            continue
        new_words = []
        for line in found or []:
            for word in getattr(line, "words", None) or []:
                if not word.text.strip():
                    continue
                r = word.bounding_rect
                if not y0 - 0.3 * bh <= r.y - margin + top + r.height / 2 <= y1 + 0.3 * bh:
                    continue
                new_words.append(SimpleNamespace(text=word.text, bounding_rect=SimpleNamespace(
                    x=r.x - margin + left_x, y=r.y - margin + top, width=r.width, height=r.height)))
        if not new_words:
            continue
        new_text = " ".join(w.text for w in new_words)
        if mine:
            old_text = " ".join(word.text for _, word in sorted(mine, key=lambda lw: lw[1].bounding_rect.x))
            new_left = min(w.bounding_rect.x for w in new_words)
            new_right = max(w.bounding_rect.x + w.bounding_rect.width for w in new_words)
            better_reach = new_right > right + 0.3 * bh or new_left < left - 0.3 * bh
            similar = difflib.SequenceMatcher(None, old_text, new_text).ratio() >= 0.6
            if not (better_reach and similar and len(new_text) >= len(old_text)):
                continue
            for line, word in mine:                            # replace the short reading
                line.words = [w for w in line.words if w is not word]
            words = [(line, word) for line, word in words if all(word is not w for _, w in mine)]
        line = SimpleNamespace(text=new_text, words=new_words)
        lines.append(line)
        words += [(line, w) for w in new_words]
    lines[:] = [line for line in lines if getattr(line, "words", None) or not hasattr(line, "words")]
    lines.sort(key=lambda line: min((w.bounding_rect.y for w in line.words), default=0))
    for line in lines:
        line.text = " ".join(w.text for w in line.words)
    return lines


END_MARKS = set('.,;:)]}"\'!?*')
START_MARKS = set('([{"\'')
ALL_MARKS = END_MARKS | START_MARKS | set("-")


def _marks_from_glyphs(gray, groups, top, base):
    """Read a run of small marks (left to right), e.g. ...") -> '...")'.
    None if any glyph isn't a recognisable mark."""
    out, i = [], 0
    while i < len(groups):
        if i + 1 < len(groups) and _classify_symbol(gray, groups[i:i + 2], top, base) == '"':
            out.append('"'); i += 2; continue
        mark = _classify_symbol(gray, [groups[i]], top, base)
        if mark not in ('"', "'", "*"):
            mark = _classify_punct(gray, groups[i], top, base)
        if mark is None:
            return None
        out.append(mark); i += 1
    return "".join(out)


def _fix_word_ends(b, boxes, gray, top, base, line_glyphs=None):
    """Compare the marks at the end (and start) of a word in the image with
    its text, and add the ones the OCR dropped: 'Browse.' + image 'Browse...")'
    -> 'Browse...")'. Only when every mark is recognised."""
    H = max(4.0, base - top)
    gap_limit = 0.6 * H
    others = [o for o in boxes if o is not b]
    next_x = min([o.x for o in others if o.x >= b.x + b.w - 1] + [b.x + b.w + 8 * H])
    prev_x = max([o.x + o.w for o in others if o.x + o.w <= b.x + 1] + [b.x - 8 * H])
    region = (max(prev_x + 1, b.x - 8 * H), top - 0.45 * H, min(next_x - 1, b.x + b.w + 8 * H), base + 0.6 * H)
    if line_glyphs is not None:      # the line's glyphs, found once for all its words
        groups = [g for g in line_glyphs if g.x0 >= region[0] and g.x1 <= region[2]]
    else:
        # (glyphs cut by the region's sides belong to the neighbouring words)
        groups = [g for g in _line_glyphs(gray, region) if not (g.touches & {"left", "right"})]
    # glyphs that another word's box claims are that word's (overlapping boxes)
    groups = [g for g in groups
              if not any(o.x <= (g.x0 + g.x1) / 2 <= o.x + o.w for o in others)]
    inside = [g for g in groups if g.x1 > b.x + 1 and g.x0 < b.x + b.w - 1]
    if not inside:
        return
    text = b.text
    end_chars, start_chars = "".join(ALL_MARKS), "".join(START_MARKS)
    new_x0, new_x1 = b.x, b.x + b.w

    # marks just past the box (a quote or bracket that sits closer to the
    # next word belongs to that word: are "hex")
    has_next = any(o.x >= b.x + b.w - 1 for o in others)
    has_prev = any(o.x + o.w <= b.x + 1 for o in others)
    tail, edge = [], b.x + b.w
    for g in groups:
        if g.x0 >= b.x + b.w - 1 and g not in inside:
            if g.x0 - edge > gap_limit or (has_next and next_x - g.x1 < g.x0 - edge):
                break
            tail.append(g); edge = g.x1
    head, edge = [], b.x
    for g in reversed(groups):
        if g.x1 <= b.x + 1 and g not in inside:
            if edge - g.x1 > gap_limit or (has_prev and g.x0 - prev_x < edge - g.x1):
                break
            head.insert(0, g); edge = g.x0

    # --- end of the word: marks at the end of the box plus those past it ---
    seq = inside + tail
    best = None
    for k in range(len(seq) - 1, 0, -1):            # keep at least one glyph that isn't a mark
        found = _marks_from_glyphs(gray, seq[k:], top, base)
        if found is None:
            break
        best = (k, found)
    have = text[len(text.rstrip(end_chars)):]
    if best and len(best[1]) > len(have) and best[1].startswith(have) and len(text) > len(have):
        text = text[:len(text) - len(have)] + best[1]           # 'started.' + image '.")'
        new_x1 = max(new_x1, seq[-1].x1)

    # --- start of the word ---
    seq = head + inside
    best = None
    for k in range(1, len(seq)):
        found = _marks_from_glyphs(gray, seq[:k], top, base)
        if found is None or not all(m in START_MARKS for m in found):
            break
        best = (k, found)
    have = text[:len(text) - len(text.lstrip(start_chars))]
    if best and len(best[1]) > len(have) and best[1].endswith(have) and len(text) > len(have):
        text = best[1] + text[len(have):]                       # '"Screen' + image '("Screen'
        new_x0 = min(new_x0, seq[0].x0)

    if text != b.text:
        b.word.text = b.text = text
    if (new_x0, new_x1) != (b.x, b.x + b.w):                  # the box now covers the marks
        r = b.word.bounding_rect
        b.word.bounding_rect = SimpleNamespace(x=new_x0, y=r.y, width=new_x1 - new_x0, height=r.height)
        b.x, b.w = new_x0, new_x1 - new_x0


def _repair_symbols(lines, gray):
    """Check the symbols OCR engines most often get wrong (Windows OCR in
    particular) against the image itself, and fix them in place:
    - bullets (• ◦ ▪), asterisks (*) and quote marks (" ') the engine left
      out are added back, at their place on the line,
    - a symbol read as a letter (* as x, " as '' , • as e or o) is corrected,
    - a 0 read as the letter O (or O as 0) is corrected from its context.
    lines: the OCR lines (words with text and bounding_rect, in the
    coordinates of gray, the prepared image)."""
    gray = gray.convert("L")
    all_words = []
    for line in lines:
        for word in getattr(line, "words", None) or []:
            r = word.bounding_rect
            all_words.append((r.x, r.y, r.x + r.width, r.y + r.height))

    def covered(g, pad=1):
        cx, cy = (g.x0 + g.x1) / 2, (g.y0 + g.y1) / 2
        return any(x0 - pad <= cx <= x1 + pad and y0 - pad <= cy <= y1 + pad for x0, y0, x1, y1 in all_words)

    for line in lines:
        words = [w for w in (getattr(line, "words", None) or []) if w.text.strip()]
        if not words:
            continue
        boxes = [SimpleNamespace(text=w.text, x=w.bounding_rect.x, y=w.bounding_rect.y,
                                 w=w.bounding_rect.width, h=w.bounding_rect.height, word=w) for w in words]
        top, base = _line_metrics(boxes, gray)
        H = max(4.0, base - top)

        # the line's ink, found once; glyphs(x0, x1) = the glyphs between x0 and x1
        # (a glyph that crosses x0 or x1 counts as touching that side, as if cut)
        line_left = max(0, min(bb.x for bb in boxes) - 8 * H)
        line_right = max(bb.x + bb.w for bb in boxes) + 8 * H
        line_blobs = [bl for bl in _components(gray, (line_left, top - 0.45 * H, line_right, base + 0.6 * H))
                      if not (bl.touches & {"top", "bottom"})]

        def glyphs(x0, x1):
            picked = []
            for bl in line_blobs:
                if bl.x1 > x0 and bl.x0 < x1:
                    sides = set()
                    if bl.x0 < x0: sides.add("left")
                    if bl.x1 > x1: sides.add("right")
                    picked.append(SimpleNamespace(x0=bl.x0, y0=bl.y0, x1=bl.x1, y1=bl.y1, n=bl.n, touches=sides))
            return _glyph_groups(picked)

        # 1) a symbol read as something else: one- or two-character words
        first_box = min(boxes, key=lambda b: b.x)
        for b in boxes:
            at_start = b is first_box and len(boxes) > 1
            core = b.text.strip(".,;:!?()[]\"'")
            if core in ("O", "o", "Ø", "ø", "©", "Θ", "θ") and core != b.text:
                # a lone 0 or O with punctuation ("Zone 0,"): judge the letter by its shape
                groups = glyphs(b.x - 1, b.x + b.w + 1)
                tall = [g for g in groups if g.y1 - g.y0 >= 0.85 * H and not (g.touches & {"top", "bottom"})]
                if len(tall) == 1 and tall[0].x1 - tall[0].x0 < 0.82 * (tall[0].y1 - tall[0].y0):
                    b.word.text = b.text = b.text.replace(core, "0", 1)
                continue
            if len(b.text) > 2 or (b.text.isalnum() and b.text not in SYMBOL_LOOKALIKES):
                continue
            groups = glyphs(b.x - 1, b.x + b.w + 1)
            groups = [g for g in groups if g.x1 > b.x and g.x0 < b.x + b.w and not (g.touches & {"top", "bottom"})]
            if not 1 <= len(groups) <= 2:
                continue
            symbol = _classify_symbol(gray, groups, top, base)
            if symbol in BULLETS and not at_start:
                symbol = None                                  # a round letter mid-line, not a bullet
            if symbol and symbol != b.text:
                b.word.text = b.text = symbol
            elif symbol is None and b.text in ("O", "o", "Ø", "ø", "©", "Θ", "θ") and len(groups) == 1:
                g = groups[0]
                w, h = g.x1 - g.x0, g.y1 - g.y0
                if h >= 0.85 * H and w < 0.82 * h:             # full height and narrow: the digit 0
                    b.word.text = b.text = "0"

        # 2) symbols at the start or end of a word: dropped, or read as a letter
        for b in boxes:
            if len(b.text) < 2 or not b.text.strip("".join(SYMBOL_LOOKALIKES)):
                continue
            edge = min(b.w, 1.3 * H)
            for side in ("left", "right"):
                x0 = b.x - 1 if side == "left" else b.x + b.w - edge
                x1 = b.x + edge if side == "left" else b.x + b.w + 1
                groups = glyphs(x0, x1)
                groups = [g for g in groups if g.x1 - g.x0 >= 1]
                if not groups:
                    continue
                inner = "right" if side == "left" else "left"
                ordered = groups if side == "left" else groups[::-1]
                trailing = ""
                if side == "right" and b.text[-1] in ":;,.!?)" and len(ordered) > 1:
                    trailing, ordered = b.text[-1], ordered[1:]   # look before the final : , . etc.
                first = ordered[0]
                if inner in first.touches:
                    continue                                   # a wide glyph: a letter
                candidates = [[first]]
                if len(ordered) > 1 and inner not in ordered[1].touches:
                    pair = sorted([first, ordered[1]], key=lambda g: g.x0)
                    candidates.insert(0, pair)
                symbol = None
                for c in candidates:
                    symbol = _classify_symbol(gray, c, top, base)
                    if symbol in ('"', "'", "*"):
                        break
                    symbol = None
                if symbol is None:
                    continue
                text = b.text[:-1] if trailing else b.text
                if symbol == "*":                              # ** or *** (e.g. **kwargs, **bold**)
                    k = 1
                    while k < 3 and len(ordered) > k and inner not in ordered[k].touches \
                            and _classify_symbol(gray, [ordered[k]], top, base) == "*":
                        k += 1
                    have = len(text) - len(text.lstrip("*")) if side == "left" else len(text) - len(text.rstrip("*"))
                    if 0 < have < k:
                        text = ("*" * k + text[have:]) if side == "left" else (text[:-have] + "*" * k)
                        b.word.text = b.text = text + trailing
                        continue
                    if have == 0 and k > 1:
                        # replace up to k look-alike marks at that edge with k stars
                        chars = text if side == "left" else text[::-1]
                        i = 0
                        while i < k and i < len(chars) - 1 and (chars[i] in STAR_LOOKALIKES or chars[i] == "*"):
                            i += 1
                        chars = "*" * k + chars[i:]
                        text = chars if side == "left" else chars[::-1]
                        b.word.text = b.text = text + trailing
                        continue
                if symbol == "'" and ((side == "left" and text[:2] in ("''", "``")) or
                                      (side == "right" and text[-2:] in ("''", "``"))):
                    symbol = '"'                               # '' is how OCR often reads "
                quotes = "\"'`‘’“”„″′"
                edge_char = text[0] if side == "left" else text[-1:]
                if symbol in ('"', "'") and edge_char and edge_char in quotes and edge_char not in ("''", "``"):
                    if not (side == "left" and text[:2] in ("''", "``")) and not (side == "right" and text[-2:] in ("''", "``")):
                        continue                               # a quote is already there
                if side == "left":
                    if text.startswith(symbol):
                        continue
                    lookalike = next((q for q in sorted(QUOTE_LOOKALIKES | STAR_LOOKALIKES, key=len, reverse=True)
                                      if text.startswith(q)), None)
                    if lookalike and len(text) > len(lookalike) and (
                            (symbol[0] == "*" and lookalike in STAR_LOOKALIKES) or
                            (symbol[0] != "*" and lookalike in QUOTE_LOOKALIKES)):
                        text = symbol + text[len(lookalike):]
                    elif first.x1 <= b.x + 0.15 * H or len(text) >= 2:
                        text = symbol + text
                else:
                    if text.endswith(symbol):
                        continue
                    lookalike = next((q for q in sorted(QUOTE_LOOKALIKES | STAR_LOOKALIKES, key=len, reverse=True)
                                      if text.endswith(q)), None)
                    if lookalike and len(text) > len(lookalike) and (
                            (symbol[0] == "*" and lookalike in STAR_LOOKALIKES) or
                            (symbol[0] != "*" and lookalike in QUOTE_LOOKALIKES)):
                        text = text[:-len(lookalike)] + symbol
                    else:
                        text = text + symbol
                b.word.text = b.text = text + trailing

        # 2b) punctuation at the end or start of a word: ..."), ("  - OCR
        #     engines often drop some of a run of small marks
        line_glyphs = _glyph_groups(line_blobs)
        for b in sorted(boxes, key=lambda bb: bb.x):
            _fix_word_ends(b, boxes, gray, top, base, line_glyphs)

        # 3) symbols left out entirely: blobs on the line not inside any word
        left = min(b.x for b in boxes)
        right = max(b.x + b.w for b in boxes)
        region = (max(0, left - 4 * H), top - 0.45 * H, right + 4 * H, base + 0.6 * H)
        all_words[:] = [(w.bounding_rect.x, w.bounding_rect.y, w.bounding_rect.x + w.bounding_rect.width,
                         w.bounding_rect.y + w.bounding_rect.height)
                        for ln in lines for w in (getattr(ln, "words", None) or [])]
        groups = [g for g in glyphs(region[0], region[2]) if not covered(g)]
        # pair up quote ticks
        found, i = [], 0
        while i < len(groups):
            g = groups[i]
            if i + 1 < len(groups):
                symbol = _classify_symbol(gray, [g, groups[i + 1]], top, base)
                if symbol == '"':
                    found.append((symbol, g.x0, groups[i + 1].x1))
                    i += 2
                    continue
            symbol = _classify_symbol(gray, [g], top, base) or _classify_punct(gray, g, top, base)
            if symbol:
                found.append((symbol, g.x0, g.x1))
            i += 1
        for symbol, x0, x1 in found:
            # touching a word: part of it (*Required, "quoted"); otherwise its own word
            before = [b for b in boxes if b.x + b.w <= x0 + 1]
            after = [b for b in boxes if b.x >= x1 - 1]
            near_after = min(after, key=lambda b: b.x) if after else None
            near_before = max(before, key=lambda b: b.x + b.w) if before else None
            punct = symbol in PUNCTUATION
            limit = (0.6 if symbol in ('"', "'") or punct else 0.25) * H
            gap_after = near_after.x - x1 if near_after is not None else 1e9
            gap_before = x0 - (near_before.x + near_before.w) if near_before is not None else 1e9
            if punct:
                # next to a word: already handled by _fix_word_ends
                if min(gap_after, gap_before) >= limit and symbol in "()[]:;":
                    # on its own (a lone dot or comma could be a speck)
                    new = SimpleNamespace(text=symbol, bounding_rect=SimpleNamespace(x=x0, y=top, width=x1 - x0, height=H))
                    pass
                    line.words.append(new)
                    boxes.append(SimpleNamespace(text=symbol, x=x0, y=top, w=x1 - x0, h=H, word=new))
                continue
            if symbol == "*" and near_after is not None and near_after.text.startswith("*") and gap_after < 0.45 * H:
                gap_after = 0                                  # ** written as two marks
            if symbol == "*" and near_before is not None and near_before.text.endswith("*") and gap_before < 0.45 * H:
                gap_before = 0
            if symbol == '"' and gap_after < limit and gap_before < limit:
                # between two words: an opening quote if the quotes so far are balanced
                opening = sum(bb.text.count('"') for bb in boxes if bb.x + bb.w <= x0) % 2 == 0
                gap_after, gap_before = (0, 1) if opening else (1, 0)
            if symbol not in BULLETS and gap_after <= gap_before and gap_after < limit \
                    and (symbol == "*" or not near_after.text.startswith(symbol)):
                near_after.word.text = near_after.text = symbol + near_after.text
                r = near_after.word.bounding_rect
                near_after.word.bounding_rect = SimpleNamespace(x=x0, y=r.y, width=r.x + r.width - x0, height=r.height)
                near_after.x, near_after.w = x0, r.x + r.width - x0
            elif symbol not in BULLETS and gap_before < limit and (symbol == "*" or not near_before.text.endswith(symbol)):
                near_before.word.text = near_before.text = near_before.text + symbol
                r = near_before.word.bounding_rect
                near_before.word.bounding_rect = SimpleNamespace(x=r.x, y=r.y, width=x1 - r.x, height=r.height)
                near_before.w = x1 - r.x
            else:
                new = SimpleNamespace(text=symbol, bounding_rect=SimpleNamespace(
                    x=x0, y=top + 0.3 * H if symbol in "•◦▪" else top, width=x1 - x0,
                    height=0.4 * H if symbol in "•◦▪" else 0.4 * H))
                line.words.append(new)
                boxes.append(SimpleNamespace(text=symbol, x=x0, y=new.bounding_rect.y, w=x1 - x0,
                                             h=new.bounding_rect.height, word=new))
        line.words.sort(key=lambda w: w.bounding_rect.x)
        line.text = " ".join(w.text for w in line.words)
    for line in lines:
        for word in getattr(line, "words", None) or []:
            word.text = _fix_zero_and_o(word.text)
        line.text = _fix_zero_and_o(line.text) if not getattr(line, "words", None) else " ".join(w.text for w in line.words)
    return lines


_NUMERIC_CHARS = set("0123456789OoØøΘθ.,:/-+$€£¥₹%#()[]")
_UNITS = ("°C", "°F", "°", "px", "ms", "s", "kb", "mb", "gb", "tb", "k", "m", "h", "min", "pt", "em", "rem", "x", "st", "nd", "rd", "th", "am", "pm")


def _fix_zero_and_o(text):
    """A 0 read as the letter O (or o, Ø, θ) inside a number becomes 0 again
    (1O.5O → 10.50, $2,OOO → $2,000, Ox00FF → 0x00FF); a 0 inside a word
    of letters becomes O (C0NTACT → CONTACT)."""
    def fix(token):
        core = token.strip("\"'“”‘’()[]{}<>,;:!?*")
        if not core:
            return token
        lower = core.lower()
        unit = ""
        for u in _UNITS:
            if lower.endswith(u.lower()) and len(core) > len(u):
                unit = core[-len(u):]
                break
        body = core[:-len(unit)] if unit else core
        prefix = ""
        if len(body) > 1 and body[0] in "vV#" and body[1] in "0123456789Oo":
            prefix, body = body[0], body[1:]
        if len(body) >= 3 and body[0] in "Oo" and body[1] == "-":
            return token[:token.index(core)] + core[0] + fix(core[1:]) + token[token.index(core) + len(core):]   # an ID such as O-1001
        if len(body) > 2 and body[0] in "0Oo" and body[1] in "xX" and all(
                c in "0123456789abcdefABCDEFOo" for c in body[2:]):
            fixed = "0x" + body[2:].replace("O", "0").replace("o", "0")
            return token.replace(core, prefix + fixed + unit)
        digits = sum(c.isdigit() for c in body)
        if body and all(c in _NUMERIC_CHARS for c in body) and any(c in "OoØøΘθ" for c in body):
            letters = [c for c in body if c in "OoØøΘθ"]
            if digits >= 1 or (len(letters) >= 2 and any(c in "$€£¥₹.,:" for c in body)) or (
                    body.endswith("%") and set(body[:-1]) <= set("Oo")):
                fixed = "".join("0" if c in "OoØøΘθ" else c for c in body)
                return token.replace(core, prefix + fixed + unit)
        # 1 or | read for the letter l between lowercase letters: Optiona11y -> Optionally
        if re.search(r"[a-z][1|]+[a-z]", core) and sum(c.isdigit() for c in core) == core.count("1") \
                and sum(c.isalpha() for c in core) >= 3:
            return token.replace(core, re.sub(r"(?<=[a-z])[1|]+(?=[a-z])", lambda m: "l" * len(m.group()), core))
        if "0" in core and core.replace("0", "").isalpha() and len(core) >= 3 and digits == core.count("0"):
            out = []
            for i, c in enumerate(core):
                if c == "0":
                    neighbours = core[max(0, i - 1):i] + core[i + 1:i + 2]
                    out.append("O" if i == 0 or neighbours.replace("0", "").isupper()
                               or not neighbours.replace("0", "") else "o")
                else:
                    out.append(c)
            return token.replace(core, "".join(out))
        return token
    return " ".join(fix(t) for t in text.split(" "))


def run_windows_ocr(image, language_tag=None):
    """Runs on the OCR thread. Returns dict(text, language, languages)."""
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.globalization import Language
    from winrt.windows.storage.streams import DataWriter
    from winrt.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat
    import winrt.windows.foundation            # noqa: F401  (async support)
    import winrt.windows.foundation.collections  # noqa: F401  (result lists)

    languages = [(lang.display_name, lang.language_tag)
                 for lang in OcrEngine.available_recognizer_languages]

    engine = None
    if language_tag:
        engine = OcrEngine.try_create_from_language(Language(language_tag))
    if engine is None:
        engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise RuntimeError(
            "No OCR language is installed in Windows. Add one in Settings > "
            "Time & language > Language & region (e.g. English), then try again.")

    prepared = _prepare_for_ocr(image, OcrEngine.max_image_dimension)

    def recognize(picture):
        """Windows OCR on a PIL image -> plain copies of the lines and word
        boxes (the WinRT objects belong to this thread), so the window can
        re-format them without running OCR again."""
        picture = picture.convert("RGBA")
        writer = DataWriter()
        writer.write_bytes(picture.tobytes())
        bitmap = SoftwareBitmap.create_copy_from_buffer(
            writer.detach_buffer(), BitmapPixelFormat.RGBA8, picture.width, picture.height)

        async def run():
            return await engine.recognize_async(bitmap)

        result = asyncio.run(run())
        found = []
        for line in result.lines:
            words = []
            for word in line.words:
                r = word.bounding_rect
                words.append(SimpleNamespace(text=word.text, bounding_rect=SimpleNamespace(
                    x=r.x, y=r.y, width=r.width, height=r.height)))
            found.append(SimpleNamespace(text=line.text, words=words))
        return found

    lines = recognize(prepared)
    try:
        _reread_lines(lines, prepared, recognize)   # lines missed or cut short: read again on their own
    except Exception as e:
        print(f"Re-reading lines skipped: {e}")
    try:
        _repair_symbols(lines, prepared)       # bullets, *, ", ) and 0 checked against the image
    except Exception as e:
        print(f"Symbol check skipped: {e}")

    return {
        "text": _layout_text(lines),
        "lines": lines,
        "language": engine.recognizer_language.display_name,
        "language_tag": engine.recognizer_language.language_tag,
        "languages": languages,
    }


def copy_to_clipboard(text):
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()


def extract_text_region():
    """Drag a rectangle on screen, then read the text inside it."""
    if recorder is not None:
        return
    if not OCR_AVAILABLE:
        messagebox.showinfo(
            "Text extraction unavailable",
            ("Text extraction uses the text recognition built into macOS (Vision).\n\n"
             "Install its Python bindings with:\n"
             f'  "{sys.executable}" -m pip install pyobjc-framework-Vision') if IS_MAC else
            ("Text extraction uses the OCR engine built into Windows 10/11.\n\n"
             "Install its Python bindings with:\n"
             f'  "{sys.executable}" -m pip install {" ".join(OCR_PACKAGES)}'),
            parent=root)
        return
    start_region_selection(ocr_selected_region)


def ocr_selected_region(x1, y1, x2, y2):
    try:
        image = grab_screen((x1, y1, x2, y2))
    except Exception as e:
        root.deiconify()
        status_var.set(f"Could not capture the area: {e}")
        return
    root.deiconify()
    root.lift()
    OcrWindow(image, source="Screen selection")


class OcrWindow:
    """Shows the extracted text: editable, with Copy / Save / language."""

    open_window = None

    def __init__(self, image, source=""):
        # Reuse one window instead of stacking many
        if OcrWindow.open_window is not None:
            try:
                OcrWindow.open_window.win.destroy()
            except Exception:
                pass
        OcrWindow.open_window = self

        self.image = image
        self.source = source
        self.language_tag = None
        self.language_map = {}
        self.first_result = True
        self.future = None
        self.lines = None                       # the OCR result, kept for re-formatting

        self.win = tk.Toplevel(root)
        self.win.title("Extracted Text")
        self.win.configure(bg=BG)
        self.win.transient(root)
        self.win.protocol("WM_DELETE_WINDOW", self.close)

        # --- header ---------------------------------------------------
        head = tk.Frame(self.win, bg=HEADER_BG)
        head.pack(fill="x")
        inner = tk.Frame(head, bg=HEADER_BG)
        inner.pack(fill="x", padx=S(20), pady=S(12))
        tk.Label(inner, image=icon("ocr", "white", 22), bg=HEADER_BG).pack(side="left")
        titles = tk.Frame(inner, bg=HEADER_BG)
        titles.pack(side="left", padx=(S(10), 0))
        tk.Label(titles, text="Extracted Text", font=F_TILE, fg="white",
                 bg=HEADER_BG).pack(anchor="w")
        self.info_label = tk.Label(titles, text="", font=F_TILE_SUB,
                                   fg=HEADER_SUB, bg=HEADER_BG)
        self.info_label.pack(anchor="w")
        tk.Frame(self.win, bg=ACCENT, height=S(3)).pack(fill="x")

        body = tk.Frame(self.win, bg=BG)
        body.pack(fill="both", expand=True, padx=S(18), pady=S(14))

        # --- language row --------------------------------------------
        lang_row = tk.Frame(body, bg=BG)
        lang_row.pack(fill="x", pady=(0, S(8)))
        tk.Label(lang_row, text="Language", font=F_BODY_BOLD, fg=TEXT,
                 bg=BG).pack(side="left", padx=(0, S(8)))
        self.language_var = tk.StringVar()
        self.language_box = ttk.Combobox(lang_row, textvariable=self.language_var,
                                         state="readonly", width=30, font=F_BODY)
        self.language_box.pack(side="left")
        self.language_box.bind("<<ComboboxSelected>>", self._language_changed)

        # Keep layout: indentation, columns and spacing as on screen (code,
        # tables, forms), shown in a fixed-width font so columns line up.
        # Off: plain flowing text.
        self.layout_var = tk.BooleanVar(value=True)
        Toggle(lang_row, "Keep layout", self.layout_var, bg=BG).pack(side="left", padx=(S(16), 0))
        self.layout_var.trace_add("write", lambda *a: self._reformat())

        open_button = FlatButton(lang_row, "Open image...", self.open_image)
        open_button.pack(side="right")

        # --- text area -----------------------------------------------
        text_frame = tk.Frame(body, bg=CARD, highlightthickness=S(2) if S(2) > 1 else 1,
                              highlightbackground=FIELD_BORDER)
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(text_frame, wrap="none", font=font(10.5, "Consolas"),
                            fg=TEXT, bg=CARD, relief="flat", bd=0,
                            padx=S(12), pady=S(10), undo=True,
                            width=72, height=18, insertbackground=TEXT,
                            selectbackground="#C7CBF9", selectforeground=TEXT)
        scroll = tk.Scrollbar(text_frame, command=self.text.yview)
        self.xscroll = tk.Scrollbar(text_frame, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=scroll.set, xscrollcommand=self.xscroll.set)
        scroll.pack(side="right", fill="y")
        self.xscroll.pack(side="bottom", fill="x")
        self.text.pack(side="left", fill="both", expand=True)

        # --- buttons -------------------------------------------------
        buttons = tk.Frame(body, bg=BG)
        buttons.pack(fill="x", pady=(S(12), 0))
        self.copy_button = FlatButton(buttons, "Copy text", self.copy, style="primary")
        self.copy_button.pack(side="left")
        FlatButton(buttons, "Save as .txt", self.save).pack(side="left", padx=(S(8), 0))
        FlatButton(buttons, "Close", self.close).pack(side="right")

        self.win.update_idletasks()
        self._center()
        self.run()

    def _center(self):
        left, top, width, height = screen_work_area()
        w = self.win.winfo_reqwidth()
        h = self.win.winfo_reqheight() + S(39)
        self.win.geometry(f"+{left + max(0, (width - w) // 2)}+{top + max(0, (height - h) // 2)}")

    # ---- OCR --------------------------------------------------------

    def run(self):
        self._set_text("Reading text...")
        self.text.configure(fg=MUTED)
        self.info_label.configure(text=f"{self.source}  ·  "
                                       f"{self.image.width} x {self.image.height} px")
        self.copy_button.config(state="disabled")
        self.future = OCR_WORKER.submit(run_ocr, self.image, self.language_tag)
        self.win.after(100, self._poll)

    def _poll(self):
        if self.future is None or not self.win.winfo_exists():
            return
        if not self.future.done():
            self.win.after(100, self._poll)
            return

        try:
            result = self.future.result()
        except Exception as e:
            self.text.configure(fg=RED)
            self._set_text(f"Text extraction failed:\n{e}")
            status_var.set("Could not extract text (see the text window)")
            return

        self.lines = result.get("lines")
        if self.lines is not None:
            result["text"] = _layout_text(self.lines, self.layout_var.get())
        self.language_map = {name: tag for name, tag in result["languages"]}
        self.language_box.configure(values=list(self.language_map))
        self.language_var.set(result["language"])
        self.language_tag = result["language_tag"]

        text = result["text"]
        self.text.configure(fg=TEXT)
        self.copy_button.config(state="normal")

        if not text:
            self._set_text("")
            self.info_label.configure(text=f"{self.source}  ·  no text found")
            status_var.set("No text found in the selected area")
            return

        self._set_text(text)
        lines = text.count("\n") + 1
        words = len(text.split())
        self.info_label.configure(
            text=f"{self.source}  ·  {words} words, {lines} lines  ·  "
                 f"{result['language']}")

        # Copy automatically the first time, so it's ready to paste
        if self.first_result:
            self.first_result = False
            copy_to_clipboard(text)
            status_var.set(f"Text extracted and copied to clipboard ({words} words)")

    def _reformat(self):
        """Keep layout switched: fixed-width font and no wrapping for the
        layout, flowing text otherwise; re-format the last result."""
        if self.layout_var.get():
            self.text.configure(font=font(10.5, "Consolas"), wrap="none")
        else:
            self.text.configure(font=font(11), wrap="word")
        if self.lines is not None:
            self._set_text(_layout_text(self.lines, self.layout_var.get()))

    def _set_text(self, value):
        self.text.delete("1.0", "end")
        self.text.insert("1.0", value)
        self.text.edit_reset()

    def _language_changed(self, event=None):
        tag = self.language_map.get(self.language_var.get())
        if tag and tag != self.language_tag:
            self.language_tag = tag
            self.run()

    # ---- buttons ----------------------------------------------------

    def current_text(self):
        return self.text.get("1.0", "end-1c")

    def copy(self):
        text = self.current_text()
        copy_to_clipboard(text)
        status_var.set(f"Text copied to clipboard ({len(text.split())} words)")
        self.copy_button.config(text="Copied!")
        self.win.after(1500, lambda: self.copy_button.winfo_exists()
                       and self.copy_button.config(text="Copy text"))

    def save(self):
        folder = save_folder or (word_doc_path.parent if word_doc_path else Path.home())
        path = filedialog.asksaveasfilename(
            parent=self.win,
            title="Save extracted text",
            initialdir=str(folder),
            initialfile=f"Text_{timestamp()}.txt",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")])
        if not path:
            return
        try:
            Path(path).write_text(self.current_text(), encoding="utf-8")
            status_var.set(f"Saved: {Path(path).name}")
        except Exception as e:
            messagebox.showerror("Could not save", str(e), parent=self.win)

    def open_image(self):
        folder = save_folder or Path.home()
        path = filedialog.askopenfilename(
            parent=self.win,
            title="Choose an image to read text from",
            initialdir=str(folder),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"),
                       ("All files", "*.*")])
        if not path:
            return
        try:
            with Image.open(path) as img:
                self.image = img.convert("RGB")
        except Exception as e:
            messagebox.showerror("Could not open image", str(e), parent=self.win)
            return
        self.source = Path(path).name
        self.first_result = True
        self.run()

    def close(self):
        OcrWindow.open_window = None
        self.win.destroy()


# ============================================================
# Blur Sensitive Information
# ============================================================
# Blurring is a tool in the Review window (shortcut B): drag a box over
# anything sensitive and that part of the screenshot gets a Gaussian blur.

BLUR_DEFAULT_RADIUS = 12
BLUR_RADII = [4, 8, 12, 20, 32]       # Review window "Strength" choices (image pixels)


def blur_region(image, box, radius):
    """Blur one rectangular area of the image in place."""
    blurred = image.crop(box)
    blurred = blurred.filter(ImageFilter.GaussianBlur(radius=radius))
    image.paste(blurred, box)


def deliver_capture(image, kind, taken_at):
    """Every new screenshot opens in the Review window (annotate, blur,
    crop...) and is saved only when the user clicks Save."""
    annotate_then_save(image, kind, taken_at)


# ============================================================
# Review & Annotate window (opens after every capture)
# ============================================================
# After each screenshot (Capture Screen, Capture Region, Full Page) the
# captured image is shown in the Review window. There the user can
# highlight, draw, add shapes and text, blur sensitive areas and crop, then:
#   Copy   - put the image on the clipboard (paste into Teams, Outlook ...)
#   Save   - save it to the folder / Word document
#   Cancel - discard it

def annotate_then_save(image, kind, taken_at):
    """Show the captured image in the Review window; save only on Save."""
    root.after(50, lambda: AnnotateEditor(
        image,
        on_save=lambda result: handle_capture(result, kind, taken_at),
        title=f"{kind} capture",
        save_label=f"Save to {doc_app()}" if word_doc_path is not None else "Save",
    ))


# ---- copy an image to the Windows clipboard ---------------------------

def image_to_dib(image):
    """The image as a Windows device-independent bitmap (BMP without its
    14-byte file header) - the format the clipboard uses."""
    buffer = BytesIO()
    image.convert("RGB").save(buffer, format="BMP")
    return buffer.getvalue()[14:]


def copy_image_to_clipboard(image):
    """Put an image on the clipboard (as a bitmap and as PNG), so it can be
    pasted into Teams, Outlook, Word, Paint, etc. Returns True on success."""
    if IS_MAC:
        return mac_copy_image(image)
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
    user32.OpenClipboard.argtypes = [ctypes.c_void_p]
    user32.SetClipboardData.argtypes = [wintypes.UINT, ctypes.c_void_p]
    user32.SetClipboardData.restype = ctypes.c_void_p
    user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    user32.RegisterClipboardFormatW.restype = wintypes.UINT

    png = BytesIO()
    image.convert("RGB").save(png, format="PNG")
    formats = [(8, image_to_dib(image)),                                   # CF_DIB
               (user32.RegisterClipboardFormatW("PNG"), png.getvalue())]   # PNG (keeps quality)

    for _ in range(10):                      # the clipboard may be busy for a moment
        if user32.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        return False
    try:
        user32.EmptyClipboard()
        ok = False
        for fmt, data in formats:
            handle = kernel32.GlobalAlloc(0x0002, len(data))          # GMEM_MOVEABLE
            if not handle:
                continue
            pointer = kernel32.GlobalLock(handle)
            ctypes.memmove(pointer, data, len(data))
            kernel32.GlobalUnlock(handle)
            if user32.SetClipboardData(fmt, handle):
                ok = True                     # the clipboard owns the memory now
            else:
                kernel32.GlobalFree(handle)
        return ok
    finally:
        user32.CloseClipboard()


class AnnotateEditor:
    """Review window shown after each capture: annotate the screenshot with
    highlighter, pen, rectangles and straight lines, then Copy (clipboard),
    Save or Cancel.

    Tools (keyboard shortcut):
      Highlighter (H)  Pen (P)  Eraser (X)  Rectangle (R)  Line (L)
      Text (T)  Blur (B)  Crop (C)
    Blur: drag a box over sensitive information; it blurs the screenshot
    itself (annotations stay sharp on top). Strength has five levels.
    Text: click where the note should go and type; Enter finishes,
    Shift+Enter starts a new line, Esc removes the note being typed.
    Hold Shift while drawing for a perfect square, or a line snapped to
    horizontal / vertical / 45 degrees.
    Crop: drag the area to keep, then Apply crop (or Enter). Undo reverts it.
    The window can be minimized, maximized and resized; the screenshot
    is scaled to fit.
    Ctrl+Z undoes the last change, Enter saves, Esc cancels.
    """

    TOOLS = [
        ("highlight", "Highlighter", "H"),
        ("pen", "Pen", "P"),
        ("eraser", "Eraser", "X"),
        ("rect", "Rectangle", "R"),
        ("line", "Line", "L"),
        ("text", "Text", "T"),
        ("blur", "Blur", "B"),
        ("crop", "Crop", "C"),
    ]
    COLORS = ["#E5484D", "#F97316", "#FACC15", "#22C55E", "#3B82F6",
              "#8B5CF6", "#EC4899", "#111827", "#FFFFFF"]
    # thickness choices in screen pixels (what you see while drawing)
    SIZES = {"highlight": [10, 16, 24, 34, 46], "default": [2, 3, 5, 8, 12],
             "text": [14, 18, 24, 32, 44]}         # text: letter size in screen px

    def __init__(self, image, on_save, title="", save_label="Save"):
        self.original = image.convert("RGB")
        self.on_save = on_save
        self.ops = []                    # everything drawn, in order (for undo)
        self.drag_start = None
        self.preview = []                # canvas items while drawing
        self.stroke = []                 # points of the current pen/highlighter stroke
        self.photo = None
        self.tool = "highlight"
        self.settings = {                # colour / size remembered per tool
            "highlight": {"color": "#FACC15", "size": 1},
            "pen": {"color": "#E5484D", "size": 1},
            "eraser": {"color": None, "size": 2},
            "rect": {"color": "#E5484D", "size": 1},
            "line": {"color": "#E5484D", "size": 1},
            "text": {"color": "#E5484D", "size": 2},
            "blur": {"color": None, "size": BLUR_RADII.index(BLUR_DEFAULT_RADIUS)},
            "crop": {"color": None, "size": 0},
        }
        self.text_edit = None            # note being typed: (widget, canvas item, image pos)
        self.crop_sel = None             # pending crop selection (canvas coords)
        self.prev_tool = "highlight"     # tool to return to after cropping
        self._relayout_job = None
        self._last_canvas_size = None

        self.win = tk.Toplevel(root)
        self.win.title("Review screenshot")
        self.win.configure(bg=BG)
        # Not transient: a normal window gets Minimize / Maximize buttons
        # and its own taskbar button, and can be resized.
        self.win.resizable(True, True)
        self.win.protocol("WM_DELETE_WINDOW", self.cancel)

        # --- header ---------------------------------------------------
        head = tk.Frame(self.win, bg=HEADER_BG)
        head.pack(fill="x")
        inner = tk.Frame(head, bg=HEADER_BG)
        inner.pack(fill="x", padx=S(20), pady=S(10))
        tk.Label(inner, image=icon("pen", "white", 22), bg=HEADER_BG).pack(side="left")
        titles = tk.Frame(inner, bg=HEADER_BG)
        titles.pack(side="left", padx=(S(10), 0))
        tk.Label(titles, text="Review & annotate" + (f"  ·  {title}" if title else ""),
                 font=F_TILE, fg="white", bg=HEADER_BG).pack(anchor="w")
        tk.Label(titles,
                 text=f"Annotate if you like, then Copy, Save or Cancel.  {KEY}+C copy, {KEY}+Z undo",
                 font=F_TILE_SUB, fg=HEADER_SUB, bg=HEADER_BG).pack(anchor="w")
        self.info = tk.Label(inner, text="", font=F_TILE_SUB, fg=HEADER_SUB, bg=HEADER_BG)
        self.info.pack(side="right")
        tk.Frame(self.win, bg=ACCENT, height=S(3)).pack(fill="x")

        body = tk.Frame(self.win, bg=BG)
        body.pack(fill="both", expand=True, padx=S(16), pady=S(10))

        # --- toolbar row 1: tools + undo/clear ------------------------
        bar1 = tk.Frame(body, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        bar1.pack(fill="x")
        tools = tk.Frame(bar1, bg=CARD)
        tools.pack(side="left", padx=S(8), pady=S(6))
        self.tool_buttons = {}
        for key, label, shortcut in self.TOOLS:
            b = tk.Label(tools, image=icon(key, TEXT, 18),
                         text=f" {label}", compound="left", font=F_BODY_BOLD,
                         bg=CARD, fg=TEXT, padx=S(9), pady=S(5), cursor="hand2",
                         highlightthickness=1, highlightbackground=CARD)
            b.pack(side="left", padx=S(2))
            b.bind("<Button-1>", lambda e, k=key: self.select_tool(k))
            b.bind("<Enter>", lambda e, k=key: self._tool_hover(k, True))
            b.bind("<Leave>", lambda e, k=key: self._tool_hover(k, False))
            self.tool_buttons[key] = b
            self.win.bind(f"<KeyPress-{shortcut.lower()}>", lambda e, k=key: self.select_tool(k))

        self.clear_button = FlatButton(bar1, "Clear all", self.clear)
        self.clear_button.pack(side="right", padx=(S(4), S(8)))
        self.undo_button = FlatButton(bar1, "Undo", self.undo)
        self.undo_button.pack(side="right")

        # --- toolbar row 2: colour and thickness -----------------------
        bar2 = tk.Frame(body, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        bar2.pack(fill="x", pady=(S(6), 0))

        self.color_frame = tk.Frame(bar2, bg=CARD)
        self.color_frame.pack(side="left", padx=S(10), pady=S(7))
        tk.Label(self.color_frame, text="Colour", font=F_BODY_BOLD, fg=TEXT,
                 bg=CARD).pack(side="left", padx=(0, S(8)))
        self.swatches = {}
        for color in self.COLORS:
            sw = tk.Frame(self.color_frame, bg=color, width=S(20), height=S(20),
                          highlightthickness=2, highlightbackground=CARD,
                          cursor="hand2")
            sw.pack(side="left", padx=S(2))
            sw.bind("<Button-1>", lambda e, c=color: self.set_color(c))
            self.swatches[color] = sw
        self.custom_swatch = tk.Label(self.color_frame, text="+", font=F_BODY_BOLD,
                                      bg=CARD, fg=TEXT, width=2, cursor="hand2",
                                      highlightthickness=2, highlightbackground=FIELD_BORDER)
        self.custom_swatch.pack(side="left", padx=(S(4), 0))
        self.custom_swatch.bind("<Button-1>", lambda e: self.pick_custom_color())

        self.size_frame = tk.Frame(bar2, bg=CARD)
        self.size_frame.pack(side="left", padx=S(14), pady=S(7))
        self.size_label = tk.Label(self.size_frame, text="Thickness", font=F_BODY_BOLD,
                                   fg=TEXT, bg=CARD)
        self.size_label.pack(side="left", padx=(0, S(8)))
        self.size_buttons = []
        for i in range(5):
            b = tk.Label(self.size_frame, image=size_dot_image(i, False), bg=CARD,
                         cursor="hand2")
            b.pack(side="left", padx=S(1))
            b.bind("<Button-1>", lambda e, i=i: self.set_size(i))
            self.size_buttons.append(b)

        # text option: white (or dark) box behind the note so it stands out
        self.text_box_var = tk.BooleanVar(value=True)
        self.text_box_toggle = Toggle(bar2, "Background", self.text_box_var)
        self.text_box_var.trace_add("write", lambda *a: self._style_text_widget())

        # crop controls (shown instead of colour / thickness for Crop)
        self.crop_frame = tk.Frame(bar2, bg=CARD)
        self.crop_hint = tk.Label(self.crop_frame, text="Drag over the part you want to keep",
                                  font=F_BODY, fg=MUTED, bg=CARD)
        self.crop_hint.pack(side="left", padx=(0, S(12)))
        self.crop_apply = FlatButton(self.crop_frame, "Apply crop", self.apply_crop,
                                     style="primary")
        self.crop_apply.pack(side="left")
        self.crop_cancel = FlatButton(self.crop_frame, "Clear selection", self.clear_crop_selection)
        self.crop_cancel.pack(side="left", padx=(S(6), 0))

        # --- bottom buttons (packed first so they stay visible when the
        #     window is made smaller) ---------------------------------
        bottom = tk.Frame(body, bg=BG)
        bottom.pack(side="bottom", fill="x", pady=(S(10), 0))
        self.copied = False
        self.copy_button = FlatButton(bottom, "Copy", self.copy,
                                      icons={"Copy": icon("copy", TEXT, 15),
                                             "Copied": icon("check", GREEN, 15)})
        self.copy_button.pack(side="left")
        FlatButton(bottom, save_label, self.save, style="primary").pack(side="right")
        FlatButton(bottom, "Cancel", self.cancel).pack(side="right", padx=(0, S(6)))

        # --- image canvas -------------------------------------------
        left, top, width, height = screen_work_area()
        max_w, max_h = int(width * 0.82), int(height * 0.56)
        self.crop_box = (0, 0) + self.original.size

        frame = tk.Frame(body, bg=CARD, highlightthickness=1, highlightbackground=FIELD_BORDER)
        frame.pack(fill="both", expand=True, pady=(S(10), 0))
        self.canvas = tk.Canvas(frame, bg="#2A2F3A", highlightthickness=0, cursor="crosshair")
        self.scroll = tk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.image_item = self.canvas.create_image(0, 0, anchor="nw")
        self.canvas.bind("<MouseWheel>", self._wheel)
        self.canvas.bind("<Button-4>", lambda e: self.tall and self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.tall and self.canvas.yview_scroll(3, "units"))

        # initial size: fit the screenshot into most of the screen
        self._layout(max_w, max_h)
        canvas_w, canvas_h = self.view_w, min(self.view_h, max_h)
        self.canvas.configure(width=canvas_w, height=canvas_h)
        # lay out again for the canvas's actual size (centring, scroll area)
        self._layout(canvas_w + (self.scroll.winfo_reqwidth() if self.tall else 0), canvas_h)

        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._drag)
        self.canvas.bind("<ButtonRelease-1>", self._release)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        self.win.bind("<Control-c>", lambda e: self.copy())
        self.win.bind("<Control-C>", lambda e: self.copy())
        self.win.bind("<Control-z>", lambda e: self.undo())
        self.win.bind("<Control-Z>", lambda e: self.undo())
        mirror_control_keys(self.win)
        self.win.bind("<Return>", self._on_enter)
        self.win.bind("<Escape>", self._on_escape)

        self._reset_layers()
        self.working = self.original.copy()
        self.select_tool("highlight")
        self.redraw()

        self.win.update_idletasks()
        w = self.win.winfo_reqwidth()
        h = self.win.winfo_reqheight() + S(39)
        # can be made smaller, but not so small the toolbar is cut off
        self.win.minsize(min(w, bar1.winfo_reqwidth() + S(40)), min(h, S(420)))
        self.win.geometry(f"+{left + max(0, (width - w) // 2)}+{top + max(0, (height - h) // 2)}")
        self.win.lift()
        self.win.focus_force()
        self.win.grab_set()   # finish (save/cancel) before doing anything else

    # ---- toolbar ----------------------------------------------------

    def _tool_hover(self, key, on):
        if key != self.tool:
            self.tool_buttons[key].configure(bg="#EEF0F4" if on else CARD)

    def select_tool(self, key):
        self._commit_text()
        if key == "crop" and self.tool != "crop":
            self.prev_tool = self.tool
        if key != "crop":
            self.clear_crop_selection()
        self.tool = key
        for k, b in self.tool_buttons.items():
            selected = k == key
            b.configure(bg=ACCENT_SOFT if selected else CARD,
                        fg=ACCENT if selected else TEXT,
                        image=icon(k, ACCENT if selected else TEXT, 18),
                        highlightbackground=ACCENT if selected else CARD)
        # Crop shows its own controls instead of colour / thickness;
        # the eraser has no colour, so the colour row is hidden for it
        for f in (self.color_frame, self.size_frame, self.crop_frame, self.text_box_toggle):
            f.pack_forget()
        if key == "crop":
            self.crop_frame.pack(side="left", padx=S(10), pady=S(5))
            self._update_crop_buttons()
        else:
            if key not in ("eraser", "blur"):
                self.color_frame.pack(side="left", padx=S(10), pady=S(7))
            self.size_frame.pack(side="left", padx=S(14), pady=S(7))
            if key == "text":
                self.text_box_toggle.pack(side="left", padx=(S(4), S(10)), pady=S(7))
        self.size_label.configure(text={"text": "Size", "blur": "Strength"}.get(key, "Thickness"))
        self.canvas.configure(cursor={"eraser": "circle", "text": "xterm"}.get(key, "crosshair"))
        self._show_color_and_size()

    def _show_color_and_size(self):
        if self.tool == "crop":
            return
        current = self.settings[self.tool]
        for color, sw in self.swatches.items():
            sw.configure(highlightbackground=TEXT if color == current["color"] else
                         (FIELD_BORDER if color == "#FFFFFF" else CARD))
        custom = current["color"] not in self.swatches
        self.custom_swatch.configure(
            bg=current["color"] if custom else CARD,
            highlightbackground=TEXT if custom else FIELD_BORDER)
        for i, b in enumerate(self.size_buttons):
            b.configure(image=size_dot_image(i, i == current["size"]))

    def set_color(self, color):
        if self.tool in self.settings:
            self.settings[self.tool]["color"] = color
            self._show_color_and_size()
            self._style_text_widget()

    def pick_custom_color(self):
        from tkinter import colorchooser
        current = self.settings.get(self.tool, {}).get("color", "#E5484D")
        result = colorchooser.askcolor(color=current, parent=self.win, title="Choose a colour")
        if result and result[1]:
            self.set_color(result[1].upper())

    def set_size(self, index):
        if self.tool in self.settings:
            self.settings[self.tool]["size"] = index
            self._show_color_and_size()
            self._style_text_widget()

    def _screen_width(self):
        """Selected thickness in screen pixels."""
        sizes = (self.SIZES["highlight"] if self.tool in ("highlight", "eraser")
                 else self.SIZES["text"] if self.tool == "text"
                 else self.SIZES["default"])
        return sizes[self.settings[self.tool]["size"]]

    # ---- drawing ----------------------------------------------------

    # Annotations are kept on transparent layers over the screenshot:
    #   highlight layer (multiplied onto the image, like a real marker),
    #   ink layer (pen), shape layer (rectangles / lines, always on top).
    # The eraser clears pixels on the highlight and ink layers only, so it
    # removes pen and highlighter marks but never shapes or the screenshot.

    def _reset_layers(self):
        size = self.original.size
        self.crop_box = (0, 0) + size
        self.base = self.original.copy()     # the screenshot itself (blur changes this)
        self.hl_layer = Image.new("RGBA", size, (0, 0, 0, 0))
        self.ink_layer = Image.new("RGBA", size, (0, 0, 0, 0))
        self.shape_layer = Image.new("RGBA", size, (0, 0, 0, 0))

    def rebuild(self):
        """Re-apply every change to a fresh copy of the original."""
        self._reset_layers()
        for op in self.ops:
            self._apply(op)
        self._compose()

    def _apply(self, op):
        kind = op["type"]
        if kind == "highlight":
            draw_path(self.hl_layer, op["points"], _hex_rgb(op["color"]) + (255,), op["width"])
        elif kind == "pen":
            draw_path(self.ink_layer, op["points"], _hex_rgb(op["color"]) + (255,), op["width"])
        elif kind == "eraser":
            for layer in (self.hl_layer, self.ink_layer):
                draw_path(layer, op["points"], (0, 0, 0, 0), op["width"])
        elif kind == "line":
            draw_line(self.shape_layer, op["points"], op["color"], op["width"])
        elif kind == "crop":
            self.crop_box = op["box"]     # annotations stay; only the view is cut
        elif kind == "blur":
            blur_region(self.base, op["box"], op["radius"])
        elif kind == "text":
            draw_text(self.shape_layer, op["pos"], op["text"], op["color"],
                      op["size"], op["box"])
        else:
            draw_shape(self.shape_layer, kind, op["box"], op["color"], op["width"])

    def _compose(self):
        """Screenshot (with any blur) + highlight (multiply) + ink + shapes
        -> self.working."""
        image = self.base.copy()
        box = self.hl_layer.getbbox()
        if box:
            region = image.crop(box)
            layer = self.hl_layer.crop(box)
            marked = ImageChops.multiply(region, layer.convert("RGB"))
            region.paste(marked, (0, 0), layer.getchannel("A"))
            image.paste(region, box)
        for layer in (self.ink_layer, self.shape_layer):
            box = layer.getbbox()
            if box:
                part = layer.crop(box)
                image.paste(part.convert("RGB"), box, part.getchannel("A"))
        if self.crop_box != (0, 0) + image.size:
            image = image.crop(self.crop_box)
        self.working = image

    def redraw(self):
        view = self.working
        if self.scale < 1.0:
            view = view.resize((self.view_w, self.view_h),
                               Image.BILINEAR if self.tall else Image.LANCZOS)
        self.view_img = view              # what is on screen (for the crop overlay)
        self._dim_img = None
        self.photo = ImageTk.PhotoImage(view)
        self.canvas.itemconfigure(self.image_item, image=self.photo)
        self.canvas.coords(self.image_item, self.off_x, self.off_y)

        n = len(self.ops)
        w, h = self.working.size
        self.info.configure(text=f"{w} \u00d7 {h} px   \u00b7   {n} change{'s' if n != 1 else ''}")
        state = "normal" if n else "disabled"
        self.undo_button.config(state=state)
        self.clear_button.config(state=state)

    # ---- mouse ------------------------------------------------------

    # ---- layout (fits the screenshot to the window) -----------------

    def _layout(self, avail_w, avail_h):
        """Choose the display scale for the current (cropped) image so it
        fits avail_w x avail_h. Tall images (full-page captures) are shown
        at a readable width with a scrollbar instead."""
        x0, y0, x1, y1 = self.crop_box
        cw, ch = max(1, x1 - x0), max(1, y1 - y0)
        sbw = self.scroll.winfo_reqwidth()
        width_scale = min(1.0, avail_w / cw)
        self.tall = ch * width_scale > avail_h * 1.3
        if self.tall:
            draw_w = max(1, avail_w - sbw)
            self.scale = min(1.0, draw_w / cw)
        else:
            draw_w = avail_w
            self.scale = min(1.0, avail_w / cw, avail_h / ch)
        self.view_w = max(1, int(cw * self.scale))
        self.view_h = max(1, int(ch * self.scale))
        self.off_x = max(0, (draw_w - self.view_w) // 2)
        self.off_y = 0 if self.tall else max(0, (avail_h - self.view_h) // 2)
        if self.tall:
            if not self.scroll.winfo_ismapped():
                self.scroll.pack(side="right", fill="y", before=self.canvas)
            self.canvas.configure(scrollregion=(0, 0, draw_w, self.view_h))
        else:
            self.scroll.pack_forget()
            self.canvas.configure(scrollregion=(0, 0, avail_w, avail_h))
            self.canvas.yview_moveto(0)

    def _on_canvas_resize(self, event):
        size = (event.width, event.height)
        if size == self._last_canvas_size:
            return
        self._last_canvas_size = size
        if self._relayout_job:
            self.win.after_cancel(self._relayout_job)
        self._relayout_job = self.win.after(60, self._relayout)

    def _relayout(self):
        """Re-fit after the window is resized / maximized or the image cropped."""
        self._relayout_job = None
        if not self.canvas.winfo_exists():
            return
        self._commit_text()
        self.canvas.update_idletasks()
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if w < 10 or h < 10:            # minimized
            return
        if self.scroll.winfo_ismapped():
            w += self.scroll.winfo_width()
        before = (self.scale, self.off_x, self.off_y, self.view_w, self.view_h,
                  self.crop_box, getattr(self, "_shown_box", None))
        self._layout(w, h)
        self._shown_box = self.crop_box
        if before == (self.scale, self.off_x, self.off_y, self.view_w, self.view_h,
                      self.crop_box, self.crop_box):
            return                       # nothing moved - keep any selection
        self.clear_crop_selection()
        self._clear_preview()
        self.redraw()

    def _wheel(self, event):
        if self.tall:
            self.canvas.yview_scroll(wheel_steps(event) * 3, "units")

    def _clamp(self, x, y):
        """Keep a canvas point on the image."""
        return (min(max(x, self.off_x), self.off_x + self.view_w),
                min(max(y, self.off_y), self.off_y + self.view_h))

    def _img(self, x, y):
        """Canvas point -> pixel position in the original screenshot."""
        return ((x - self.off_x) / self.scale + self.crop_box[0],
                (y - self.off_y) / self.scale + self.crop_box[1])

    def _img_box(self, x0, y0, x1, y1):
        """Canvas rectangle -> whole-pixel box in the original, kept inside
        the current crop."""
        ax, ay = self._img(min(x0, x1), min(y0, y1))
        bx, by = self._img(max(x0, x1), max(y0, y1))
        cx0, cy0, cx1, cy1 = self.crop_box
        return (max(cx0, int(ax)), max(cy0, int(ay)),
                min(cx1, int(round(bx))), min(cy1, int(round(by))))

    def _point(self, event):
        """Mouse position in view coordinates (accounts for scrolling)."""
        return self._clamp(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))

    def _shape_end(self, event):
        """End point of a shape; Shift makes a square, or snaps a line
        to horizontal / vertical / 45 degrees."""
        x, y = self._point(event)
        if event.state & 0x0001 and self.tool == "line":
            x0, y0 = self.drag_start
            dx, dy = x - x0, y - y0
            if abs(dx) > 2.414 * abs(dy):        # within 22.5 deg of horizontal
                y = y0
            elif abs(dy) > 2.414 * abs(dx):      # within 22.5 deg of vertical
                x = x0
            else:                                # diagonal
                side = max(abs(dx), abs(dy))
                x = x0 + side * (1 if dx >= 0 else -1)
                y = y0 + side * (1 if dy >= 0 else -1)
            x, y = self._clamp(x, y)
        elif event.state & 0x0001 and self.tool == "rect":
            x0, y0 = self.drag_start
            side = max(abs(x - x0), abs(y - y0))
            x = x0 + side * (1 if x >= x0 else -1)
            y = y0 + side * (1 if y >= y0 else -1)
            x, y = self._clamp(x, y)
        return x, y

    def _clear_preview(self):
        for item in self.preview:
            self.canvas.delete(item)
        self.preview = []

    def _press(self, event):
        if self.text_edit is not None:       # click elsewhere = finish the note
            self._commit_text()
            self.drag_start = None
            return
        if self.tool == "text":
            self.drag_start = None
            self._start_text(*self._point(event))
            return
        self.drag_start = self._point(event)
        self.stroke = [self.drag_start]
        self._clear_preview()
        if self.tool == "crop":
            self.crop_sel = None
            self._update_crop_buttons()

    def _drag(self, event):
        if self.drag_start is None:
            return
        if self.tall:   # scroll while dragging near the top/bottom edge
            if event.y < S(24):
                self.canvas.yview_scroll(-1, "units")
            elif event.y > self.canvas.winfo_height() - S(24):
                self.canvas.yview_scroll(1, "units")

        tool = self.tool
        if tool in ("pen", "highlight", "eraser"):
            p = self._point(event)
            last = self.stroke[-1]
            if abs(p[0] - last[0]) + abs(p[1] - last[1]) < 2:
                return
            self.stroke.append(p)
            if tool == "eraser":
                # show the eraser's path; marks under it disappear on release
                item = self.canvas.create_line(
                    *last, *p, fill="#FFFFFF", width=self._screen_width(),
                    capstyle="round", joinstyle="round", stipple="gray50")
            else:
                setting = self.settings[tool]
                item = self.canvas.create_line(
                    *last, *p, fill=setting["color"], width=self._screen_width(),
                    capstyle="round", joinstyle="round",
                    stipple="gray50" if tool == "highlight" else "")
            self.preview.append(item)
            return

        x0, y0 = self.drag_start
        x1, y1 = self._shape_end(event)
        self._clear_preview()
        if tool == "crop":
            self._draw_crop_selection(x0, y0, x1, y1)
            return
        setting = self.settings[tool]
        width = self._screen_width()
        if tool == "line":
            item = self.canvas.create_line(x0, y0, x1, y1, fill=setting["color"],
                                           width=width, capstyle="round")
        elif tool == "blur":
            # dashed outline of the area that will be blurred
            self.preview.append(self.canvas.create_rectangle(x0, y0, x1, y1, outline="#FFFFFF", width=2))
            item = self.canvas.create_rectangle(x0, y0, x1, y1, outline=ACCENT, width=2, dash=(6, 4))
        else:
            item = self.canvas.create_rectangle(x0, y0, x1, y1, outline=setting["color"], width=width)
        self.preview.append(item)

    def _release(self, event):
        if self.drag_start is None:
            return
        tool = self.tool
        s = self.scale
        if tool == "crop":
            x0, y0 = self.drag_start
            x1, y1 = self._point(event)
            self.drag_start = None
            self._clear_preview()
            if abs(x1 - x0) >= 5 and abs(y1 - y0) >= 5:
                self.crop_sel = (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
                self._draw_crop_selection(*self.crop_sel)
            self._update_crop_buttons()
            return
        self._clear_preview()

        if tool in ("pen", "highlight", "eraser"):
            p = self._point(event)
            if p != self.stroke[-1]:
                self.stroke.append(p)
            setting = self.settings[tool]
            op = {"type": tool, "color": setting["color"],
                  "width": max(1, round(self._screen_width() / s)),
                  "points": [self._img(x, y) for x, y in self.stroke]}
            if tool == "eraser" and not (self.ink_layer.getbbox() or self.hl_layer.getbbox()):
                # nothing to erase
                self.drag_start = None
                self.stroke = []
                return
        elif tool == "line":
            x0, y0 = self.drag_start
            x1, y1 = self._shape_end(event)
            self.drag_start = None
            if abs(x1 - x0) + abs(y1 - y0) < 3:      # just a click
                return
            setting = self.settings[tool]
            op = {"type": "line", "color": setting["color"],
                  "width": max(1, round(self._screen_width() / s)),
                  "points": [self._img(x0, y0), self._img(x1, y1)]}
        else:
            x0, y0 = self.drag_start
            x1, y1 = self._shape_end(event)
            box = self._img_box(x0, y0, x1, y1)
            self.drag_start = None
            if box[2] - box[0] < 3 or box[3] - box[1] < 3:
                return
            setting = self.settings[tool]
            if tool == "blur":
                op = {"type": "blur", "box": box, "radius": BLUR_RADII[setting["size"]]}
            else:
                op = {"type": tool, "box": box, "color": setting["color"],
                      "width": max(1, round(self._screen_width() / s))}

        self.drag_start = None
        self.stroke = []
        self.ops.append(op)
        self._apply(op)
        self._compose()
        self.redraw()

    # ---- text -------------------------------------------------------

    def _text_font(self):
        return (UI_FAMILY, -self._screen_width(), "bold")

    def _start_text(self, x, y):
        """Open a small typing box where the user clicked."""
        box = tk.Text(self.canvas, width=10, height=1, wrap="none", undo=True,
                      relief="solid", bd=1, padx=S(4), pady=S(2),
                      highlightthickness=0, insertwidth=2)
        # keep typed letters away from the window's shortcuts (H, P, Enter...)
        box.bindtags((str(box), "Text", "all"))
        box.bind("<Return>", lambda e: (self._commit_text(), "break")[1])
        box.bind("<KP_Enter>", lambda e: (self._commit_text(), "break")[1])
        box.bind("<Shift-Return>", lambda e: (box.insert("insert", "\n"), self._fit_text_widget(), "break")[2])
        box.bind("<Escape>", lambda e: (self._discard_text(), "break")[1])
        box.bind("<KeyPress>", lambda e: self.win.after_idle(self._fit_text_widget))
        box.bind("<KeyRelease>", lambda e: self._fit_text_widget())
        pad_x = int(box.cget("padx")) + 1
        pad_y = int(box.cget("pady")) + 1
        item = self.canvas.create_window(x - pad_x, y - pad_y, window=box, anchor="nw")
        self.text_edit = (box, item, self._img(x, y))
        self._style_text_widget()
        box.focus_set()

    def _style_text_widget(self):
        if self.text_edit is None or self.tool != "text":
            return
        box = self.text_edit[0]
        color = self.settings["text"]["color"]
        light = sum(_hex_rgb(color)) > 600
        bg = "#FFFFFF" if not light else "#1F2430"
        if not self.text_box_var.get():
            bg = "#F4F6F8" if not light else "#2A2F3A"
        box.configure(font=self._text_font(), fg=color, bg=bg, insertbackground=color)
        self._fit_text_widget()

    def _fit_text_widget(self):
        """Grow the typing box (in pixels) to fit what has been typed."""
        if self.text_edit is None:
            return
        import tkinter.font as tkfont
        box, item, _ = self.text_edit
        font = tkfont.Font(font=box.cget("font"))
        lines = box.get("1.0", "end-1c").split("\n")
        extra = 2 * (int(box.cget("padx")) + int(box.cget("bd"))) + font.measure("MM")
        width = max(font.measure("M" * 6), max(font.measure(l) for l in lines)) + extra
        height = (font.metrics("linespace") * len(lines)
                  + 2 * (int(box.cget("pady")) + int(box.cget("bd"))) + 2)
        self.canvas.itemconfigure(item, width=width, height=height)
        box.xview_moveto(0)

    def _commit_text(self):
        """Turn the note being typed into an annotation."""
        if self.text_edit is None:
            return
        box, item, pos = self.text_edit
        text = box.get("1.0", "end-1c").rstrip()
        self.text_edit = None
        self.canvas.delete(item)
        box.destroy()
        self.win.focus_set()
        if not text.strip():
            return
        op = {"type": "text", "pos": pos, "text": text,
              "color": self.settings["text"]["color"],
              "size": max(6, round(self._screen_width() / self.scale))
                      if self.tool == "text" else
                      max(6, round(self.SIZES["text"][self.settings["text"]["size"]] / self.scale)),
              "box": bool(self.text_box_var.get())}
        self.ops.append(op)
        self._apply(op)
        self._compose()
        self.redraw()

    def _discard_text(self):
        if self.text_edit is None:
            return
        box, item, _ = self.text_edit
        self.text_edit = None
        self.canvas.delete(item)
        box.destroy()
        self.win.focus_set()

    # ---- crop -------------------------------------------------------

    def _draw_crop_selection(self, x0, y0, x1, y1):
        """Dim everything outside the selection and outline it."""
        x0, x1 = sorted((x0, x1))
        y0, y1 = sorted((y0, y1))
        L, T = self.off_x, self.off_y
        # darken the screenshot outside the selection (drawn as an image,
        # so it looks the same on every system)
        if self._dim_img is None:
            self._dim_img = Image.blend(self.view_img.convert("RGB"),
                                        Image.new("RGB", self.view_img.size, (0, 0, 0)), 0.55)
        sel = (int(x0 - L), int(y0 - T), int(x1 - L), int(y1 - T))
        shaded = self._dim_img.copy()
        if sel[2] > sel[0] and sel[3] > sel[1]:
            shaded.paste(self.view_img.crop(sel), sel[:2])
        self._dim_photo = ImageTk.PhotoImage(shaded)
        self.preview.append(self.canvas.create_image(L, T, anchor="nw", image=self._dim_photo))
        self.preview.append(self.canvas.create_rectangle(
            x0, y0, x1, y1, outline="#FFFFFF", width=2))
        self.preview.append(self.canvas.create_rectangle(
            x0, y0, x1, y1, outline=ACCENT, width=2, dash=(6, 4)))
        ix0, iy0, ix1, iy1 = self._img_box(x0, y0, x1, y1)
        self.crop_hint.configure(text=f"Selection: {ix1 - ix0} \u00d7 {iy1 - iy0} px", fg=TEXT)

    def _update_crop_buttons(self):
        state = "normal" if self.crop_sel else "disabled"
        self.crop_apply.config(state=state)
        self.crop_cancel.config(state=state)
        if not self.crop_sel:
            self.crop_hint.configure(text="Drag over the part you want to keep", fg=MUTED)

    def clear_crop_selection(self):
        if self.crop_sel is not None:
            self.crop_sel = None
            self._clear_preview()
        if hasattr(self, "crop_apply"):
            self._update_crop_buttons()

    def apply_crop(self):
        if not self.crop_sel:
            return
        box = self._img_box(*self.crop_sel)
        self.clear_crop_selection()
        if box[2] - box[0] < 2 or box[3] - box[1] < 2 or box == self.crop_box:
            return
        self.ops.append({"type": "crop", "box": box})
        self.rebuild()
        self.select_tool(self.prev_tool)    # back to drawing
        self._relayout()

    def _on_enter(self, event=None):
        if self.crop_sel:
            self.apply_crop()
        else:
            self.save()

    def _on_escape(self, event=None):
        if self.crop_sel:
            self.clear_crop_selection()
        else:
            self.cancel()

    # ---- undo / save ------------------------------------------------

    def undo(self):
        if self.text_edit is not None:       # undo while typing = drop that note
            self._discard_text()
            return
        if self.ops:
            was_crop = self.ops.pop()["type"] == "crop"
            self.rebuild()
            if was_crop:
                self._relayout()
            else:
                self.redraw()

    def clear(self):
        self._discard_text()
        if self.ops:
            had_crop = any(op["type"] == "crop" for op in self.ops)
            self.ops.clear()
            self.rebuild()
            if had_crop:
                self._relayout()
            else:
                self.redraw()

    def save(self):
        self._commit_text()
        result = self.working.copy()
        self.win.grab_release()
        if self.on_save(result) is False:     # e.g. user cancelled "Save as"
            self.win.grab_set()
            return
        print(f"Saved with {len(self.ops)} annotation(s)")
        self.win.destroy()

    def copy(self):
        """Copy the screenshot (with annotations) to the clipboard. The
        window stays open, so it can still be saved afterwards."""
        self._commit_text()
        try:
            ok = copy_image_to_clipboard(self.working)
        except Exception as e:
            print(f"Copy failed: {e}")
            ok = False
        if not ok:
            messagebox.showerror("Copy failed",
                                 "Could not copy the image to the clipboard. Please try again.",
                                 parent=self.win)
            return
        self.copied = True
        self.copied_ops = len(self.ops)
        self.copy_button.config(text="Copied")
        self.win.after(1500, lambda: self.copy_button.winfo_exists()
                       and self.copy_button.config(text="Copy"))
        status_var.set(f"Screenshot copied to the clipboard - paste it with {KEY}+V")

    def cancel(self):
        self._commit_text()
        # nothing to lose if it was copied after the last change
        copied_latest = self.copied and getattr(self, "copied_ops", -1) == len(self.ops)
        if self.ops and not copied_latest and not messagebox.askyesno(
                "Discard screenshot?",
                "Close without saving? The screenshot will be discarded.",
                parent=self.win):
            return
        self.win.grab_release()
        self.win.destroy()
        status_var.set("Copied to the clipboard - not saved to the folder"
                       if self.copied else "Screenshot discarded")


def _hex_rgb(color):
    color = color.lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def draw_path(layer, points, rgba, width):
    """Stroke a freehand path with round ends onto an RGBA layer. With a
    fully transparent colour it erases (pixels are replaced, not blended)."""
    d = ImageDraw.Draw(layer)
    if len(points) > 1:
        d.line(points, fill=rgba, width=width, joint="curve")
    r = width / 2
    for x, y in (points if len(points) <= 2 else (points[0], points[-1])):
        d.ellipse([x - r, y - r, x + r, y + r], fill=rgba)
    if rgba[3] == 0 and len(points) > 2:
        # round joints along the eraser path so it leaves no gaps
        for x, y in points[1:-1]:
            d.ellipse([x - r, y - r, x + r, y + r], fill=rgba)


def draw_stroke(image, points, color, width):
    """Freehand pen stroke with round ends and joins."""
    d = ImageDraw.Draw(image)
    rgb = _hex_rgb(color)
    if len(points) > 1:
        d.line(points, fill=rgb, width=width, joint="curve")
    r = width / 2
    for x, y in (points[0], points[-1]) if len(points) > 1 else points:
        d.ellipse([x - r, y - r, x + r, y + r], fill=rgb)


def draw_highlight(image, points, color, width):
    """Highlighter: colours the background like a real marker while the
    text underneath stays dark and readable (multiply blend)."""
    pad = width
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    box = (max(0, int(min(xs) - pad)), max(0, int(min(ys) - pad)),
           min(image.width, int(max(xs) + pad) + 1), min(image.height, int(max(ys) + pad) + 1))
    if box[2] <= box[0] or box[3] <= box[1]:
        return
    region = image.crop(box)
    mask = Image.new("L", region.size, 0)
    md = ImageDraw.Draw(mask)
    local = [(x - box[0], y - box[1]) for x, y in points]
    if len(local) > 1:
        md.line(local, fill=255, width=width, joint="curve")
    r = width / 2
    for x, y in (local[0], local[-1]) if len(local) > 1 else local:
        md.ellipse([x - r, y - r, x + r, y + r], fill=255)
    tint = Image.new("RGB", region.size, _hex_rgb(color))
    marked = ImageChops.multiply(region, tint)
    region.paste(marked, (0, 0), mask)
    image.paste(region, box)


def draw_shape(image, kind, box, color, width):
    d = ImageDraw.Draw(image)
    rgb = _hex_rgb(color) + ((255,) if image.mode == "RGBA" else ())
    x0, y0, x1, y1 = box
    if kind == "rect":
        d.rectangle(box, outline=rgb, width=width)


TEXT_FONT_FILES = ["segoeuib.ttf", "arialbd.ttf", "seguisb.ttf", "segoeui.ttf", "arial.ttf",
                   "/System/Library/Fonts/Supplemental/Arial Bold.ttf", "/Library/Fonts/Arial Bold.ttf",
                   "/System/Library/Fonts/HelveticaNeue.ttc", "/System/Library/Fonts/Helvetica.ttc",
                   "DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
_TEXT_FONT_CACHE = {}


def text_font(size):
    """Bold UI font at the given pixel size (Segoe UI on Windows)."""
    size = max(6, int(size))
    if size not in _TEXT_FONT_CACHE:
        from PIL import ImageFont
        font = None
        for name in TEXT_FONT_FILES:
            try:
                font = ImageFont.truetype(name, size)
                break
            except OSError:
                continue
        if font is None:
            try:
                font = ImageFont.load_default(size=size)
            except TypeError:                  # very old Pillow
                font = ImageFont.load_default()
        _TEXT_FONT_CACHE[size] = font
    return _TEXT_FONT_CACHE[size]


def draw_text(layer, pos, text, color, size, box=True):
    """A text note. With box=True it sits on a rounded background in a
    contrasting colour, with an outline in the text colour."""
    d = ImageDraw.Draw(layer)
    font = text_font(size)
    rgb = _hex_rgb(color)
    spacing = max(2, int(size * 0.25))
    x, y = pos
    bbox = d.multiline_textbbox((x, y), text, font=font, spacing=spacing)
    if box:
        pad = max(4, int(size * 0.35))
        light = sum(rgb) > 600
        fill = (31, 36, 48, 235) if light else (255, 255, 255, 235)
        d.rounded_rectangle([bbox[0] - pad, bbox[1] - pad, bbox[2] + pad, bbox[3] + pad],
                            radius=pad, fill=fill, outline=rgb + (255,),
                            width=max(1, size // 14))
    d.multiline_text((x, y), text, font=font, fill=rgb + (255,), spacing=spacing)


def draw_line(image, points, color, width):
    """Straight line with round ends."""
    d = ImageDraw.Draw(image)
    rgb = _hex_rgb(color) + ((255,) if image.mode == "RGBA" else ())
    (x0, y0), (x1, y1) = points
    d.line([(x0, y0), (x1, y1)], fill=rgb, width=width)
    if width > 2:
        r = width / 2
        for x, y in ((x0, y0), (x1, y1)):
            d.ellipse([x - r, y - r, x + r, y + r], fill=rgb)


SIZE_DOT_CACHE = {}


def size_dot_image(index, selected):
    """Thickness choice: a dot that grows with the size."""
    key = (index, selected)
    if key not in SIZE_DOT_CACHE:
        n = S(26) * 4
        img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        if selected:
            d.rounded_rectangle([0, 0, n - 1, n - 1], radius=n // 5,
                                fill=_hex_rgb(ACCENT_SOFT), outline=_hex_rgb(ACCENT), width=6)
        r = n * (0.08 + 0.065 * index)
        c = n / 2
        d.ellipse([c - r, c - r, c + r, c + r], fill=_hex_rgb(ACCENT if selected else TEXT))
        SIZE_DOT_CACHE[key] = ImageTk.PhotoImage(img.resize((n // 4, n // 4), Image.LANCZOS))
    return SIZE_DOT_CACHE[key]


# ============================================================
# Full-Page (Scrolling) Capture
# ============================================================
# For long pages and single-page applications that don't fit on screen:
# the user selects the area that scrolls (e.g. the web page inside the
# browser). The app scrolls it to the top, then scrolls down step by
# step with the mouse wheel, grabbing the area each time, and stitches
# the grabs into one tall image. Fixed headers/footers (common in
# single-page apps) are detected so they appear only once.
# While it runs, the main window becomes a small progress bar with a
# Stop button (Esc also stops); what was captured so far is kept.

SCROLL_MAX_STEPS = 80          # safety limit (infinite-scroll pages)
SCROLL_MAX_HEIGHT = 20000      # pixels
SCROLL_MAX_SECONDS = 180       # overall time limit
SCROLL_SETTLE_MS = 350         # wait for smooth scrolling / lazy content


def _row_signatures(image):
    """One small byte string per pixel row (rows are compared exactly)."""
    gray = image.convert("L")
    width = max(1, gray.width // 4)
    small = gray.resize((width, gray.height), Image.BOX)
    data = small.tobytes()
    return [data[i * width:(i + 1) * width] for i in range(small.height)]


def _is_plain(row):
    return max(row) - min(row) < 12


def _rows_similar(x, y):
    """Two pixel rows are 'similar' if they match except for a small part
    of their width - e.g. a blinking cursor, spinner or clock."""
    if x == y:
        return True
    limit = max(2, len(x) // 8)             # up to ~12% of the row width
    different = 0
    for p, q in zip(x, y):
        if p - q > 24 or q - p > 24:
            different += 1
            if different > limit:
                return False
    return True


def _nearly_same(a, b):
    """True if two grabs are the same apart from small animations."""
    if len(a) != len(b):
        return False
    allowed = max(2, len(a) // 60)          # up to ~1.7% of the rows
    different = 0
    for x, y in zip(a, b):
        if not _rows_similar(x, y):
            different += 1
            if different > allowed:
                return False
    return True


def _fixed_band(a, b, indexes):
    """Rows that didn't move between two grabs, scanning from one edge
    (a fixed header or footer). Short runs of changed rows - e.g. an
    animated icon in the header - are tolerated."""
    band = 0
    misses = 0
    for count, i in enumerate(indexes, start=1):
        if _rows_similar(a[i], b[i]):
            band = count
            misses = 0
        else:
            misses += 1
            if misses > 4:
                break
    return band


def _fixed_bands(a, b):
    height = len(a)
    limit = int(height * 0.35)
    top = _fixed_band(a, b, range(limit))
    bottom = _fixed_band(a, b, range(height - 1, height - 1 - limit, -1))
    return top, bottom


def _find_scroll_shift(a_rows, b_rows, top, bottom, expected=None):
    """How many pixels the content moved up between grab a and grab b,
    looking only at the scrolling middle band. None if it can't be found."""
    height = len(a_rows)
    a = a_rows[top:height - bottom]
    b = b_rows[top:height - bottom]
    m = len(a)
    if m < 20:
        return None

    index = {}
    for i, row in enumerate(a):
        index.setdefault(row, []).append(i)

    # distinctive rows near the top of b, looked up in a
    candidates, seen = set(), set()
    for j, row in enumerate(b[:max(1, int(m * 0.7))]):
        if _is_plain(row) or row in seen or len(index.get(row, ())) > 3:
            continue
        seen.add(row)
        for i in index.get(row, ()):
            if 0 < i - j < m:
                candidates.add(i - j)
        if len(seen) >= 24:
            break

    best = None
    min_overlap = max(10, int(m * 0.06))
    for shift in candidates:
        overlap = m - shift
        if overlap < min_overlap:
            continue
        allowed = max(2, int(overlap * 0.05))
        bad = 0
        for k in range(overlap):
            if not _rows_similar(a[shift + k], b[k]):   # tolerate small animations
                bad += 1
                if bad > allowed:
                    break
        else:
            score = (bad, abs(shift - expected) if expected else 0, shift)
            if best is None or score < best[0]:
                best = (score, shift)
    return best[1] if best else None


class ScrollingCapture:
    """Scrolls a screen area and stitches the grabs (runs on Tk timers)."""

    active = None

    def __init__(self, bbox, on_done, on_progress=None):
        self.bbox = bbox
        self.on_done = on_done            # on_done(image or None, message)
        self.on_progress = on_progress or (lambda text: None)
        x1, y1, x2, y2 = bbox
        # wheel over the right edge of the area (usually the scrollbar),
        # so hover effects on the page don't change what is captured
        self.mouse_at = (max(x1, x2 - 12), (y1 + y2) // 2)
        self.saved_cursor = None
        self.taken_at = datetime.now()
        self.started = time.monotonic()
        self.first = None
        self.prev_rows = None
        self.pieces = []
        self.last = None
        self.top = self.bottom = None
        self.height = 0
        self.steps = 0
        self.notches = 1
        self.px_per_notch = None
        self.top_tries = 0
        self.stop_requested = False
        self.stopped_early = False
        self.finished = False
        self.method = "message"          # "message" (pointer stays free) or "cursor"
        self.scroll_seen = False         # has scrolling visibly worked yet?

    # ---- Windows input (overridable for testing) -----------------------

    def _wheel(self, notches):
        """Scroll the page. Positive = down."""
        if IS_MAC:                      # "message": scroll event at the point; "cursor": move there too
            mac_scroll(self.mouse_at[0], self.mouse_at[1], notches, self.method != "message")
            return
        if self.method == "message":
            self._wheel_message(notches)
        else:
            self._wheel_cursor(notches)

    def _wheel_message(self, notches):
        """Send mouse-wheel messages straight to the window under the
        scroll point. The mouse pointer is NOT moved, so the user can
        still click Stop."""
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        user32.WindowFromPoint.argtypes = [wintypes.POINT]
        user32.WindowFromPoint.restype = wintypes.HWND
        user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                        wintypes.WPARAM, wintypes.LPARAM]
        x, y = int(self.mouse_at[0]), int(self.mouse_at[1])
        hwnd = user32.WindowFromPoint(wintypes.POINT(x, y))
        if not hwnd:
            return
        lparam = ((y & 0xFFFF) << 16) | (x & 0xFFFF)          # screen position
        remaining = int(-120 * notches)
        while remaining:
            delta = max(-1200, min(1200, remaining))           # 10 notches per message
            remaining -= delta
            wparam = (delta & 0xFFFF) << 16
            user32.PostMessageW(hwnd, 0x020A, wparam, lparam)  # WM_MOUSEWHEEL

    def _wheel_cursor(self, notches):
        """Fallback for apps that ignore wheel messages: move the pointer
        there and turn the wheel (Esc stops, since the pointer is busy)."""
        user32 = ctypes.windll.user32
        user32.SetCursorPos(int(self.mouse_at[0]), int(self.mouse_at[1]))
        user32.mouse_event(0x0800, 0, 0, int(-120 * notches), 0)   # MOUSEEVENTF_WHEEL

    def _grab(self):
        return grab_screen(self.bbox).convert("RGB")

    def _escape_pressed(self):
        """Esc is down now, or was pressed since the last check (works
        whichever window has the focus)."""
        try:
            if IS_MAC:
                return mac_escape_down()
            return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8001)
        except Exception:
            return False

    def _watch_escape(self):
        """Check for Esc every 50 ms for the whole capture."""
        if self.finished:
            return
        if self._escape_pressed():
            self.stop()
        root.after(50, self._watch_escape)

    def _save_cursor(self):
        try:
            if IS_MAC:
                self.saved_cursor = mac_pointer()
                return
            from ctypes import wintypes
            point = wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(point))
            self.saved_cursor = (point.x, point.y)
        except Exception:
            self.saved_cursor = None

    def _restore_cursor(self):
        try:
            if self.saved_cursor and IS_MAC:
                mac["Quartz"].CGWarpMouseCursorPosition(self.saved_cursor)
            elif self.saved_cursor:
                ctypes.windll.user32.SetCursorPos(*self.saved_cursor)
        except Exception:
            pass

    # ---- helpers -----------------------------------------------------------

    def _later(self, ms, func, *args):
        """Run func after ms; any error ends the capture with a message
        instead of silently stopping."""
        def run():
            if self.finished:
                return
            try:
                func(*args)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._finish(error=f"Full-page capture failed: {e}")
        root.after(ms, run)

    def _should_stop(self):
        if self.stop_requested:
            return True
        if time.monotonic() - self.started > SCROLL_MAX_SECONDS:
            return True
        return False

    def stop(self):
        """Stop button / Esc: finish now and save what has been captured."""
        if not self.stop_requested and not self.finished:
            self.stop_requested = True
            self.on_progress("Stopping - saving...")

    # ---- flow ------------------------------------------------------------

    def start(self):
        ScrollingCapture.active = self
        self._escape_pressed()        # clear an Esc press from before the capture
        self._watch_escape()
        self._save_cursor()
        self.top_tries = 0
        self.on_progress("Scrolling to the top...")
        self._later(0, self._begin)

    def _begin(self):
        self.prev_rows = _row_signatures(self._grab())
        self._to_top()

    def _stable_grab(self, then, tries=0, previous=None):
        """Grab until the picture stops changing (scrolling has settled)."""
        image = self._grab()
        rows = _row_signatures(image)
        if previous is not None and (_nearly_same(rows, previous[1]) or tries >= 5):
            then(image, rows)
            return
        self._later(120, self._stable_grab, then, tries + 1, (image, rows))

    def _to_top(self):
        """Scroll up until nothing changes any more."""
        if self._should_stop():
            # stopped before scrolling down: save what's on screen now
            self.stopped_early = True
            self.first = self.last = self._grab()
            return self._finish()
        self._wheel(-25)
        self._later(SCROLL_SETTLE_MS, self._stable_grab, self._after_up)

    def _after_up(self, image, rows):
        self.top_tries += 1
        if not _nearly_same(rows, self.prev_rows):
            self.scroll_seen = True
        if _nearly_same(rows, self.prev_rows) or self.top_tries >= 20:
            self.first = image
            self.last = image
            self.prev_rows = rows
            self.height = image.height
            self.on_progress("Capturing...  screen 1")
            self._step_down()
        else:
            self.prev_rows = rows
            self._to_top()

    def _step_down(self):
        if self._should_stop():
            self.stopped_early = True
            return self._finish()
        if self.steps >= SCROLL_MAX_STEPS or self.height >= SCROLL_MAX_HEIGHT:
            self.stopped_early = True
            return self._finish()
        self._wheel(self.notches)
        self._later(SCROLL_SETTLE_MS, self._stable_grab, self._after_down)

    def _after_down(self, image, rows):
        if _nearly_same(rows, self.prev_rows):
            if self.steps == 0 and not self.scroll_seen and self.method == "message":
                # this app ignored the wheel messages: start over, moving the
                # pointer to scroll instead (Esc still stops)
                print("Full-page capture: switching to pointer scrolling")
                self.method = "cursor"
                self.top_tries = 0
                self.on_progress("Scrolling to the top...")
                self.prev_rows = rows
                return self._to_top()
            return self._finish()                        # reached the end
        self.scroll_seen = True

        if self.top is None:
            self.top, self.bottom = _fixed_bands(self.prev_rows, rows)

        expected = self.px_per_notch * self.notches if self.px_per_notch else None
        shift = _find_scroll_shift(self.prev_rows, rows, self.top, self.bottom, expected)
        if shift is None:
            if self.steps == 0:
                return self._finish(error=(
                    "The selected area didn't scroll in a way that could be stitched.\n"
                    "Select only the scrolling page content (not the whole browser), "
                    "and make sure the page itself scrolls with the mouse wheel."))
            self.stopped_early = True
            return self._finish()

        # keep the newly revealed rows at the bottom of the scrolling band
        h = image.height
        self.pieces.append(image.crop((0, h - self.bottom - shift, image.width, h - self.bottom)))
        self.height += shift
        self.last = image
        self.prev_rows = rows
        self.steps += 1
        self.on_progress(f"Capturing...  screen {self.steps + 1}  ·  {self.height} px")

        # after the first step: scroll ~65% of the visible band each time
        if self.px_per_notch is None:
            self.px_per_notch = max(1, shift / self.notches)
            band = h - self.top - self.bottom
            self.notches = max(1, int(band * 0.65 / self.px_per_notch))
        self._step_down()

    def _finish(self, cancelled=False, error=None):
        if self.finished:
            return
        self.finished = True
        ScrollingCapture.active = None
        self._restore_cursor()
        if error:
            self.on_done(None, error)
            return
        if cancelled or self.first is None:
            self.on_done(None, "Full-page capture cancelled")
            return

        w, h = self.first.size
        bottom = self.bottom or 0
        head = self.first.crop((0, 0, w, h - bottom))
        total = head.height + sum(p.height for p in self.pieces) + bottom
        result = Image.new("RGB", (w, total), "white")
        result.paste(head, (0, 0))
        y = head.height
        for piece in self.pieces:
            result.paste(piece, (0, y))
            y += piece.height
        if bottom:
            result.paste(self.last.crop((0, h - bottom, w, h)), (0, y))

        screens = self.steps + 1
        note = " (stopped early)" if self.stopped_early else ""
        self.on_done(result, f"{screens} screen{'s' if screens != 1 else ''} stitched, "
                             f"{w} x {total} px{note}")


def full_page_capture():
    """Button: select the scrolling area, then capture the whole page."""
    if recorder is not None or ScrollingCapture.active is not None:
        return
    if IS_MAC and not mac_accessibility_allowed(ask=True):
        messagebox.showinfo(
            "Allow scrolling",
            "Full Page scrolls the page for you, which macOS only allows after you agree.\n\n"
            "Open System Settings > Privacy & Security > Accessibility, switch on "
            "Screen Capture Utility, then try again.", parent=root)
        return
    if not (IS_WINDOWS or IS_MAC):
        messagebox.showinfo("Not available", "Full-page capture needs Windows or macOS.", parent=root)
        return
    if not ensure_save_folder():
        status_var.set("No save folder selected - capture cancelled")
        return
    start_region_selection(
        start_scrolling_capture,
        hint="Drag over the page content that scrolls (not the address bar or tabs)  "
             "·  then don't touch the mouse  ·  Esc cancels")


def start_scrolling_capture(x1, y1, x2, y2):
    if y2 - y1 < 60 or x2 - x1 < 60:
        root.deiconify()
        status_var.set("Selection too small for a full-page capture")
        return

    area = {"left": x1, "top": y1, "width": x2 - x1, "height": y2 - y1}

    def progress(text):
        scroll_label.config(text=text)

    def done(image, message):
        root.unbind("<Escape>")
        exit_compact_mode(scroll_bar)
        if image is None:
            status_var.set(message)
            if "failed" in message or "couldn't" in message or "didn't" in message:
                messagebox.showwarning("Full-page capture", message, parent=root)
            return
        print(f"Full-page capture: {message}")
        deliver_capture(image, "Full page", capture.taken_at)
        root.after(300, lambda: status_var.set(
            f"{status_var.get()}\nFull page: {message}"))

    capture = ScrollingCapture((x1, y1, x2, y2), done, progress)
    scroll_stop_button.config(command=capture.stop, state="normal")
    root.bind("<Escape>", lambda e: capture.stop())
    enter_compact_mode(scroll_bar, area, "Capturing full page")
    capture.start()


# ============================================================
# Compare Images
# ============================================================
# Image 1 is the SOURCE; Image 2 is checked against it. Every difference is
# marked in red in BOTH images: changed text, changed or moved
# graphics, content that is missing from Image 2 and content that was added.
#
#   1. If the images are the same shape but a different size (e.g. captured
#      at another zoom level), Image 2 is resized to match Image 1.
#   2. Auto-align: separately captured regions are often a few pixels apart,
#      so the best small shift between them is found first.
#   3. Line by line: both images are cut into horizontal bands (lines of
#      text, rows of controls) and the bands of Image 2 are matched to those
#      of Image 1 in order, like a text compare. So when a line is removed or
#      added, the lines below it are still matched with their counterparts
#      instead of all being reported as changed.
#   4. Matched lines are compared pixel by pixel (with a tolerance, so
#      anti-aliasing and JPEG noise are ignored). Lines found in only one
#      image are reported as missing from / added in Image 2.
#   5. The changed pixels are grouped into numbered differences. The images
#      are shown exactly as they are (nothing is painted over the content):
#      each difference gets a red outline with a small gap around it, a bar
#      in the margin beside the same rows and its number in that margin, in
#      Image 1 and in Image 2. When content is missing, a dashed red line
#      marks the spot where it should be.
#   Pictures without clear lines (photos, charts) are compared pixel by
#   pixel; whichever method explains the differences with less red wins.
# Pure Pillow - no extra packages.

COMPARE_RED = (230, 0, 0)
COMPARE_HEX = "#E60000"
COMPARE_MIN_PIXELS = 3               # smaller specks are ignored (noise)
COMPARE_SIG = (96, 8)                # size of the small "fingerprint" of a line
KIND_CHANGED = "Changed"
KIND_MISSING = "Missing in Image 2"
KIND_ADDED = "Added in Image 2"


@dataclass
class Difference:
    kind: str                    # KIND_CHANGED / KIND_MISSING / KIND_ADDED
    box1: tuple                  # (x0, y0, x1, y1) in Image 1
    box2: tuple                  # (x0, y0, x1, y1) in Image 2
    marker1: bool = False        # True: not in Image 1 - box1 is a line where it would be
    marker2: bool = False        # True: not in Image 2 - box2 is a line where it should be


@dataclass
class CompareResult:
    image1: Image.Image          # Image 1 (the source)
    image2: Image.Image          # Image 2 (resized to Image 1 if needed)
    mask1: Image.Image           # "L": 255 = differing pixel in Image 1
    mask2: Image.Image           # "L": 255 = differing pixel in Image 2
    diffs: list                  # [Difference, ...] in reading order of Image 1
    changed_pixels: int
    compared_pixels: int
    notes: list
    method: str = "pixels"       # "pixels" or "lines"

    def count(self, kind):
        return sum(1 for d in self.diffs if d.kind == kind)


def _max_channel_diff(a, b):
    """Per-pixel difference: the largest change in R, G or B (so a colour
    change is found even when brightness stays the same)."""
    r, g, bl = ImageChops.difference(a, b).split()
    return ImageChops.lighter(ImageChops.lighter(r, g), bl)


def _changed_mask(a, b, tolerance):
    """255 where the images really differ. A pixel only counts as changed
    if it matches nothing in the other image's 3x3 neighbourhood (by more
    than the tolerance), so anti-aliasing, JPEG noise and sub-pixel shifts
    are ignored while real edits are found. Tolerance 0 = exact compare."""
    if tolerance <= 0:
        return _max_channel_diff(a, b).point(lambda v: 255 if v else 0)
    amin, amax = a.filter(ImageFilter.MinFilter(3)), a.filter(ImageFilter.MaxFilter(3))
    bmin, bmax = b.filter(ImageFilter.MinFilter(3)), b.filter(ImageFilter.MaxFilter(3))

    def max_channel(img):
        r, g, bl = img.split()
        return ImageChops.lighter(ImageChops.lighter(r, g), bl)
    a_out = max_channel(ImageChops.lighter(ImageChops.subtract(a, bmax), ImageChops.subtract(bmin, a)))
    b_out = max_channel(ImageChops.lighter(ImageChops.subtract(b, amax), ImageChops.subtract(amin, b)))
    return ImageChops.lighter(a_out, b_out).point(lambda v: 255 if v > tolerance else 0)


def _shift_score(a, b, dx, dy, min_overlap=0.5):
    """Mean grey-level difference with image a placed at (dx, dy) over b,
    or None if they would overlap too little."""
    aw, ah = a.size
    bw, bh = b.size
    x0, y0 = max(0, dx), max(0, dy)
    x1, y1 = min(bw, dx + aw), min(bh, dy + ah)
    if x1 - x0 < min(aw, bw) * min_overlap or y1 - y0 < min(ah, bh) * min_overlap:
        return None
    diff = ImageChops.difference(a.crop((x0 - dx, y0 - dy, x1 - dx, y1 - dy)),
                                 b.crop((x0, y0, x1, y1)))
    return ImageStat.Stat(diff).mean[0]


def _best_shift(score, cx, cy, radius):
    best = None
    for dy in range(cy - radius, cy + radius + 1):
        for dx in range(cx - radius, cx + radius + 1):
            s = score(dx, dy)
            if s is None:
                continue
            key = (round(s, 4), dx * dx + dy * dy)        # ties: smallest shift
            if best is None or key < best[0]:
                best = (key, dx, dy)
    return best


def find_alignment(a, b):
    """Shift (dx, dy) that lines up Image 1 (a) with Image 2 (b): pixel (x, y)
    of a matches pixel (x + dx, y + dy) of b. (0, 0) unless shifting is
    clearly better."""
    ga, gb = a.convert("L"), b.convert("L")
    biggest = max(ga.width, ga.height, gb.width, gb.height)

    def scaled(img, f):
        if f <= 1:
            return img
        return img.resize((max(1, round(img.width / f)), max(1, round(img.height / f))), Image.BOX)

    # 1) coarse: small copies (about 400 px), +-10 steps
    f1 = max(1.0, biggest / 400)
    sa, sb = scaled(ga, f1), scaled(gb, f1)
    coarse = _best_shift(lambda dx, dy: _shift_score(sa, sb, dx, dy), 0, 0, 10)
    if coarse is None:
        return 0, 0

    # 2) medium: about 1600 px
    f2 = max(1.0, biggest / 1600)
    ma, mb = scaled(ga, f2), scaled(gb, f2)
    cx, cy = round(coarse[1] * f1 / f2), round(coarse[2] * f1 / f2)
    medium = _best_shift(lambda dx, dy: _shift_score(ma, mb, dx, dy), cx, cy,
                         int(math.ceil(f1 / f2)) + 1)
    zero = _shift_score(ma, mb, 0, 0)
    if medium is None or zero is None:
        return 0, 0
    dx, dy = round(medium[1] * f2), round(medium[2] * f2)

    # 3) full resolution, on the most detailed parts of the image (text and
    #    edges give a precise match; plain areas or lines running in one
    #    direction don't), so it stays fast even for 4K screenshots
    if f2 > 1:
        r = int(math.ceil(f2)) + 1
        x0, y0 = max(0, dx) + r, max(0, dy) + r
        x1 = min(gb.width, dx + ga.width) - r
        y1 = min(gb.height, dy + ga.height) - r
        if x1 - x0 > 60 and y1 - y0 > 60:
            tile = min(600, max(60, (x1 - x0) // 4), max(60, (y1 - y0) // 4))
            edges = gb.filter(ImageFilter.FIND_EDGES) if biggest <= 2000 else \
                scaled(gb, f2).filter(ImageFilter.FIND_EDGES)
            ef = 1 if biggest <= 2000 else f2
            tiles = []
            for ty in range(y0, y1 - tile + 1, tile):
                for tx in range(x0, x1 - tile + 1, tile):
                    box = tuple(int(v / ef) for v in (tx, ty, tx + tile, ty + tile))
                    tiles.append((ImageStat.Stat(edges.crop(box)).mean[0], (tx, ty, tx + tile, ty + tile)))
            tiles.sort(reverse=True)
            windows = [(w, gb.crop(w)) for _, w in tiles[:4]]

            def window_score(sx, sy):
                return sum(ImageStat.Stat(ImageChops.difference(
                    ga.crop((w[0] - sx, w[1] - sy, w[2] - sx, w[3] - sy)), bw)).mean[0]
                    for w, bw in windows)
            if windows:
                fine = _best_shift(window_score, dx, dy, r)
                if fine is not None:
                    dx, dy = fine[1], fine[2]

    # Only shift when it really is a better match
    best = _shift_score(ma, mb, round(dx / f2), round(dy / f2))
    if best is None or best >= zero * 0.9 or zero - best < 0.3:
        return 0, 0
    return dx, dy


def find_regions(mask):
    """Group changed pixels into boxes [(x0, y0, x1, y1), ...]."""
    W, H = mask.size
    cell = int(min(15, max(4, round(max(W, H) / 500))))
    gw, gh = int(math.ceil(W / cell)), int(math.ceil(H / cell))
    grid = mask.resize((gw, gh), Image.BOX).point(lambda v: 255 if v else 0)
    grid = grid.filter(ImageFilter.MaxFilter(5))          # join nearby changes (e.g. words of a line)
    data = grid.tobytes()
    seen = bytearray(gw * gh)
    boxes = []
    for start in range(gw * gh):
        if not data[start] or seen[start]:
            continue
        seen[start] = 1
        stack = [start]
        x0 = x1 = start % gw
        y0 = y1 = start // gw
        while stack:
            i = stack.pop()
            x, y = i % gw, i // gw
            x0, x1, y0, y1 = min(x0, x), max(x1, x), min(y0, y), max(y1, y)
            for j in ((i - 1) if x > 0 else -1, (i + 1) if x < gw - 1 else -1,
                      (i - gw) if y > 0 else -1, (i + gw) if y < gh - 1 else -1):
                if j >= 0 and data[j] and not seen[j]:
                    seen[j] = 1
                    stack.append(j)
        box = (x0 * cell, y0 * cell, min(W, (x1 + 1) * cell), min(H, (y1 + 1) * cell))
        tight = mask.crop(box).getbbox()
        if tight is None:
            continue
        tight = (box[0] + tight[0], box[1] + tight[1], box[0] + tight[2], box[1] + tight[3])
        if mask.crop(tight).histogram()[255] >= COMPARE_MIN_PIXELS:
            boxes.append(tight)

    # merge boxes that touch or overlap (after padding)
    gap = cell
    merged = True
    while merged:
        merged = False
        out = []
        for b in boxes:
            for k, o in enumerate(out):
                if (b[0] <= o[2] + gap and o[0] <= b[2] + gap and
                        b[1] <= o[3] + gap and o[1] <= b[3] + gap):
                    out[k] = (min(b[0], o[0]), min(b[1], o[1]), max(b[2], o[2]), max(b[3], o[3]))
                    merged = True
                    break
            else:
                out.append(b)
        boxes = out
    # reading order: top to bottom, then left to right (rows of ~cell*4 px)
    boxes.sort(key=lambda b: (b[1] // (cell * 4), b[0]))
    return boxes


def _count(mask):
    return mask.histogram()[255]


def _paint(mask, part, xy):
    """mask = mask OR part, with part placed at xy."""
    x, y = xy
    box = (x, y, x + part.width, y + part.height)
    mask.paste(ImageChops.lighter(mask.crop(box), part), box[:2])


def _row_flags(mask):
    """For each row of an "L" mask: True if any pixel in it is set."""
    w, h = mask.size
    m = mask
    while w > 1:                       # fold the columns in half until one is left
        half = (w + 1) // 2
        m = ImageChops.lighter(m.crop((0, 0, half, h)), m.crop((w - half, 0, w, h)))
        w = half
    return [v > 0 for v in m.tobytes()]


def _content_pixels(gray, tol):
    """255 on everything that isn't the background (the most common shade)."""
    hist = gray.histogram()
    bg = max(range(256), key=lambda v: hist[v])
    return gray.point(lambda v: 255 if abs(v - bg) > tol else 0)


def find_bands(gray, tol=12):
    """Horizontal bands of content - lines of text, rows of buttons -
    separated by rows that just repeat the row above (background, or the
    plain middle of a panel). [(y0, y1), ...]"""
    W, H = gray.size
    if H < 3 or W < 1:
        return [(0, H)]
    step = ImageChops.difference(gray.crop((0, 1, W, H)), gray.crop((0, 0, W, H - 1)))
    busy = [False] + _row_flags(step.point(lambda v: 255 if v > tol else 0))
    bands = []
    y = 0
    while y < H:
        if not busy[y]:
            y += 1
            continue
        s = 0 if y == 1 else y
        while y < H and busy[y]:
            y += 1
        if bands and s - bands[-1][1] <= 2:      # thick stroke inside a line, not a gap
            bands[-1][1] = y
        else:
            bands.append([s, y])
    return [tuple(b) for b in bands]


def _band_matches(ba, bb, sa, sb):
    """Match the bands of Image 1 (ba) and Image 2 (bb) in order, like a
    text compare: first the lines that look the same (longest common
    sequence), then, between those, lines that changed. [(i, j), ...]"""
    na, nb = len(ba), len(bb)
    means_a = [ImageStat.Stat(s).mean[0] for s in sa]
    means_b = [ImageStat.Stat(s).mean[0] for s in sb]
    cache = {}

    def diff(i, j):
        k = (i, j)
        if k not in cache:
            cache[k] = ImageStat.Stat(ImageChops.difference(sa[i], sb[j])).mean[0]
        return cache[k]

    def height(bands, i):
        return bands[i][1] - bands[i][0]

    def same(i, j):
        return (abs(height(ba, i) - height(bb, j)) <= 1 and abs(means_a[i] - means_b[j]) <= 4
                and diff(i, j) <= 4)

    # 1) lines that look the same: longest common sequence
    L = [[0] * (nb + 1) for _ in range(na + 1)]
    S = [[False] * nb for _ in range(na)]
    for i in range(na - 1, -1, -1):
        Li, Ln, Si = L[i], L[i + 1], S[i]
        for j in range(nb - 1, -1, -1):
            if same(i, j):
                Si[j] = True
                Li[j] = Ln[j + 1] + 1
            else:
                Li[j] = Ln[j] if Ln[j] >= Li[j + 1] else Li[j + 1]
    anchors = []
    i = j = 0
    while i < na and j < nb:
        if S[i][j] and L[i][j] == L[i + 1][j + 1] + 1:
            anchors.append((i, j))
            i += 1
            j += 1
        elif L[i + 1][j] >= L[i][j + 1]:
            i += 1
        else:
            j += 1

    # 2) between two matched lines: pair up lines of a similar height (changed lines)
    def pair_gap(ai, bj):
        n, m = len(ai), len(bj)

        def score(p, q):
            ha, hb = height(ba, ai[p]), height(bb, bj[q])
            if abs(ha - hb) > max(3, 0.5 * max(ha, hb)):
                return 0
            return 2 - min(diff(ai[p], bj[q]), 100) / 100
        D = [[0.0] * (m + 1) for _ in range(n + 1)]
        for p in range(n - 1, -1, -1):
            for q in range(m - 1, -1, -1):
                s = score(p, q)
                D[p][q] = max(D[p + 1][q], D[p][q + 1], (D[p + 1][q + 1] + s) if s else 0)
        out, p, q = [], 0, 0
        while p < n and q < m:
            s = score(p, q)
            if s and abs(D[p][q] - (D[p + 1][q + 1] + s)) < 1e-9:
                out.append((ai[p], bj[q]))
                p += 1
                q += 1
            elif D[p + 1][q] >= D[p][q + 1]:
                p += 1
            else:
                q += 1
        return out

    pairs = []
    for (i0, j0), (i1, j1) in zip([(-1, -1)] + anchors, anchors + [(na, nb)]):
        ai, bj = list(range(i0 + 1, i1)), list(range(j0 + 1, j1))
        if ai and bj:
            pairs += pair_gap(ai, bj)
        if i1 < na:
            pairs.append((i1, j1))
    return sorted(pairs)


def _mapper(spans, h_from, h_to):
    """Map a y in one image to the matching y in the other, using matched
    row spans [(from_y0, from_y1, to_y0, to_y1), ...]."""
    def f(y):
        if not spans:
            return max(0, min(h_to, round(y * h_to / max(1, h_from))))
        prev = None
        for s0, s1, d0, d1 in spans:
            if y < s0:
                if prev is None:
                    return max(0, min(d0, y + d0 - s0))
                return max(prev[3], min(d0, y + prev[3] - prev[1]))
            if y < s1:
                return min(d1, d0 + (y - s0))
            prev = (s0, s1, d0, d1)
        return max(prev[3], min(h_to, y + prev[3] - prev[1]))
    return f


class _Match:
    """How Image 2 lines up with Image 1, and which pixels differ."""

    def __init__(self, a, b):
        self.mask1 = Image.new("L", a.size, 0)     # differences, in Image 1
        self.mask2 = Image.new("L", b.size, 0)     # differences, in Image 2
        self.missing1 = Image.new("L", a.size, 0)  # content not found in Image 2
        self.added2 = Image.new("L", b.size, 0)    # content not found in Image 1
        self.dx = 0
        self.map12 = self.map21 = None
        self.moved = False
        self.compared = a.width * a.height

    @property
    def area(self):
        return _count(self.mask1) + _count(self.mask2)


def _compare_pixels(a, b, tolerance, dx, dy, resized):
    """Straight pixel compare: (x, y) of Image 1 <-> (x + dx, y + dy) of Image 2."""
    m = _Match(a, b)
    m.dx = dx
    Wa, Ha = a.size
    Wb, Hb = b.size
    ov = (max(0, -dx), max(0, -dy), min(Wa, Wb - dx), min(Ha, Hb - dy))
    if ov[2] > ov[0] and ov[3] > ov[1]:
        part = _changed_mask(a.crop(ov), b.crop((ov[0] + dx, ov[1] + dy, ov[2] + dx, ov[3] + dy)),
                             tolerance)
        if resized and part.width > 8 and part.height > 8:
            # resizing blurs the outermost pixels - don't report them
            inner = Image.new("L", part.size, 0)
            inner.paste(255, (2, 2, part.width - 2, part.height - 2))
            part = ImageChops.darker(part, inner)
        _paint(m.mask1, part, ov[:2])
        _paint(m.mask2, part, (ov[0] + dx, ov[1] + dy))
    if a.size != b.size:
        # content outside the shared area exists in only one of the images
        tol = max(10, tolerance // 2)
        out1 = Image.new("L", a.size, 255)
        out2 = Image.new("L", b.size, 255)
        if ov[2] > ov[0] and ov[3] > ov[1]:
            out1.paste(0, ov)
            out2.paste(0, (ov[0] + dx, ov[1] + dy, ov[2] + dx, ov[3] + dy))
        m.missing1 = ImageChops.darker(out1, _content_pixels(a.convert("L"), tol))
        m.added2 = ImageChops.darker(out2, _content_pixels(b.convert("L"), tol))
        m.mask1 = ImageChops.lighter(m.mask1, m.missing1)
        m.mask2 = ImageChops.lighter(m.mask2, m.added2)
    m.map12 = lambda y: max(0, min(Hb, y + dy))
    m.map21 = lambda y: max(0, min(Ha, y - dy))
    return m


def _compare_lines(a, b, tolerance, dx):
    """Line-by-line compare (see the notes at the top of this section).
    None if the images don't consist of lines."""
    Wa, Ha = a.size
    Wb, Hb = b.size
    xa0, xa1 = max(0, -dx), min(Wa, Wb - dx)
    if xa1 - xa0 < 0.5 * min(Wa, Wb):
        return None
    ga, gb = a.convert("L"), b.convert("L")
    ca, cb = ga.crop((xa0, 0, xa1, Ha)), gb.crop((xa0 + dx, 0, xa1 + dx, Hb))
    tol = max(10, tolerance // 2)
    ba, bb = find_bands(ca, tol), find_bands(cb, tol)
    if len(ba) < 2 or len(bb) < 2 or len(ba) * len(bb) > 600 * 600:
        return None
    sa = [ca.crop((0, s, ca.width, e)).resize(COMPARE_SIG, Image.BOX) for s, e in ba]
    sb = [cb.crop((0, s, cb.width, e)).resize(COMPARE_SIG, Image.BOX) for s, e in bb]
    pairs = _band_matches(ba, bb, sa, sb)
    if not pairs:
        return None

    m = _Match(a, b)
    m.dx = dx
    spans = []

    def compare_rows(ya, yb, h):
        part = _changed_mask(a.crop((xa0, ya, xa1, ya + h)), b.crop((xa0 + dx, yb, xa1 + dx, yb + h)),
                             tolerance)
        _paint(m.mask1, part, (xa0, ya))
        _paint(m.mask2, part, (xa0 + dx, yb))

    for i, j in pairs:
        (a0, a1), (b0, b1) = ba[i], bb[j]
        h = max(a1 - a0, b1 - b0)
        best = None
        for d in (0, -1, 1, -2, 2):                  # line up the rows exactly
            y2 = b0 + d
            hh = min(h, Ha - a0, Hb - y2)
            if y2 < 0 or hh <= 0:
                continue
            s = ImageStat.Stat(ImageChops.difference(ca.crop((0, a0, ca.width, a0 + hh)),
                                                     cb.crop((0, y2, cb.width, y2 + hh)))).mean[0]
            if best is None or s < best[0] - 0.01:
                best = (s, y2, hh)
        if best is None:
            continue
        _, y2, hh = best
        compare_rows(a0, y2, hh)
        spans.append((a0, a0 + hh, y2, y2 + hh))
        if y2 != a0:
            m.moved = True

    # the plain rows between two matched lines (same height in both images)
    edges = [(0, 0, 0, 0)] + spans + [(Ha, Ha, Hb, Hb)]
    for p, q in zip(edges, edges[1:]):
        ha, hb = q[0] - p[1], q[2] - p[3]
        if ha > 0 and ha == hb:
            compare_rows(p[1], p[3], ha)

    # lines found in only one of the images
    paired_a = {i for i, _ in pairs}
    paired_b = {j for _, j in pairs}
    for i, (s, e) in enumerate(ba):
        if i not in paired_a:
            part = _content_pixels(ca.crop((0, s, ca.width, e)), tol)
            _paint(m.missing1, part, (xa0, s))
    for j, (s, e) in enumerate(bb):
        if j not in paired_b:
            part = _content_pixels(cb.crop((0, s, cb.width, e)), tol)
            _paint(m.added2, part, (xa0 + dx, s))
    m.mask1 = ImageChops.lighter(m.mask1, m.missing1)
    m.mask2 = ImageChops.lighter(m.mask2, m.added2)

    m.map12 = _mapper(spans, Ha, Hb)
    m.map21 = _mapper([(s2, e2, s1, e1) for s1, e1, s2, e2 in spans], Hb, Ha)
    if any(i not in paired_a for i in range(len(ba))) or any(j not in paired_b for j in range(len(bb))):
        m.moved = True
    return m


def _plain(img, box):
    if box[2] - box[0] < 1 or box[3] - box[1] < 1:
        return True
    return ImageStat.Stat(img.crop(box).convert("L")).stddev[0] < 4


def _differences(a, b, m):
    """Turn the masks into numbered differences, each with a box in both images."""
    Wa, Ha = a.size
    Wb, Hb = b.size
    dx = m.dx
    diffs = []
    claimed = Image.new("L", b.size, 0)          # parts of Image 2 already explained
    for box in find_regions(m.mask1):
        x0, y0, x1, y1 = box
        total = _count(m.mask1.crop(box))
        missing = _count(m.missing1.crop(box))
        bx0, bx1 = max(0, min(Wb, x0 + dx)), max(0, min(Wb, x1 + dx))
        by0 = m.map12(y0)
        by1 = m.map12(y1 - 1) + 1
        if (total and missing >= total * 0.5) or bx1 <= bx0 or by0 >= Hb:
            y = max(0, min(by0, Hb) - 3)              # in the gap just above the next line
            diffs.append(Difference(KIND_MISSING, box, (bx0, y, max(bx1, bx0 + 1), y), marker2=True))
            continue
        by1 = max(by0 + 1, min(Hb, by1))
        # include the differing pixels of Image 2 right around it (e.g. a longer word)
        pad = 4
        near = (max(0, bx0 - pad), max(0, by0 - pad), min(Wb, bx1 + pad), min(Hb, by1 + pad))
        tight = m.mask2.crop(near).getbbox()
        if tight:
            bx0, by0 = min(bx0, near[0] + tight[0]), min(by0, near[1] + tight[1])
            bx1, by1 = max(bx1, near[0] + tight[2]), max(by1, near[1] + tight[3])
        box2 = (bx0, by0, bx1, by1)
        claimed.paste(255, (max(0, bx0 - pad), max(0, by0 - pad), min(Wb, bx1 + pad), min(Hb, by1 + pad)))
        p1, p2 = _plain(a, box), _plain(b, box2)
        kind = KIND_MISSING if (p2 and not p1) else KIND_ADDED if (p1 and not p2) else KIND_CHANGED
        diffs.append(Difference(kind, box, box2))
    rest = ImageChops.darker(ImageChops.lighter(m.added2, m.mask2), ImageChops.invert(claimed))
    for box in find_regions(rest):
        x0, y0, x1, y1 = box
        ax0, ax1 = max(0, min(Wa, x0 - dx)), max(0, min(Wa, x1 - dx))
        y = max(0, min(m.map21(y0), Ha) - 3)
        diffs.append(Difference(KIND_ADDED, (ax0, y, max(ax1, ax0 + 1), y), box, marker1=True))
    diffs.sort(key=lambda d: (d.box1[1], d.box1[0], d.box2[1]))
    return diffs


def compare_images(a, b, tolerance=30, align=True):
    """Compare Image 2 (b) against Image 1 (a, the source). tolerance: 0-255
    colour change that is still treated as 'the same'."""
    a, b = a.convert("RGB"), b.convert("RGB")
    notes = []
    resized = False
    if a.size != b.size:
        ra, rb = a.width / a.height, b.width / b.height
        if abs(ra - rb) / ra < 0.01 and 0.4 < b.width / a.width < 2.5:
            b = b.resize(a.size, Image.LANCZOS)
            resized = True
            notes.append("Image 2 was resized to match Image 1")
    dx, dy = find_alignment(a, b) if align else (0, 0)

    best = _compare_pixels(a, b, tolerance, dx, dy, resized)
    method = "pixels"
    if not (a.size == b.size and best.area < 0.004 * a.width * a.height):
        lines = _compare_lines(a, b, tolerance, dx)
        if lines is not None and lines.area < best.area * 0.9:
            best, method = lines, "lines"
    if method == "pixels" and (dx, dy) != (0, 0):
        notes.append(f"aligned by {dx:+d}, {dy:+d} px")
    if method == "lines" and best.moved:
        notes.append("lines matched up (content moved up or down)")
    if a.size != b.size:
        notes.append(f"sizes differ ({a.width}×{a.height} vs {b.width}×{b.height})")

    diffs = _differences(a, b, best)
    return CompareResult(image1=a, image2=b, mask1=best.mask1, mask2=best.mask2, diffs=diffs,
                         changed_pixels=_count(best.mask1), compared_pixels=max(1, best.compared),
                         notes=notes, method=method)


COMPARE_GUTTER_BG = (238, 240, 245)    # the margin strip beside each image


def gutter_width(*images):
    """Width of the margin strip where the change bars and numbers go."""
    return max(48, round(max(im.width for im in images) * 0.09))


def side_by_side(image1, image2):
    """[margin | Image 1]  [margin | Image 2] with a label above each. The
    images are shown exactly as they are - nothing is drawn over them.
    Returns the picture, where each image is, and where each margin is."""
    g = gutter_width(image1, image2)
    gap, bar = 24, 34
    x2 = g + image1.width + gap
    W, H = x2 + g + image2.width, max(image1.height, image2.height) + bar
    out = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(out)
    d.rectangle([0, bar, g - 1, H], fill=COMPARE_GUTTER_BG)
    d.rectangle([x2, bar, x2 + g - 1, H], fill=COMPARE_GUTTER_BG)
    out.paste(image1, (g, bar))
    out.paste(image2, (x2 + g, bar))
    d.rectangle([g + image1.width, 0, x2 - 1, H], fill=(201, 206, 216))
    font = text_font(16)
    d.text((4, 8), "IMAGE 1  (SOURCE)", fill=(27, 31, 42), font=font)
    d.text((x2 + 4, 8), "IMAGE 2", fill=(27, 31, 42), font=font)
    return out, [(g, bar), (x2 + g, bar)], [(0, g), (x2, x2 + g)]


def with_margin(image):
    """One image with the margin strip on its left."""
    g = gutter_width(image)
    out = Image.new("RGB", (g + image.width, image.height), COMPARE_GUTTER_BG)
    out.paste(image, (g, 0))
    return out, (g, 0), (0, g)


def _dashed_line(d, x0, x1, y, color, width, dash=10):
    x = x0
    while x < x1:
        d.line([x, y, min(x1, x + dash), y], fill=color, width=width)
        x += dash * 2


def draw_difference_boxes(image, diffs, offsets, gutters=None, selected=None):
    """The marks drawn into the picture itself (for Copy / Save): an outline
    around each difference (with a gap, so it never covers the content), a
    bar in the margin over the same rows, and the number in the margin.
    offsets / gutters: [Image 1, Image 2], None where that image isn't shown."""
    out = image.copy()
    W, H = out.size
    d = ImageDraw.Draw(out)
    lw = max(2, round(max(W, H) / 700))
    pad = lw + 5
    font = text_font(max(13, lw * 6))
    gutters = gutters or [None, None]
    next_free = [0, 0]                           # stack numbers that would overlap
    for n, diff in enumerate(diffs, 1):
        for k, off in enumerate(offsets):
            if off is None:
                continue
            x0, y0, x1, y1 = diff.box1 if k == 0 else diff.box2
            marker = diff.marker1 if k == 0 else diff.marker2
            ox, oy = off
            width = lw * 2 if n == selected else lw
            if marker:
                x1 = max(x1, x0 + 40)
                _dashed_line(d, x0 + ox, x1 + ox, y0 + oy, "white", width + 2)
                _dashed_line(d, x0 + ox, x1 + ox, y0 + oy, COMPARE_RED, width)
                by0, by1 = y0 + oy - pad, y0 + oy + pad
            else:
                bx0, by0, bx1, by1 = x0 + ox - pad, y0 + oy - pad, x1 + ox + pad, y1 + oy + pad
                d.rectangle([bx0 - 1, by0 - 1, bx1 + 1, by1 + 1], outline="white", width=width + 2)
                d.rectangle([bx0, by0, bx1, by1], outline=COMPARE_RED, width=width)
            label = str(n)
            tb = d.textbbox((0, 0), label, font=font)
            tw, th = tb[2] - tb[0], tb[3] - tb[1]
            gut = gutters[k] if k < len(gutters) else None
            if gut:
                gx0, gx1 = gut
                bw = max(4, round((gx1 - gx0) * 0.12))
                d.rectangle([gx1 - bw - 3, max(by0, oy), gx1 - 4, max(by1, by0 + 6)], fill=COMPARE_RED)
                lx = gx0 + 3
                ty = max(by0, next_free[k], oy)
                next_free[k] = ty + th + 12
            else:
                lx = (x0 + ox - pad)
                ty = by0 - th - 10 if by0 - th - 10 >= 0 else by1 + 2
            d.rectangle([lx, ty, lx + tw + 10, ty + th + 10], fill=COMPARE_RED)
            d.text((lx + 5 - tb[0], ty + 5 - tb[1]), label, fill="white", font=font)
    return out


def zoom_rects(result, diff, ratio=None):
    """Close-up areas around one difference: (rect in Image 1, rect in Image 2),
    the same size so the two can be shown next to each other."""
    boxes = [diff.box1, diff.box2]
    sizes = [(max(1, b[2] - b[0]), max(1, b[3] - b[1])) for b in boxes]
    w = max(s[0] for s in sizes)
    h = max(s[1] for s in sizes)
    ctx = max(30, int(max(w, h) * 0.35))
    rw, rh = w + 2 * ctx, h + 2 * ctx
    if ratio:                                   # widen / heighten to the display's shape
        if rw / rh < ratio:
            rw = int(rh * ratio)
        else:
            rh = int(rw / ratio)
    rects = []
    for box, img in zip(boxes, (result.image1, result.image2)):
        W, H = img.size
        cw, ch = min(rw, W), min(rh, H)
        cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
        x0 = int(max(0, min(W - cw, cx - cw / 2)))
        y0 = int(max(0, min(H - ch, cy - ch / 2)))
        rects.append((x0, y0, x0 + cw, y0 + ch))
    return rects


def difference_thumbnail(result, diff, size, images=None):
    """Zoomed close-up of one difference: Image 1 | arrow | Image 2."""
    max_w, max_h = size
    imgs = images or (result.image1, result.image2)
    rects = zoom_rects(result, diff, ratio=max_w / max_h)
    arrow_w = 26
    out = Image.new("RGB", (max_w * 2 + arrow_w, max_h), (255, 255, 255))
    rw = max(r[2] - r[0] for r in rects)
    rh = max(r[3] - r[1] for r in rects)
    scale = min(4.0, max_w / rw, max_h / rh)
    method = Image.NEAREST if scale >= 2 else (Image.BICUBIC if scale >= 1 else Image.LANCZOS)
    for k, (img, r) in enumerate(zip(imgs, rects)):
        crop = img.crop(r)
        tw, th = max(1, int(crop.width * scale)), max(1, int(crop.height * scale))
        crop = crop.resize((tw, th), method)
        d = ImageDraw.Draw(crop)
        box = diff.box1 if k == 0 else diff.box2
        marker = diff.marker1 if k == 0 else diff.marker2
        bx0, by0 = (box[0] - r[0]) * scale, (box[1] - r[1]) * scale
        bx1, by1 = (box[2] - r[0]) * scale, (box[3] - r[1]) * scale
        if marker:
            _dashed_line(d, bx0, max(bx1, bx0 + 20), by0, COMPARE_RED, 2, dash=5)
        else:
            d.rectangle([bx0 - 2, by0 - 2, bx1 + 1, by1 + 1], outline=COMPARE_RED, width=2)
        px = k * (max_w + arrow_w) + (max_w - tw) // 2
        py = (max_h - th) // 2
        out.paste(crop, (px, py))
        ImageDraw.Draw(out).rectangle([px - 1, py - 1, px + tw, py + th], outline=(201, 206, 216))
    d = ImageDraw.Draw(out)
    ax, ay = max_w + 5, max_h // 2
    d.line([ax, ay, ax + 14, ay], fill=(75, 85, 99), width=3)
    d.polygon([(ax + 12, ay - 6), (ax + 19, ay), (ax + 12, ay + 6)], fill=(75, 85, 99))
    return out


class CompareWindow:
    """Load or capture two images, compare them, and show the differences."""

    current = None
    SLOT_TITLES = ("IMAGE 1  ·  SOURCE", "IMAGE 2  ·  CHECKED AGAINST IMAGE 1")
    VIEWS = (("side", "Side by side"), ("image1", "Image 1"), ("image2", "Image 2"))
    MAX_LISTED = 150

    def __init__(self):
        if CompareWindow.current is not None and CompareWindow.current.win.winfo_exists():
            CompareWindow.current.show()
            return
        CompareWindow.current = self
        self.images = [None, None]
        self.names = ["", ""]
        self.result = None
        self.bases = {}           # view -> (image, [offset of Image 1, offset of Image 2])
        self.hover = None         # number of the difference under the mouse
        self.drawn = ([None, None], None, [None, None])   # offsets, clips, margins on screen now
        self.selected = None      # number of the selected difference (1-based)
        self.zoom = None          # number of the difference zoomed in on
        self.photo = None
        self.thumbs = [None, None]
        self.item_photos = []
        self.items = []
        self.busy = False
        self._resize_job = None
        self.transform = None     # (scale, ox, oy, vx0, vy0) of the current drawing

        self.win = tk.Toplevel(root)
        self.win.title("Compare images")
        self.win.configure(bg=BG)
        self.win.protocol("WM_DELETE_WINDOW", self.close)

        # --- header ---------------------------------------------------
        head = tk.Frame(self.win, bg=HEADER_BG)
        head.pack(fill="x")
        inner = tk.Frame(head, bg=HEADER_BG)
        inner.pack(fill="x", padx=S(20), pady=S(10))
        tk.Label(inner, image=icon("compare", "white", 22), bg=HEADER_BG).pack(side="left")
        titles = tk.Frame(inner, bg=HEADER_BG)
        titles.pack(side="left", padx=(S(10), 0))
        tk.Label(titles, text="Compare images", font=F_TILE, fg="white", bg=HEADER_BG).pack(anchor="w")
        tk.Label(titles, text="Image 1 is the source. Differences are outlined in red in both images "
                              "(- - - = missing). Point at one to see its match; click to zoom in.",
                 font=F_TILE_SUB, fg=HEADER_SUB, bg=HEADER_BG).pack(anchor="w")
        tk.Frame(self.win, bg=ACCENT, height=S(3)).pack(fill="x")

        body = tk.Frame(self.win, bg=BG)
        body.pack(fill="both", expand=True, padx=S(16), pady=S(12))

        # --- the two image slots -----------------------------------------
        slots = tk.Frame(body, bg=BG)
        slots.pack(fill="x")
        slots.columnconfigure((0, 1), weight=1, uniform="slots")
        self.slot = [self._make_slot(slots, i) for i in range(2)]

        # --- options + Compare --------------------------------------------
        opts = tk.Frame(body, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        opts.pack(fill="x", pady=(S(8), 0))
        row = tk.Frame(opts, bg=CARD)
        row.pack(fill="x", padx=S(12), pady=S(7))
        tk.Label(row, text="Tolerance", font=F_BODY_BOLD, fg=TEXT, bg=CARD).pack(side="left")
        self.tol_var = tk.DoubleVar(value=30)
        ttk.Scale(row, from_=0, to=120, orient="horizontal", variable=self.tol_var,
                  length=S(150)).pack(side="left", padx=(S(8), S(6)))
        self.tol_label = tk.Label(row, font=F_BODY, fg=MUTED, bg=CARD, width=15, anchor="w")
        self.tol_label.pack(side="left")
        self.tol_var.trace_add("write", lambda *a: self._tol_text())
        self._tol_text()
        self.align_var = tk.BooleanVar(value=True)
        Toggle(row, "Auto-align", self.align_var).pack(side="left", padx=(S(6), 0))
        self.compare_button = FlatButton(row, "Compare", self.compare, style="primary")
        self.compare_button.pack(side="right")
        self.swap_button = FlatButton(row, "Swap images", self.swap)
        self.swap_button.pack(side="right", padx=(0, S(8)))

        # --- result header: summary + views ---------------------------------
        res_head = tk.Frame(body, bg=BG)
        res_head.pack(fill="x", pady=(S(8), S(6)))
        self.summary = tk.Label(res_head, text="", font=F_BODY_BOLD, fg=TEXT, bg=BG, anchor="w",
                                justify="left")
        self.summary.pack(side="left", fill="x", expand=True)
        self.legend = tk.Label(res_head, text="", font=F_TILE_SUB, fg=MUTED, bg=BG)
        self.legend.pack(side="left", padx=(0, S(10)))
        self.view_var = tk.StringVar(value="side")
        self.view_buttons = []
        for value, label in reversed(self.VIEWS):
            b = tk.Radiobutton(res_head, text=label, value=value, variable=self.view_var,
                               indicatoron=False, command=self._view_changed, font=F_BODY_BOLD,
                               bg=CARD, fg=TEXT, selectcolor=ACCENT_SOFT, activebackground="#EEF0F4",
                               relief="flat", bd=0, highlightthickness=1, highlightbackground=FIELD_BORDER,
                               padx=S(12), pady=S(4), cursor="hand2")
            b.pack(side="right", padx=(S(4), 0))
            self.view_buttons.append(b)

        # bottom buttons (packed before the result so they stay visible)
        bottom = tk.Frame(body, bg=BG)
        bottom.pack(side="bottom", fill="x", pady=(S(10), 0))
        self.copy_button = FlatButton(bottom, "Copy result", self.copy_result,
                                      icons={"Copy result": icon("copy", TEXT, 15),
                                             "Copied": icon("check", GREEN, 15)})
        self.copy_button.pack(side="left")
        self.save_button = FlatButton(bottom, "Save result...", self.save_result)
        self.save_button.pack(side="left", padx=(S(8), 0))
        FlatButton(bottom, "Close", self.close).pack(side="right")

        # --- result: picture (left) + list of differences (right) ----------
        result_row = tk.Frame(body, bg=BG)
        result_row.pack(fill="both", expand=True)

        side = tk.Frame(result_row, bg=CARD, highlightthickness=1, highlightbackground=BORDER,
                        width=S(330))
        side.pack(side="right", fill="y", padx=(S(10), 0))
        side.pack_propagate(False)
        side_head = tk.Frame(side, bg=CARD)
        side_head.pack(fill="x", padx=S(10), pady=(S(8), S(4)))
        tk.Label(side_head, text="DIFFERENCES", font=F_SECTION, fg=TEXT, bg=CARD).pack(side="left")
        self.count_label = tk.Label(side_head, text="", font=F_SECTION, fg=COMPARE_HEX, bg=CARD)
        self.count_label.pack(side="left", padx=(S(6), 0))
        nav = tk.Frame(side, bg=CARD)
        nav.pack(fill="x", padx=S(10), pady=(0, S(6)))
        self.prev_button = FlatButton(nav, "Previous", lambda: self.step(-1))
        self.prev_button.pack(side="left")
        self.next_button = FlatButton(nav, "Next", lambda: self.step(1))
        self.next_button.pack(side="left", padx=(S(6), 0))
        self.all_button = FlatButton(nav, "Show all", self.show_all)
        self.all_button.pack(side="right")
        tk.Frame(side, bg="#E3E6EC", height=1).pack(fill="x")
        list_wrap = tk.Frame(side, bg=CARD)
        list_wrap.pack(fill="both", expand=True)
        self.list_canvas = tk.Canvas(list_wrap, bg=CARD, highlightthickness=0)
        list_scroll = tk.Scrollbar(list_wrap, orient="vertical", command=self.list_canvas.yview)
        self.list_canvas.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side="right", fill="y")
        self.list_canvas.pack(side="left", fill="both", expand=True)
        self.list_frame = tk.Frame(self.list_canvas, bg=CARD)
        self.list_window = self.list_canvas.create_window(0, 0, anchor="nw", window=self.list_frame)
        self.list_frame.bind("<Configure>", lambda e: self.list_canvas.configure(
            scrollregion=self.list_canvas.bbox("all")))
        self.list_canvas.bind("<Configure>", lambda e: self.list_canvas.itemconfigure(
            self.list_window, width=e.width))
        for w in (self.list_canvas, self.list_frame):
            w.bind("<MouseWheel>", self._list_wheel)

        frame = tk.Frame(result_row, bg=CARD, highlightthickness=1, highlightbackground=FIELD_BORDER)
        frame.pack(side="left", fill="both", expand=True)
        self.canvas = tk.Canvas(frame, bg="#2A2F3A", highlightthickness=0, height=S(300))
        self.scroll = tk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self._schedule_redraw())
        self.canvas.bind("<MouseWheel>", self._canvas_wheel)
        self.canvas.bind("<Button-1>", self._canvas_click)
        self.canvas.bind("<Motion>", self._canvas_motion)
        self.canvas.bind("<Leave>", lambda e: self._canvas_motion(type("E", (), {"x": -999, "y": -999})()))
        self.canvas.bind("<Double-Button-1>", lambda e: self.show_all())

        self.win.bind("<Return>", lambda e: self.compare())
        self.win.bind("<Escape>", lambda e: self.show_all() if self.zoom else self.close())
        self.win.bind("<Left>", lambda e: self.step(-1))
        self.win.bind("<Right>", lambda e: self.step(1))
        self.win.bind("<Control-c>", lambda e: self.copy_result())
        self.win.bind("<Control-C>", lambda e: self.copy_result())
        mirror_control_keys(self.win)
        self.win.bind("<Down>", lambda e: self.step(1))
        self.win.bind("<Up>", lambda e: self.step(-1))

        left, top, width, height = screen_work_area()
        w, h = min(S(1280), int(width * 0.94)), min(S(940), int(height * 0.92))
        self.win.geometry(f"{w}x{h}+{left + (width - w) // 2}+{top + max(0, (height - h) // 2 - S(10))}")
        self.win.minsize(min(w, S(900)), min(h, S(600)))
        self._update_state()
        self._fill_list()
        self.show_result()
        self.show()

    # ---- slots ---------------------------------------------------------

    def _make_slot(self, parent, i):
        card = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.grid(row=0, column=i, sticky="nsew", padx=(0, S(5)) if i == 0 else (S(5), 0))
        top = tk.Frame(card, bg=CARD)
        top.pack(fill="x", padx=S(12), pady=(S(8), S(4)))
        tk.Frame(top, bg=ACCENT, width=S(4), height=S(14)).pack(side="left", padx=(0, S(8)))
        tk.Label(top, text=self.SLOT_TITLES[i], font=F_SECTION, fg=TEXT, bg=CARD).pack(side="left")
        info = tk.Label(top, text="No image yet", font=F_TILE_SUB, fg=MUTED, bg=CARD)
        info.pack(side="right")
        preview = tk.Canvas(card, bg="#2A2F3A", highlightthickness=0, height=S(110))
        preview.pack(fill="x", padx=S(12))
        preview.bind("<Configure>", lambda e, i=i: self._draw_thumb(i))
        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(fill="x", padx=S(12), pady=S(8))
        FlatButton(buttons, "Open...", lambda: self.open_file(i)).pack(side="left")
        FlatButton(buttons, "Capture region", lambda: self.capture_region(i)).pack(side="left", padx=(S(6), 0))
        FlatButton(buttons, "Full screen", lambda: self.capture_screen(i)).pack(side="left", padx=(S(6), 0))
        FlatButton(buttons, "Paste", lambda: self.paste(i)).pack(side="left", padx=(S(6), 0))
        return {"info": info, "preview": preview}

    def _draw_thumb(self, i):
        c = self.slot[i]["preview"]
        c.delete("all")
        cw, ch = max(10, c.winfo_width()), max(10, c.winfo_height())
        img = self.images[i]
        if img is None:
            c.create_text(cw // 2, ch // 2, fill="#9AA3B2", font=F_BODY,
                          text="Open, capture or paste an image")
            return
        thumb = img.copy()
        thumb.thumbnail((cw - 8, ch - 8), Image.LANCZOS)
        self.thumbs[i] = ImageTk.PhotoImage(thumb)
        c.create_image(cw // 2, ch // 2, image=self.thumbs[i])

    def set_image(self, i, image, name):
        self.images[i] = image.convert("RGB")
        self.names[i] = name
        w, h = self.images[i].size
        self.slot[i]["info"].configure(text=f"{name}  ·  {w} × {h} px", fg=TEXT)
        self._draw_thumb(i)
        self.clear_result()
        self._update_state()

    def _update_state(self):
        both = all(img is not None for img in self.images)
        self.compare_button.config(state="normal" if both and not self.busy else "disabled")
        self.swap_button.config(state="normal" if any(self.images) and not self.busy else "disabled")
        has = self.result is not None
        n = len(self.result.diffs) if has else 0
        self.copy_button.config(state="normal" if has else "disabled")
        self.save_button.config(state="normal" if has else "disabled")
        for b in self.view_buttons:
            b.configure(state="normal" if has else "disabled")
        self.prev_button.config(state="normal" if n else "disabled")
        self.next_button.config(state="normal" if n else "disabled")
        self.all_button.config(state="normal" if self.zoom else "disabled")

    def _tol_text(self):
        v = int(round(self.tol_var.get()))
        word = ("exact" if v == 0 else "strict" if v < 20 else "normal" if v < 60 else "relaxed")
        self.tol_label.configure(text=f"{v}  ({word})")

    # ---- loading images ------------------------------------------------

    def show(self):
        self.win.deiconify()
        self.win.lift()
        self.win.focus_force()

    def open_file(self, i):
        path = filedialog.askopenfilename(
            parent=self.win, title=f"Open Image {i + 1}",
            initialdir=str(save_folder) if save_folder else str(Path.home()),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"), ("All files", "*.*")])
        if not path:
            return
        try:
            with Image.open(path) as im:
                im.load()
                self.set_image(i, im, Path(path).name)
        except Exception as e:
            messagebox.showerror("Could not open image", f"{path}\n\n{e}", parent=self.win)

    def paste(self, i):
        try:
            data = ImageGrab.grabclipboard()
        except Exception as e:
            data = None
            print(f"Clipboard read failed: {e}")
        if isinstance(data, list):                     # files copied in Explorer
            for p in data:
                try:
                    with Image.open(p) as im:
                        im.load()
                        self.set_image(i, im, Path(p).name)
                        return
                except Exception:
                    continue
            data = None
        if data is None or not hasattr(data, "size"):
            messagebox.showinfo("Nothing to paste",
                                "There is no image on the clipboard.\n\n"
                                "Copy a screenshot or an image file first, then click Paste.",
                                parent=self.win)
            return
        self.set_image(i, data, "Pasted image")

    def capture_region(self, i):
        if recorder is not None:
            return
        self.win.withdraw()
        start_region_selection(lambda x1, y1, x2, y2: self._grab(i, (x1, y1, x2, y2)),
                               hint=f"Drag over the area for Image {i + 1}  ·  Esc to cancel",
                               on_cancel=self.show)

    def capture_screen(self, i):
        if recorder is not None:
            return
        self.win.withdraw()
        root.withdraw()
        root.after(500, lambda: self._grab(i, None))

    def _grab(self, i, bbox):
        try:
            image = grab_screen(bbox)
        except Exception as e:
            image = None
            messagebox.showerror("Capture failed", str(e), parent=root)
        root.deiconify()
        self.show()
        if image is not None:
            self.set_image(i, image, "Screen region" if bbox else "Full screen")

    def swap(self):
        self.images.reverse()
        self.names.reverse()
        for i in range(2):
            img = self.images[i]
            if img is None:
                self.slot[i]["info"].configure(text="No image yet", fg=MUTED)
            else:
                self.slot[i]["info"].configure(
                    text=f"{self.names[i]}  ·  {img.width} × {img.height} px", fg=TEXT)
            self._draw_thumb(i)
        self.clear_result()
        self._update_state()

    # ---- comparing -----------------------------------------------------

    def compare(self):
        if self.busy or not all(img is not None for img in self.images):
            return
        self.busy = True
        self.clear_result()
        self.compare_button.config(text="Comparing...")
        self.summary.configure(text="Comparing...", fg=MUTED)
        self._update_state()
        self.show_result()
        a, b = self.images
        tol, align = int(round(self.tol_var.get())), self.align_var.get()
        box = {}

        def work():
            try:
                res = compare_images(a, b, tolerance=tol, align=align)
                box["bases"] = self._make_bases(res)
                box["thumbs"] = [difference_thumbnail(res, d, (S(132), S(92)))
                                 for d in res.diffs[:self.MAX_LISTED]]
                box["result"] = res
            except Exception as e:
                box["error"] = e

        thread = threading.Thread(target=work, daemon=True)
        thread.start()

        def poll():
            if thread.is_alive():
                self.win.after(100, poll)
                return
            self.busy = False
            self.compare_button.config(text="Compare")
            if "error" in box:
                print(f"Compare failed: {box['error']}")
                self.summary.configure(text=f"Could not compare the images: {box['error']}", fg=RED)
            else:
                self.result, self.bases = box["result"], box["bases"]
                self._summarise()
                self._fill_list(box["thumbs"])
            self._update_state()
            self.show_result()
        poll()

    @staticmethod
    def _make_bases(res):
        """Each view: (picture, [where Image 1 is, where Image 2 is], [their margins])."""
        side, offsets, gutters = side_by_side(res.image1, res.image2)
        m1, o1, g1 = with_margin(res.image1)
        m2, o2, g2 = with_margin(res.image2)
        return {"side": (side, offsets, gutters),
                "image1": (m1, [o1, None], [g1, None]),
                "image2": (m2, [None, o2], [None, g2])}

    def _summarise(self):
        r = self.result
        n = len(r.diffs)
        extra = ("   ·   " + "; ".join(r.notes)) if r.notes else ""
        if n == 0:
            self.summary.configure(text="No differences - Image 2 matches Image 1." + extra, fg=GREEN)
            self.legend.configure(text="")
            status_var.set("Compare images: no differences found")
            return
        parts = []
        for kind, word in ((KIND_CHANGED, "changed"), (KIND_MISSING, "missing in Image 2"),
                           (KIND_ADDED, "added in Image 2")):
            k = r.count(kind)
            if k:
                parts.append(f"{k} {word}")
        self.summary.configure(
            text=f"{n} difference{'s' if n != 1 else ''}:  " + ",  ".join(parts) + extra, fg=COMPARE_HEX)
        self.legend.configure(text="")
        status_var.set(f"Compare images: {n} difference{'s' if n != 1 else ''} found")

    def clear_result(self):
        self.result = None
        self.bases = {}
        self.hover = None
        self.selected = None
        self.zoom = None
        self.summary.configure(text="", fg=TEXT)
        if hasattr(self, "legend"):
            self.legend.configure(text="")
        if hasattr(self, "list_frame"):
            self._fill_list()
        if hasattr(self, "all_button"):
            self._update_state()
        self.show_result()

    # ---- list of differences --------------------------------------------

    def _fill_list(self, thumbs=None):
        for w in self.list_frame.winfo_children():
            w.destroy()
        self.items, self.item_photos = [], []
        if self.result is None or not self.result.diffs:
            self.count_label.configure(text="")
            msg = ("Comparing..." if self.busy else
                   "No differences - the images match." if self.result is not None else
                   "The differences will be listed here - each one shown close up, Image 1 next to Image 2.")
            tk.Label(self.list_frame, text=msg, font=F_BODY, fg=MUTED, bg=CARD, wraplength=S(280),
                     justify="left").pack(anchor="w", padx=S(12), pady=S(12))
            return
        diffs = self.result.diffs
        shown = diffs[:self.MAX_LISTED]
        self.count_label.configure(text=str(len(diffs)))
        for n, diff in enumerate(shown, 1):
            kind = diff.kind
            item = tk.Frame(self.list_frame, bg=CARD, cursor="hand2", highlightthickness=1,
                            highlightbackground=CARD)
            item.pack(fill="x", padx=S(6), pady=(S(4), 0))
            head = tk.Frame(item, bg=CARD)
            head.pack(fill="x", padx=S(6), pady=(S(4), S(2)))
            badge = tk.Label(head, text=str(n), font=F_BODY_BOLD, fg="white", bg=COMPARE_HEX,
                             padx=S(6))
            badge.pack(side="left")
            title = tk.Label(head, text=kind, font=F_BODY_BOLD, fg=TEXT, bg=CARD)
            title.pack(side="left", padx=(S(8), 0))
            size = tk.Label(head, text=self._where(diff), font=F_TILE_SUB, fg=MUTED, bg=CARD)
            size.pack(side="right")
            widgets = [item, head, badge, title, size]
            if thumbs is not None and n - 1 < len(thumbs):
                photo = ImageTk.PhotoImage(thumbs[n - 1])
                self.item_photos.append(photo)
                pic = tk.Label(item, image=photo, bg=CARD)
                pic.pack(padx=S(6), pady=(0, S(2)))
                caps = tk.Frame(item, bg=CARD)
                caps.pack(fill="x", padx=S(6), pady=(0, S(4)))
                c1 = tk.Label(caps, text="Image 1", font=F_TILE_SUB, fg=MUTED, bg=CARD)
                c1.pack(side="left", expand=True)
                c2 = tk.Label(caps, text="Image 2", font=F_TILE_SUB, fg=MUTED, bg=CARD)
                c2.pack(side="left", expand=True)
                widgets += [pic, caps, c1, c2]
            for w in widgets:
                w.bind("<Button-1>", lambda e, n=n: self.select(n, zoom=True))
                w.bind("<MouseWheel>", self._list_wheel)
            self.items.append((item, widgets))
        if len(diffs) > len(shown):
            tk.Label(self.list_frame, text=f"+ {len(diffs) - len(shown)} more (boxed in the picture)",
                     font=F_TILE_SUB, fg=MUTED, bg=CARD).pack(anchor="w", padx=S(12), pady=S(6))
        self.list_canvas.yview_moveto(0)

    def _where(self, diff):
        """Short position hint, e.g. 'line at y 120'."""
        if diff.marker2:
            return "not in Image 2"
        if diff.marker1:
            return "not in Image 1"
        return f"{diff.box1[2] - diff.box1[0]} × {diff.box1[3] - diff.box1[1]} px"

    def _list_wheel(self, event):
        self.list_canvas.yview_scroll(wheel_steps(event) * 3, "units")

    def _highlight_item(self):
        for k, (item, widgets) in enumerate(self.items, 1):
            on = k == self.selected
            bg = ACCENT_SOFT if on else CARD
            for w in widgets:
                if not (isinstance(w, tk.Label) and w.cget("bg") == COMPARE_HEX):
                    w.configure(bg=bg)
            item.configure(highlightbackground=ACCENT if on else CARD)
        if self.selected and self.selected <= len(self.items):
            item = self.items[self.selected - 1][0]
            self.list_canvas.update_idletasks()
            total = max(1, self.list_frame.winfo_height())
            top, bottom = item.winfo_y() / total, (item.winfo_y() + item.winfo_height()) / total
            first, last = self.list_canvas.yview()
            if top < first or bottom > last:
                self.list_canvas.yview_moveto(max(0.0, top - 0.02))

    def select(self, n, zoom=True):
        if self.result is None or not (1 <= n <= len(self.result.diffs)):
            return
        self.selected = n
        if zoom:
            self.zoom = n
        self._highlight_item()
        self._update_state()
        self.show_result()

    def step(self, delta):
        if self.result is None or not self.result.diffs:
            return
        n = len(self.result.diffs)
        cur = self.selected or (0 if delta > 0 else n + 1)
        self.select((cur - 1 + delta) % n + 1, zoom=True)

    def show_all(self):
        self.zoom = None
        self._update_state()
        self.show_result()

    # ---- showing the result -------------------------------------------

    def _view_changed(self):
        self._update_state()
        self.show_result()

    def _schedule_redraw(self):
        if self._resize_job:
            self.win.after_cancel(self._resize_job)
        self._resize_job = self.win.after(60, self.show_result)

    def _zoomed_view(self, cw, ch):
        """Close-up of the selected difference for the current view:
        (image, offsets, clips)."""
        diff = self.result.diffs[self.zoom - 1]
        view = self.view_var.get()
        h1, h2 = self.result.image1, self.result.image2
        if view == "side":
            r1, r2 = zoom_rects(self.result, diff)
            zw, zh = max(r1[2] - r1[0], r2[2] - r2[0]), max(r1[3] - r1[1], r2[3] - r2[1])
            gap, bar = max(12, zw // 25), max(22, zh // 12)
            comp = Image.new("RGB", (zw * 2 + gap, zh + bar), (255, 255, 255))
            comp.paste(h1.crop(r1), (0, bar))
            comp.paste(h2.crop(r2), (zw + gap, bar))
            d = ImageDraw.Draw(comp)
            d.rectangle([zw, 0, zw + gap - 1, zh + bar], fill=(201, 206, 216))
            font = text_font(min(18, max(11, bar - 10)))
            d.text((4, 3), "IMAGE 1  (SOURCE)", fill=(27, 31, 42), font=font)
            d.text((zw + gap + 4, 3), "IMAGE 2", fill=(27, 31, 42), font=font)
            offsets = [(-r1[0], bar - r1[1]), (zw + gap - r2[0], bar - r2[1])]
            clips = [(0, bar, zw, bar + zh), (zw + gap, bar, zw * 2 + gap, bar + zh)]
            return comp, offsets, clips
        k = 0 if view == "image1" else 1
        r = zoom_rects(self.result, diff, ratio=cw / max(1, ch))[k]
        img = (h1, h2)[k]
        offsets = [None, None]
        offsets[k] = (-r[0], -r[1])
        return img.crop(r), offsets, None

    def show_result(self):
        self._resize_job = None
        c = self.canvas
        c.delete("all")
        cw, ch = max(20, c.winfo_width()), max(20, c.winfo_height())
        current = self.bases.get(self.view_var.get())
        if current is None:
            self.scroll.pack_forget()
            self.transform = None
            msg = ("Comparing..." if self.busy else
                   "Click Compare to check Image 2 against Image 1" if all(img is not None for img in self.images) else
                   "Load Image 1 (the source) and Image 2 above")
            c.create_text(cw // 2, ch // 2, text=msg, fill="#9AA3B2", font=F_TILE)
            c.configure(scrollregion=(0, 0, cw, ch))
            return
        base, offsets, gutters = current
        clips = None
        if self.zoom:
            gutters = [None, None]
            base, offsets, clips = self._zoomed_view(cw, ch)
            vw, vh = base.size
            scale = min(4.0, (cw - 8) / vw, (ch - 8) / vh)
            tall = False
        else:
            vw, vh = base.size
            fit_w = min(1.0, (cw - 4) / vw)
            fit_both = min(fit_w, (ch - 4) / vh)
            tall = vh > vw and fit_both < fit_w * 0.6
            scale = min(1.0, (cw - 20) / vw) if tall else fit_both
        size = (max(1, int(vw * scale)), max(1, int(vh * scale)))
        shown = base if size == base.size else base.resize(size, Image.LANCZOS if scale < 1 else Image.BICUBIC)
        self.photo = ImageTk.PhotoImage(shown)
        if tall:
            if not self.scroll.winfo_ismapped():
                self.scroll.pack(side="right", fill="y", before=c)
            ox, oy = max(0, (cw - size[0]) // 2), 0
            c.configure(scrollregion=(0, 0, cw, size[1]))
        else:
            self.scroll.pack_forget()
            ox, oy = (cw - size[0]) // 2, (ch - size[1]) // 2
            c.configure(scrollregion=(0, 0, cw, ch))
            c.yview_moveto(0)
        c.create_image(ox, oy, anchor="nw", image=self.photo)
        self.transform = (scale, ox, oy, 0, 0)
        self.drawn_offsets = offsets
        self.drawn = (offsets, clips, gutters)
        self._draw_boxes()

    def _screen_box(self, diff, k, off):
        """Where difference box k (0 = Image 1, 1 = Image 2) is on the canvas."""
        scale, ox, oy, vx0, vy0 = self.transform
        x0, y0, x1, y1 = diff.box1 if k == 0 else diff.box2
        dx, dy = off
        return (ox + (x0 + dx - vx0) * scale, oy + (y0 + dy - vy0) * scale,
                ox + (x1 + dx - vx0) * scale, oy + (y1 + dy - vy0) * scale)

    def _draw_boxes(self):
        """The marks, drawn at screen size so they stay clear at any zoom and
        never cover the content: a red outline with a gap and a white edge
        around each difference, a bar in the margin beside the same rows,
        and the number in the margin. Missing content: a dashed red line.
        The selected difference, and the one under the mouse, are drawn
        bolder - in both images, so you can see which parts belong together."""
        c = self.canvas
        c.delete("marks")
        if not self.result or not self.result.diffs or not self.transform:
            return
        offsets, clips, gutters = self.drawn
        scale, ox, oy = self.transform[:3]
        min_px = S(16)
        badge_font = font(8.5, "Segoe UI", "bold")
        bw_base, bh = S(8), S(18)
        next_free = [-1e9, -1e9]
        for n, diff in enumerate(self.result.diffs, 1):
            strong = n in (self.selected, self.hover)
            for k, off in enumerate(offsets):
                if off is None:
                    continue
                marker = diff.marker1 if k == 0 else diff.marker2
                box = diff.box1 if k == 0 else diff.box2
                if clips:                      # zoomed side by side: keep each box in its half
                    cl = clips[k]
                    bx = (box[0] + off[0], box[1] + off[1], box[2] + off[0], box[3] + off[1])
                    if bx[2] < cl[0] or bx[0] > cl[2] or bx[3] < cl[1] or bx[1] > cl[3]:
                        continue
                sx0, sy0, sx1, sy1 = self._screen_box(diff, k, off)
                tags = ("marks", f"diff{n}")
                w = 4 if strong else 2
                if marker:
                    sx1 = max(sx1, sx0 + S(40))
                    c.create_line(sx0, sy0, sx1, sy0, fill="white", width=w + 3, tags=("marks",))
                    c.create_line(sx0, sy0, sx1, sy0, fill=COMPARE_HEX, width=w + 1, dash=(6, 4), tags=tags)
                    c.create_rectangle(sx0 - 4, sy0 - 7, sx1 + 4, sy0 + 7, outline="", fill="", tags=tags)
                    by0, by1 = sy0 - 3, sy0 + 3
                    bx0 = sx0
                else:
                    if sx1 - sx0 < min_px:         # tiny changes still get a box you can see
                        m = (sx0 + sx1) / 2
                        sx0, sx1 = m - min_px / 2, m + min_px / 2
                    if sy1 - sy0 < min_px:
                        m = (sy0 + sy1) / 2
                        sy0, sy1 = m - min_px / 2, m + min_px / 2
                    gap = 5                          # keep the outline off the content
                    bx0, by0, bx1, by1 = sx0 - gap, sy0 - gap, sx1 + gap, sy1 + gap
                    c.create_rectangle(bx0, by0, bx1, by1, outline="white", width=w + 3, tags=("marks",))
                    c.create_rectangle(bx0, by0, bx1, by1, outline=COMPARE_HEX, width=w, tags=tags)
                # number + bar in the margin, or (no margin / too narrow) just above the box
                bw = bw_base + 7 * len(str(n))
                gut = gutters[k] if gutters and gutters[k] else None
                if gut:
                    gx0, gx1 = ox + gut[0] * scale, ox + gut[1] * scale
                else:
                    gx0 = gx1 = None
                if gut and gx1 - gx0 >= bw + 8:
                    c.create_rectangle(gx1 - (7 if strong else 5), by0, gx1 - 2, max(by1, by0 + 6),
                                       fill=COMPARE_HEX, outline="", tags=tags)
                    lx = gx0 + max(1, (gx1 - gx0 - 8 - bw) / 2)
                    ty = max(by0, next_free[k])
                    next_free[k] = ty + bh + 2
                elif marker:
                    lx, ty = bx0 - bw - 4, sy0 - bh / 2
                    if lx < 0:
                        lx, ty = bx0, by0 - bh - 2
                else:
                    lx, ty = bx0, (by0 - bh - 1 if by0 - bh - 1 > 0 else by1 + 1)
                c.create_rectangle(lx, ty, lx + bw, ty + bh, fill=COMPARE_HEX,
                                   outline="white" if strong else COMPARE_HEX, width=2 if strong else 1, tags=tags)
                c.create_text(lx + bw / 2, ty + bh / 2, text=str(n), fill="white", font=badge_font, tags=tags)

    def _diff_at(self, x, y):
        """Number of the difference at canvas point (x, y), or None."""
        if not self.result or not self.result.diffs or not self.transform:
            return None
        for item in reversed(self.canvas.find_overlapping(x - 3, y - 3, x + 3, y + 3)):
            for tag in self.canvas.gettags(item):
                if tag.startswith("diff"):
                    return int(tag[4:])
        for n, diff in enumerate(self.result.diffs, 1):
            for k, off in enumerate(self.drawn_offsets):
                if off is None:
                    continue
                sx0, sy0, sx1, sy1 = self._screen_box(diff, k, off)
                if sx0 - 8 <= x <= sx1 + 8 and sy0 - 8 <= y <= sy1 + 8:
                    return n
        return None

    def _canvas_motion(self, event):
        """Pointing at a difference outlines it boldly in both images."""
        n = self._diff_at(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if n != self.hover:
            self.hover = n
            self.canvas.configure(cursor="hand2" if n else "")
            self._draw_boxes()

    def _canvas_click(self, event):
        """Click a difference to select (and zoom to) it."""
        n = self._diff_at(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if n:
            self.select(n, zoom=True)

    def _canvas_wheel(self, event):
        if self.scroll.winfo_ismapped():
            self.canvas.yview_scroll(wheel_steps(event) * 3, "units")

    # ---- output --------------------------------------------------------

    def result_image(self):
        """The current view as an image, with numbered boxes drawn in."""
        current = self.bases.get(self.view_var.get())
        if current is None:
            return None
        base, offsets, gutters = current
        return draw_difference_boxes(base, self.result.diffs, offsets, gutters, self.selected)

    def copy_result(self):
        img = self.result_image()
        if img is None:
            return
        try:
            ok = copy_image_to_clipboard(img)
        except Exception as e:
            print(f"Copy failed: {e}")
            ok = False
        if not ok:
            messagebox.showerror("Copy failed", "Could not copy the image to the clipboard.",
                                 parent=self.win)
            return
        self.copy_button.config(text="Copied")
        self.win.after(1500, lambda: self.copy_button.winfo_exists()
                       and self.copy_button.config(text="Copy result"))
        status_var.set(f"Comparison copied to the clipboard - paste it with {KEY}+V")

    def save_result(self):
        """Open the result in the Review window (annotate, then save to the
        folder or the selected Word / Excel / PowerPoint file)."""
        img = self.result_image()
        if img is None or not ensure_save_folder():
            return
        taken_at = datetime.now()
        AnnotateEditor(img, on_save=lambda result: handle_capture(result, "Comparison", taken_at),
                       title="Comparison result",
                       save_label=f"Save to {doc_app()}" if word_doc_path is not None else "Save")

    def close(self):
        CompareWindow.current = None
        self.win.destroy()


def open_compare_window():
    if recorder is not None:
        return
    CompareWindow()


# ============================================================
# Capture Full Screen
# ============================================================

def capture_screen():
    if not ensure_save_folder():
        status_var.set("No save folder selected - capture cancelled")
        return

    # Hide the utility window
    root.withdraw()

    # Wait until the window is completely hidden
    root.after(500, perform_full_screen_capture)


def perform_full_screen_capture():
    try:
        taken_at = datetime.now()
        screenshot = grab_screen()
        deliver_capture(screenshot, "Full screen", taken_at)

    except Exception as e:
        print(f"Error capturing screen: {e}")

    finally:
        root.deiconify()
        root.lift()


# ============================================================
# Region Selection Overlay (shared by screenshot and recording)
# ============================================================

def start_region_selection(on_selected, hint=None, on_cancel=None):
    """
    Hide the main window and show the selection overlay.
    on_selected(left, top, right, bottom) is called with screen
    coordinates after the user drags out a rectangle.
    """
    root.withdraw()
    root.after(200, lambda: create_selection_overlay(on_selected, hint, on_cancel))


def create_selection_overlay(on_selected, hint=None, on_cancel=None):
    overlay = tk.Toplevel(root)
    if IS_MAC:     # macOS "full screen" opens a new Space; cover the screen instead
        overlay.overrideredirect(True)
        overlay.geometry(f"{overlay.winfo_screenwidth()}x{overlay.winfo_screenheight()}+0+0")
    else:
        overlay.attributes("-fullscreen", True)
    overlay.attributes("-topmost", True)
    overlay.attributes("-alpha", 0.25)
    overlay.configure(bg="black")

    canvas = tk.Canvas(
        overlay,
        cursor="cross",
        bg="black",
        highlightthickness=0
    )
    canvas.pack(fill="both", expand=True)

    if hint:
        overlay.update_idletasks()
        canvas.create_text(overlay.winfo_width() // 2, S(40), text=hint,
                           fill="white", font=font(14, "Segoe UI Semibold"))

    start_x = 0
    start_y = 0
    start_x_root = 0
    start_y_root = 0
    rectangle = None

    def mouse_down(event):
        nonlocal start_x, start_y, start_x_root, start_y_root, rectangle
        start_x = event.x
        start_y = event.y
        start_x_root = event.x_root
        start_y_root = event.y_root
        rectangle = canvas.create_rectangle(
            start_x, start_y, start_x, start_y,
            outline="red",
            width=2
        )

    def mouse_move(event):
        if rectangle is not None:
            canvas.coords(rectangle, start_x, start_y, event.x, event.y)

    def mouse_up(event):
        # Screen coordinates of the selection
        x1 = min(start_x_root, event.x_root)
        y1 = min(start_y_root, event.y_root)
        x2 = max(start_x_root, event.x_root)
        y2 = max(start_y_root, event.y_root)

        overlay.destroy()

        # Make sure selection has some size
        if x2 - x1 < 2 or y2 - y1 < 2:
            root.deiconify()
            root.lift()
            status_var.set("Selection too small - cancelled")
            if on_cancel:
                on_cancel()
            return

        # Wait for overlay to disappear
        root.after(200, lambda: on_selected(x1, y1, x2, y2))

    def cancel_selection(event):
        overlay.destroy()
        root.deiconify()
        root.lift()
        status_var.set("Selection cancelled")
        if on_cancel:
            on_cancel()

    canvas.bind("<ButtonPress-1>", mouse_down)
    canvas.bind("<B1-Motion>", mouse_move)
    canvas.bind("<ButtonRelease-1>", mouse_up)
    overlay.bind("<Escape>", cancel_selection)

    overlay.focus_force()


# ============================================================
# Capture Selected Region (screenshot)
# ============================================================

def select_region():
    if not ensure_save_folder():
        status_var.set("No save folder selected - capture cancelled")
        return

    start_region_selection(capture_selected_region)


def capture_selected_region(x1, y1, x2, y2):
    try:
        taken_at = datetime.now()
        screenshot = grab_screen((x1, y1, x2, y2))
        deliver_capture(screenshot, "Region", taken_at)

    except Exception as e:
        print(f"Error capturing selected region: {e}")

    finally:
        root.deiconify()
        root.lift()


# ============================================================
# Screen Recorder
# ============================================================
# Video: screen frames are grabbed on a fixed clock and piped into
# ffmpeg, which encodes them as H.264 into a temporary file.
#
# Audio: the microphone and the computer's own sound (loopback) are
# each recorded into a temporary WAV file. Every audio chunk is placed
# on the SAME clock the video uses (the recording's elapsed time, not
# counting pauses), so gaps are filled with silence and drift is
# corrected. That keeps sound and picture in sync.
#
# When recording stops, ffmpeg mixes the audio tracks and combines them
# with the video into the final .mp4.

def hidden_temp_path(final_path, suffix):
    """A temporary file next to the final recording, e.g.
    .Recording_2026-09-29_21-30-00.mic.wav"""
    return final_path.with_name(f".{final_path.stem}{suffix}")


class AudioTrack:
    """One audio source written to a WAV file, aligned to the recording clock."""

    def __init__(self, label, path, rate, channels, clock):
        self.label = label
        self.path = path
        self.rate = rate
        self.channels = channels
        self.clock = clock                 # the ScreenRecorder
        self.frame_bytes = 2 * channels    # 16-bit samples
        self.frames = 0                    # audio frames written so far
        self.lock = threading.Lock()
        self.closed = False
        self.received_audio = False

        self.wav = wave.open(str(path), "wb")
        self.wav.setnchannels(channels)
        self.wav.setsampwidth(2)
        self.wav.setframerate(rate)

    def feed(self, data, frame_count):
        """Called from the audio callback with raw 16-bit PCM bytes."""
        if self.clock.paused or frame_count <= 0:
            return

        with self.lock:
            if self.closed:
                return

            self.received_audio = True

            # Where this chunk belongs on the recording timeline
            end_time = self.clock.elapsed()
            expected_start = int(end_time * self.rate) - frame_count
            difference = expected_start - self.frames
            tolerance = int(self.rate * AUDIO_MAX_DRIFT)

            if difference > tolerance:
                # Audio is behind the clock (e.g. the computer was silent
                # and Windows sent nothing): fill the gap with silence.
                self.wav.writeframes(b"\x00" * (difference * self.frame_bytes))
                self.frames += difference

            elif difference < -tolerance:
                # Audio is ahead of the clock: drop the extra samples.
                drop = min(frame_count, -difference)
                data = data[drop * self.frame_bytes:]
                frame_count -= drop

            if frame_count > 0:
                self.wav.writeframes(data[:frame_count * self.frame_bytes])
                self.frames += frame_count

    def finish(self, duration):
        """Pad with silence to the full video length and close the file."""
        with self.lock:
            if self.closed:
                return
            self.closed = True
            target = int(duration * self.rate)
            if self.frames < target:
                self.wav.writeframes(
                    b"\x00" * ((target - self.frames) * self.frame_bytes))
                self.frames = target
            self.wav.close()


class AudioCapture:
    """Records the default microphone and/or the default speakers
    (loopback = everything the computer plays) using WASAPI."""

    def __init__(self, final_path, clock, want_mic, want_system):
        self.final_path = final_path
        self.clock = clock
        self.want_mic = want_mic
        self.want_system = want_system
        self.pa = None
        self.streams = []
        self.tracks = []
        self.notes = []

    def start(self):
        if not (self.want_mic or self.want_system):
            return
        if not AUDIO_AVAILABLE:
            self.notes.append("audio recording isn't available")
            return

        try:
            self.pa = pyaudio.PyAudio()
            wasapi = self.pa.get_host_api_info_by_type(pyaudio.paWASAPI)
        except Exception as e:
            self.notes.append(f"audio unavailable ({e})")
            return

        if self.want_system:
            try:
                self._open(self._find_loopback(wasapi), "system",
                           "computer sound")
            except Exception as e:
                self.notes.append(f"computer sound not recorded ({e})")

        if self.want_mic:
            try:
                index = wasapi["defaultInputDevice"]
                if index is None or index < 0:
                    raise RuntimeError("no microphone found")
                self._open(self.pa.get_device_info_by_index(index), "mic",
                           "microphone")
            except Exception as e:
                self.notes.append(f"microphone not recorded ({e})")

    def _find_loopback(self, wasapi):
        try:
            return self.pa.get_default_wasapi_loopback()
        except Exception:
            pass

        speakers = self.pa.get_device_info_by_index(wasapi["defaultOutputDevice"])
        if speakers.get("isLoopbackDevice"):
            return speakers
        for loopback in self.pa.get_loopback_device_info_generator():
            if speakers["name"] in loopback["name"]:
                return loopback
        raise RuntimeError("no loopback device for the default speakers")

    def _open(self, device, file_label, label):
        rate = int(device["defaultSampleRate"])
        channels = max(1, int(device["maxInputChannels"]))
        path = hidden_temp_path(self.final_path, f".{file_label}.wav")

        track = AudioTrack(label, path, rate, channels, self.clock)

        def callback(in_data, frame_count, time_info, status):
            track.feed(in_data, frame_count)
            return (None, pyaudio.paContinue)

        try:
            stream = self.pa.open(
                format=pyaudio.paInt16,
                channels=channels,
                rate=rate,
                input=True,
                input_device_index=device["index"],
                frames_per_buffer=1024,
                stream_callback=callback,
            )
            stream.start_stream()
        except Exception:
            track.finish(0)
            try:
                path.unlink()
            except Exception:
                pass
            raise

        self.streams.append(stream)
        self.tracks.append(track)
        print(f"Recording {label} from: {device['name']}")

    def stop(self):
        for stream in self.streams:
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
        self.streams.clear()
        if self.pa is not None:
            try:
                self.pa.terminate()
            except Exception:
                pass
            self.pa = None


class FfmpegVideoSink:
    """Encodes raw RGB frames to H.264 with ffmpeg."""

    def __init__(self, path, width, height, fps):
        self.path = path
        self.log_path = hidden_temp_path(path, ".log")
        self.log = open(self.log_path, "wb")
        self.proc = subprocess.Popen(
            [FFMPEG_EXE, "-y", "-loglevel", "error",
             "-f", "rawvideo", "-pix_fmt", "rgb24",
             "-s", f"{width}x{height}", "-r", str(fps), "-i", "-",
             "-an", "-c:v", "libx264", "-preset", "veryfast", "-crf", "23",
             "-pix_fmt", "yuv420p", str(path)],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=self.log,
            creationflags=NO_WINDOW,
        )

    def write(self, frame_bytes):
        self.proc.stdin.write(frame_bytes)

    def close(self):
        error = None
        try:
            self.proc.stdin.close()
        except Exception:
            pass
        if self.proc.wait() != 0:
            error = "Video encoding failed."
        self.log.close()
        try:
            text = self.log_path.read_text(errors="replace").strip()
            if error and text:
                error += " " + text.splitlines()[-1]
            self.log_path.unlink()
        except Exception:
            pass
        return error


class Cv2VideoSink:
    """Fallback when ffmpeg isn't available: OpenCV, no audio."""

    def __init__(self, path, width, height, fps):
        self.width, self.height = width, height
        self.writer = cv2.VideoWriter(str(path), cv2.VideoWriter_fourcc(*"mp4v"),
                                      fps, (width, height))
        if not self.writer.isOpened():
            raise RuntimeError("Could not open the video file for writing.")

    def write(self, frame_bytes):
        frame = np.frombuffer(frame_bytes, dtype=np.uint8).reshape(
            self.height, self.width, 3)
        self.writer.write(cv2.cvtColor(frame, cv2.COLOR_RGB2BGR))

    def close(self):
        self.writer.release()
        return None


class ScreenRecorder:
    """
    Records an area of the screen (plus optional audio) to an .mp4 file.
    Frames are written on a fixed clock so the video plays back at
    real-time speed even if capturing is sometimes slow. Recording can
    be paused and resumed; paused time is left out of video and audio.
    """

    def __init__(self, filename, area, fps=RECORDING_FPS,
                 record_mic=False, record_system=False):
        self.filename = filename
        self.fps = fps

        # Video encoders need even width/height
        width = area["width"] - (area["width"] % 2)
        height = area["height"] - (area["height"] % 2)

        self.area = {
            "left": area["left"],
            "top": area["top"],
            "width": width,
            "height": height,
        }

        self.stop_event = threading.Event()
        self.lock = threading.Lock()
        self.thread = None
        self.error = None
        self.warning = None
        self.frames_written = 0
        self.start_time = None
        self.paused = False
        self.pause_started = None
        self.paused_total = 0.0

        self.use_ffmpeg = FFMPEG_EXE is not None
        self.video_path = (hidden_temp_path(filename, ".video.mp4")
                           if self.use_ffmpeg else filename)
        self.audio = AudioCapture(filename, self,
                                  record_mic and self.use_ffmpeg,
                                  record_system and self.use_ffmpeg)

    # ---- clock -----------------------------------------------------

    def elapsed(self):
        """Recorded time in seconds (not counting paused time)."""
        with self.lock:
            if self.start_time is None:
                return 0
            now = self.pause_started if self.paused else time.perf_counter()
            return now - self.start_time - self.paused_total

    def pause(self):
        with self.lock:
            if not self.paused and self.start_time is not None:
                self.paused = True
                self.pause_started = time.perf_counter()

    def resume(self):
        with self.lock:
            if self.paused:
                self.paused_total += time.perf_counter() - self.pause_started
                self.paused = False
                self.pause_started = None

    # ---- start / stop ----------------------------------------------

    def start(self):
        # Video and audio share this one clock, which keeps them in sync
        with self.lock:
            self.start_time = time.perf_counter()
        self.audio.start()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def is_running(self):
        return self.thread is not None and self.thread.is_alive()

    def duration(self):
        return self.frames_written / self.fps

    def stop(self):
        """Stop capturing (may take a moment while ffmpeg finishes)."""
        self.stop_event.set()
        self.audio.stop()
        if self.thread is not None:
            self.thread.join()

    def finalize(self):
        """Combine video and audio into the final .mp4 file.
        Runs after stop(), in a background thread."""
        tracks = self.audio.tracks
        duration = self.duration()

        for track in tracks:
            track.finish(duration)

        try:
            if self.error or not self.use_ffmpeg:
                return

            if self.frames_written == 0:
                self.error = "No frames were captured, so the video is empty."
                return

            if not tracks:
                os.replace(self.video_path, self.filename)
                return

            command = [FFMPEG_EXE, "-y", "-loglevel", "error",
                       "-i", str(self.video_path)]
            for track in tracks:
                command += ["-i", str(track.path)]

            # Bring every track to 48 kHz stereo, then mix them together
            filters = []
            for i in range(len(tracks)):
                filters.append(
                    f"[{i + 1}:a]aresample=48000,"
                    f"aformat=sample_fmts=fltp:channel_layouts=stereo[a{i}]")
            if len(tracks) == 1:
                audio_out = "[a0]"
            else:
                inputs = "".join(f"[a{i}]" for i in range(len(tracks)))
                filters.append(
                    f"{inputs}amix=inputs={len(tracks)}:duration=longest:"
                    f"normalize=0,alimiter=limit=0.95[mix]")
                audio_out = "[mix]"

            command += ["-filter_complex", ";".join(filters),
                        "-map", "0:v", "-map", audio_out,
                        "-c:v", "copy", "-c:a", "aac", "-b:a", "192k",
                        "-t", f"{duration:.3f}",
                        "-movflags", "+faststart",
                        str(self.filename)]

            result = subprocess.run(command, capture_output=True,
                                    creationflags=NO_WINDOW)

            if result.returncode != 0:
                # Keep the video even if adding the sound failed
                details = result.stderr.decode(errors="replace").strip()
                print(f"Adding audio failed: {details}")
                os.replace(self.video_path, self.filename)
                self.warning = "the sound couldn't be added (see console)"

        except Exception as e:
            self.error = f"Could not save the recording: {e}"

        finally:
            for path in [self.video_path] + [t.path for t in tracks]:
                if path != self.filename:
                    try:
                        path.unlink()
                    except Exception:
                        pass

    def audio_summary(self):
        labels = [t.label for t in self.audio.tracks]
        parts = []
        if labels:
            parts.append("with " + " + ".join(labels))
        elif self.audio.want_mic or self.audio.want_system:
            parts.append("without sound")
        parts += self.audio.notes
        if self.warning:
            parts.append(self.warning)
        return "; ".join(parts)

    # ---- capture loop ----------------------------------------------

    def _run(self):
        width = self.area["width"]
        height = self.area["height"]

        try:
            if self.use_ffmpeg:
                sink = FfmpegVideoSink(self.video_path, width, height, self.fps)
            else:
                sink = Cv2VideoSink(self.video_path, width, height, self.fps)
        except Exception as e:
            self.error = str(e)
            return

        frame_interval = 1.0 / self.fps

        bbox = (
            self.area["left"],
            self.area["top"],
            self.area["left"] + width,
            self.area["top"] + height,
        )

        last_frame = None

        try:
            while not self.stop_event.is_set():
                if self.paused:
                    self.stop_event.wait(0.05)
                    continue

                # The moment this frame shows (on the shared clock)
                capture_time = self.elapsed()
                shot = grab_screen(bbox)

                # Paused while grabbing? Drop this frame.
                if self.paused:
                    continue

                shot = shot.convert("RGB")
                if shot.size != (width, height):
                    shot = shot.crop((0, 0, width, height))
                frame_bytes = shot.tobytes()

                # Put the frame at its own time in the video. If capturing
                # fell behind, repeat the previous frame to fill the gap,
                # so the video always runs in real time (and in sync with
                # the audio, which uses the same clock).
                index = round(capture_time * self.fps)
                filler = last_frame if last_frame is not None else frame_bytes
                while self.frames_written < index:
                    sink.write(filler)
                    self.frames_written += 1
                if self.frames_written <= index:
                    sink.write(frame_bytes)
                    self.frames_written += 1
                last_frame = frame_bytes

                # Sleep until the next frame is due
                delay = self.frames_written * frame_interval - self.elapsed()
                if delay > 0:
                    self.stop_event.wait(delay)

            # Hold the last frame until the moment Stop was pressed
            if last_frame is not None:
                end_index = round(self.elapsed() * self.fps)
                while self.frames_written < end_index:
                    sink.write(last_frame)
                    self.frames_written += 1

        except Exception as e:
            self.error = str(e)

        finally:
            close_error = sink.close()
            if close_error and not self.error:
                self.error = close_error


# ============================================================
# Recording Mode (compact control bar)
# ============================================================
# While recording, the main window shrinks into a small control bar
# (blinking dot, timer, Pause/Resume, Stop). It is a normal window with
# a title bar, so it is always visible and shows on the taskbar.
#
# On Windows 10 (version 2004) and later the bar is also excluded from
# screen capture, so it doesn't appear in the video. For region
# recordings it is additionally placed outside the recorded area.

WDA_NONE = 0x00000000
WDA_EXCLUDEFROMCAPTURE = 0x00000011
GA_ROOT = 2

DOT_ON = "#ff4b4b"
DOT_PAUSED = "#fbbc04"


def window_handle(window):
    user32 = ctypes.windll.user32
    return user32.GetAncestor(window.winfo_id(), GA_ROOT)


def set_capture_excluded(window, excluded):
    """Hide/show a window in screenshots and recordings.
    Returns True if Windows accepted the setting."""
    try:
        window.update_idletasks()
        if IS_MAC:
            return mac_set_windows_shared(not excluded)
        affinity = WDA_EXCLUDEFROMCAPTURE if excluded else WDA_NONE
        return bool(ctypes.windll.user32.SetWindowDisplayAffinity(
            window_handle(window), affinity))
    except Exception as e:
        print(f"Could not change capture visibility: {e}")
        return False


def virtual_screen_bounds():
    """(left, top, width, height) of the whole desktop, all monitors."""
    try:
        user32 = ctypes.windll.user32
        bounds = (user32.GetSystemMetrics(76), user32.GetSystemMetrics(77),
                  user32.GetSystemMetrics(78), user32.GetSystemMetrics(79))
        if bounds[2] > 0 and bounds[3] > 0:
            return bounds
    except Exception:
        pass
    return (0, 0, root.winfo_screenwidth(), root.winfo_screenheight())


saved_main_position = None
bar_excluded = False
dot_visible = True


def bar_position_for(area, bar=None):
    """Top-left corner for the control bar: just below the recorded
    area, else just above it, else inside its top edge."""
    bar = bar or rec_bar
    root.update_idletasks()
    w = bar.winfo_reqwidth() + 20          # + window border
    h = bar.winfo_reqheight() + 45         # + title bar
    vx, vy, vw, vh = virtual_screen_bounds()

    left, top = area["left"], area["top"]
    right, bottom = left + area["width"], top + area["height"]
    gap = 10

    x = left + (area["width"] - w) // 2
    x = max(vx + gap, min(x, vx + vw - w - gap))

    if bottom + gap + h <= vy + vh - 50:      # below (clear of the taskbar)
        y = bottom + gap
    elif top - gap - h >= vy:                 # above
        y = top - gap - h
    else:                                     # inside, near the top
        y = top + gap

    return x, y


def enter_recording_mode(area):
    """Turn the main window into the small control bar."""
    global saved_main_position, bar_excluded, dot_visible

    saved_main_position = (root.winfo_x(), root.winfo_y())

    main_frame.pack_forget()
    rec_bar.pack()
    rec_time_label.config(text="00:00")
    rec_pause_button.config(text="Pause")
    rec_dot.itemconfigure(rec_dot_id, fill=DOT_ON, state="normal")
    dot_visible = True

    root.title("Recording")
    root.attributes("-topmost", True)

    x, y = bar_position_for(area)
    root.geometry(f"+{x}+{y}")
    root.deiconify()
    root.lift()
    root.update()

    bar_excluded = set_capture_excluded(root, True)
    if bar_excluded:
        print("Recording controls are hidden from the video.")
    else:
        print("Note: the recording controls can't be hidden from capture here.")


def enter_compact_mode(bar, area, title):
    """Show only a small bar (e.g. full-page capture progress) instead of
    the main window, placed outside the captured area and hidden from
    screen capture where Windows supports it."""
    global saved_main_position
    saved_main_position = (root.winfo_x(), root.winfo_y())
    main_frame.pack_forget()
    bar.pack()
    root.title(title)
    root.attributes("-topmost", True)
    x, y = bar_position_for(area, bar)
    root.geometry(f"+{x}+{y}")
    root.deiconify()
    root.lift()
    root.update()
    set_capture_excluded(root, True)


def exit_compact_mode(bar):
    set_capture_excluded(root, False)
    bar.pack_forget()
    main_frame.pack()
    root.title("Screen Capture Utility")
    root.attributes("-topmost", False)
    if saved_main_position is not None:
        root.geometry(f"+{saved_main_position[0]}+{saved_main_position[1]}")
    root.deiconify()
    root.lift()


def exit_recording_mode():
    """Turn the control bar back into the main window."""
    set_capture_excluded(root, False)

    rec_bar.pack_forget()
    main_frame.pack()

    root.title("Screen Capture Utility")
    root.attributes("-topmost", False)

    if saved_main_position is not None:
        root.geometry(f"+{saved_main_position[0]}+{saved_main_position[1]}")

    root.deiconify()
    root.lift()


# ============================================================
# Recording Controls
# ============================================================

recorder = None


def primary_monitor_area():
    if IS_MAC:                     # macOS records in screen points
        return {"left": 0, "top": 0,
                "width": root.winfo_screenwidth(), "height": root.winfo_screenheight()}
    # A plain grab() (without all_screens) captures the primary monitor
    width, height = ImageGrab.grab().size
    return {
        "left": 0,
        "top": 0,
        "width": width,
        "height": height,
    }


def toggle_screen_recording():
    """Start recording the primary monitor, or stop the current recording."""
    if recorder is not None:
        stop_recording()
        return

    if not RECORDING_AVAILABLE or saving_thread is not None:
        return

    if not ensure_save_folder():
        status_var.set("No save folder selected - recording cancelled")
        return

    area = primary_monitor_area()

    # Hide the full window first, then show the bar and start
    root.withdraw()
    root.after(300, lambda: start_recording(area))


def record_region():
    if recorder is not None or not RECORDING_AVAILABLE or saving_thread is not None:
        return

    if not ensure_save_folder():
        status_var.set("No save folder selected - recording cancelled")
        return

    start_region_selection(start_region_recording)


def start_region_recording(x1, y1, x2, y2):
    start_recording({
        "left": x1,
        "top": y1,
        "width": x2 - x1,
        "height": y2 - y1,
    })


def start_recording(area):
    global recorder

    try:
        enter_recording_mode(area)

        filename = get_recording_filename()
        recorder = (MacScreenRecorder if IS_MAC else ScreenRecorder)(
            filename, area,
            record_mic=mic_var.get() and AUDIO_AVAILABLE,
            record_system=system_audio_var.get() and SYSTEM_AUDIO_AVAILABLE,
        )
        recorder.start()
        print(f"Recording started: {filename}")

    except Exception as e:
        recorder = None
        exit_recording_mode()
        print(f"Could not start recording: {e}")
        status_var.set("Could not start recording (see console)")
        messagebox.showerror("Recording failed", str(e), parent=root)
        return

    # Give Windows a moment to show the bar, then start the timer
    root.after(200, update_recording_status)


def toggle_pause():
    if recorder is None:
        return

    if recorder.paused:
        recorder.resume()
        rec_pause_button.config(text="Pause")
        rec_dot.itemconfigure(rec_dot_id, fill=DOT_ON, state="normal")
        print("Recording resumed")
    else:
        recorder.pause()
        rec_pause_button.config(text="Resume")
        rec_dot.itemconfigure(rec_dot_id, fill=DOT_PAUSED, state="normal")
        print("Recording paused")


saving_thread = None


def stop_recording():
    """Stop recording, then save the video (and sound) in the background
    so the window stays responsive."""
    global recorder, saving_thread

    if recorder is None:
        return

    current = recorder
    recorder = None

    exit_recording_mode()
    set_record_buttons_state("disabled")
    status_var.set("Saving recording... please wait")

    def work():
        current.stop()
        current.finalize()

    saving_thread = threading.Thread(target=work, daemon=True)
    saving_thread.start()
    root.after(200, lambda: check_saving_done(current))


def check_saving_done(current):
    global saving_thread

    if saving_thread is not None and saving_thread.is_alive():
        root.after(200, lambda: check_saving_done(current))
        return

    saving_thread = None
    set_record_buttons_state("normal")

    if current.error:
        print(f"Error while recording: {current.error}")
        status_var.set(f"Recording failed: {current.error}")
        messagebox.showerror("Recording failed", current.error, parent=root)
        return

    print("Recording saved successfully:")
    print(current.filename)
    message = f"Saved: {current.filename.name}  ({current.duration():.1f} s)"
    audio = current.audio_summary()
    if audio:
        message += f" - {audio}"
    if not bar_excluded:
        message += ("\nNote: the controls may appear in the video" +
                    ("." if IS_MAC else " (hiding them needs Windows 10 version 2004 or later)."))
    status_var.set(message)


def set_record_buttons_state(state):
    if not RECORDING_AVAILABLE:
        state = "disabled"
    record_button.config(state=state)
    record_region_button.config(state=state)


def update_recording_status():
    """Refresh the timer on the control bar (runs on the UI thread)."""
    global dot_visible

    if recorder is None:
        return

    # The recording thread stopped on its own because of an error
    if not recorder.is_running():
        stop_recording()
        return

    seconds = int(recorder.elapsed())
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    rec_time_label.config(
        text=f"{hours}:{minutes:02d}:{seconds:02d}" if hours
        else f"{minutes:02d}:{seconds:02d}"
    )

    # Blink the dot while recording; solid yellow while paused
    if not recorder.paused:
        dot_visible = not dot_visible
        rec_dot.itemconfigure(rec_dot_id,
                              state="normal" if dot_visible else "hidden")

    root.after(500, update_recording_status)


# ============================================================
# Close Handling
# ============================================================

def on_close():
    # While recording, the window's X button just stops the recording
    if recorder is not None:
        stop_recording()
        return

    # During a full-page capture, the X button stops it (keeps what's captured)
    if ScrollingCapture.active is not None:
        ScrollingCapture.active.stop()
        return

    # Don't quit while a recording is still being saved
    if saving_thread is not None and saving_thread.is_alive():
        messagebox.showinfo(
            "Still saving",
            "The recording is still being saved. Please wait a moment.",
            parent=root
        )
        return

    # Make sure every screenshot ends up in the selected document
    while pending_captures and not flush_pending(interactive=True):
        retry = messagebox.askretrycancel(
            f"{doc_app()} file is open",
            f"{len(pending_captures)} capture(s) haven't been added to:\n\n"
            f"{word_doc_path}\n\n"
            f"The file is probably open in {doc_app()}. Close it there, "
            "then click Retry.\n\n"
            "Cancel saves them as separate .png files instead.",
            parent=root
        )
        if not retry:
            saved = save_pending_as_png()
            if saved:
                messagebox.showinfo(
                    "Saved as .png files",
                    f"{len(saved)} capture(s) were saved in:\n\n{saved[0].parent}",
                    parent=root
                )
            break

    root.destroy()


# ============================================================
# Main Window
# ============================================================

startup_step("packages loaded")
root = tk.Tk()
root.withdraw()   # shown, centred, once the layout is built
startup_step(f"window system ready (Tk {root.tk.call('info', 'patchlevel')}, "
             f"{root.tk.call('tk', 'windowingsystem')})")
root.title("Screen Capture Utility")
root.resizable(False, False)
root.protocol("WM_DELETE_WINDOW", on_close)

folder_var = tk.StringVar()
doc_var = tk.StringVar()
keep_png_var = tk.BooleanVar(value=False)
status_var = tk.StringVar(value="Choose a folder to save screenshots and recordings.")


# ============================================================
# Theme
# ============================================================

def screen_work_area():
    """(left, top, width, height) of the usable area (without the taskbar)
    of the monitor the mouse is on - the one the user is looking at."""
    if IS_MAC:
        try:
            return mac_work_area()
        except Exception:
            return 0, 25, root.winfo_screenwidth(), root.winfo_screenheight() - 25
    try:
        from ctypes import wintypes
        user32 = ctypes.windll.user32

        class MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT),
                        ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]

        user32.MonitorFromPoint.argtypes = [wintypes.POINT, wintypes.DWORD]
        user32.MonitorFromPoint.restype = ctypes.c_void_p
        user32.GetMonitorInfoW.argtypes = [ctypes.c_void_p, ctypes.POINTER(MONITORINFO)]

        cursor = wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(cursor))
        monitor = user32.MonitorFromPoint(cursor, 2)   # 2 = nearest monitor
        info = MONITORINFO()
        info.cbSize = ctypes.sizeof(MONITORINFO)
        if monitor and user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            r = info.rcWork
            return r.left, r.top, r.right - r.left, r.bottom - r.top
    except Exception:
        pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight()



from PIL import ImageDraw, ImageTk

# Colours
BG = "#E8EBF1"            # window background (darker so cards stand out)
CARD = "#FFFFFF"          # card background
CARD_HOVER = "#F4F4FF"    # tile hover
BORDER = "#C9CED8"          # card / tile borders
FIELD_BORDER = "#9AA3B2"    # text box borders
FIELD_READONLY = "#F2F4F8"  # read-only text box background
TEXT = "#1B1F2A"
MUTED = "#4B5563"
FAINT = "#7B8494"
ACCENT = "#4F46E5"        # indigo
ACCENT_HOVER = "#4338CA"
ACCENT_SOFT = "#EEF0FF"
RED = "#E5484D"
RED_SOFT = "#FDEDEE"
GREEN = "#16A34A"
AMBER = "#D97706"
HEADER_BG = "#141824"
HEADER_SUB = "#C5CAD6"
BAR_BG = "#141824"

# Size of everything follows the screen:
#   DPI_SCALE - Windows display scaling (100% = 1.0, 150% = 1.5)
#   UI_FIT    - how much room the screen has: the layout is designed for a
#               screen a bit bigger than 1920 x 1080; on smaller screens (laptops, 1366 x 768)
#               everything shrinks, on large ones it grows a little, so the
#               window always takes up a similar share of the screen.
# On a Mac, Tk works in points and macOS itself makes Retina screens sharp,
# so no extra scaling. (Some Mac set-ups report a nonsensical screen
# resolution to Tk, which would blow every icon up to an enormous size.)
_RAW_DPI = root.winfo_fpixels("1i")
DPI_SCALE = 1.0 if IS_MAC else min(4.0, max(1.0, _RAW_DPI / 96.0))


def _screen_fit():
    _, _, width, height = screen_work_area()
    logical_w, logical_h = width / DPI_SCALE, height / DPI_SCALE
    return max(0.72, min(1.1, logical_w / 1900, logical_h / 1150))


UI_FIT = _screen_fit()
UI_SCALE = DPI_SCALE * UI_FIT
startup_step(f"scale: screen reports {_RAW_DPI:.0f} dpi -> {DPI_SCALE:.2f} x {UI_FIT:.2f}; "
             f"work area {screen_work_area()}")
FONT_SCALE = DPI_SCALE * max(0.86, UI_FIT)      # text never gets too small to read


def S(pixels):
    return int(round(pixels * UI_SCALE))


UI_FAMILY = "Helvetica Neue" if IS_MAC else "Segoe UI"
MAC_FONTS = {"Segoe UI": ("Helvetica Neue", ()), "Segoe UI Semibold": ("Helvetica Neue", ("bold",)),
             "Consolas": ("Menlo", ())}


def font(points, family="Segoe UI", *style):
    """A font sized for this screen (Tk: negative size = pixels). On a Mac
    the Windows fonts are swapped for their macOS counterparts."""
    if IS_MAC and family in MAC_FONTS:
        family, extra = MAC_FONTS[family]
        style = tuple(dict.fromkeys(extra + style))
    return (family, -max(9, int(round(points * 4 / 3 * FONT_SCALE))), *style)


def wheel_steps(event):
    """Mouse wheel / trackpad as whole steps (Windows sends 120 per notch,
    macOS small numbers)."""
    if IS_MAC:
        return -int(event.delta)
    return int(-event.delta / 120)


# Fonts (Segoe UI on Windows; Tk falls back elsewhere)
F_TITLE = font(15, "Segoe UI Semibold")
F_SUB = font(9.5)
F_SECTION = font(9, "Segoe UI Semibold")
F_BODY = font(10)
F_BODY_BOLD = font(10, "Segoe UI Semibold")
F_TILE = font(10.5, "Segoe UI Semibold")
F_TILE_SUB = font(9)
F_TIMER = font(13, "Consolas", "bold")


root.configure(bg=BG)


# ============================================================
# Icons (drawn with Pillow, 4x supersampled for smooth edges)
# ============================================================

ICONS = {}   # keep references so Tk doesn't garbage-collect them


def _canvas(size):
    big = size * 4
    return Image.new("RGBA", (big, big), (0, 0, 0, 0)), big


def _finish(img, size):
    return img.resize((size, size), Image.LANCZOS)


def draw_icon(kind, color, size, accent=RED):
    img, n = _canvas(size)
    d = ImageDraw.Draw(img)
    w = max(4, n // 14)                      # stroke width
    u = n / 24.0                             # 24-unit grid

    def box(x0, y0, x1, y1):
        return [x0 * u, y0 * u, x1 * u, y1 * u]

    if kind in ("screen", "record"):
        d.rounded_rectangle(box(2, 4, 22, 17), radius=2 * u, outline=color, width=w)
        d.line(box(12, 17, 12, 20), fill=color, width=w)
        d.line(box(8, 20.5, 16, 20.5), fill=color, width=w)
        if kind == "record":
            d.ellipse(box(9, 7.5, 15, 13.5), fill=accent)

    elif kind in ("region", "record_region"):
        L = 5
        for (x, y, dx, dy) in [(3, 3, 1, 1), (21, 3, -1, 1), (3, 21, 1, -1), (21, 21, -1, -1)]:
            d.line([x * u, y * u, (x + dx * L) * u, y * u], fill=color, width=w)
            d.line([x * u, y * u, x * u, (y + dy * L) * u], fill=color, width=w)
        if kind == "record_region":
            d.ellipse(box(8.5, 8.5, 15.5, 15.5), fill=accent)
        else:
            d.rounded_rectangle(box(8, 8, 16, 16), radius=1.2 * u, fill=color)

    elif kind == "eraser":
        d.polygon([(13 * u, 3.5 * u), (21 * u, 11.5 * u), (12.5 * u, 20 * u), (4.5 * u, 12 * u)],
                  outline=color, width=w)
        d.polygon([(4.5 * u, 12 * u), (8.5 * u, 8 * u), (16.5 * u, 16 * u), (12.5 * u, 20 * u)], fill=color)
        d.line(box(9, 21, 21, 21), fill=color, width=w)

    elif kind == "copy":
        d.rounded_rectangle(box(8, 8, 21, 21), radius=2 * u, outline=color, width=w)
        d.line([(4 * u, 16 * u), (4 * u, 5 * u)], fill=color, width=w)
        d.line([(4 * u, 4 * u), (15 * u, 4 * u)], fill=color, width=w)
        d.arc(box(3, 3, 7, 7), start=180, end=270, fill=color, width=w)

    elif kind == "check":
        d.line([(4 * u, 12.5 * u), (9.5 * u, 18 * u), (20 * u, 6.5 * u)], fill=color, width=int(w * 1.4),
               joint="curve")

    elif kind == "highlight":
        d.polygon([(14 * u, 3 * u), (21 * u, 10 * u), (12 * u, 19 * u), (5 * u, 12 * u)], fill=color)
        d.polygon([(5 * u, 12 * u), (12 * u, 19 * u), (9 * u, 21 * u), (3 * u, 15 * u)], fill=color)
        d.line(box(2, 22, 13, 22), fill=color, width=w)

    elif kind == "pen":
        d.line([(5 * u, 19 * u), (18 * u, 6 * u)], fill=color, width=int(w * 2.2))
        d.polygon([(3 * u, 21 * u), (4 * u, 16.5 * u), (7.5 * u, 20 * u)], fill=color)
        d.line([(16 * u, 4 * u), (20 * u, 8 * u)], fill=color, width=int(w * 1.2))

    elif kind == "rect":
        d.rectangle(box(3, 5, 21, 19), outline=color, width=w)

    elif kind == "text":
        d.line([(4 * u, 5 * u), (20 * u, 5 * u)], fill=color, width=int(w * 1.2))
        d.line([(12 * u, 5 * u), (12 * u, 20 * u)], fill=color, width=int(w * 1.2))
        d.line([(9 * u, 20 * u), (15 * u, 20 * u)], fill=color, width=w)

    elif kind == "crop":
        d.line([(6 * u, 2 * u), (6 * u, 18 * u), (22 * u, 18 * u)], fill=color, width=w, joint="curve")
        d.line([(2 * u, 6 * u), (18 * u, 6 * u), (18 * u, 22 * u)], fill=color, width=w, joint="curve")

    elif kind == "line":
        d.line([(4 * u, 20 * u), (20 * u, 4 * u)], fill=color, width=int(w * 1.3))
        for x, y in ((4, 20), (20, 4)):
            r = w * 0.65
            d.ellipse([x * u - r, y * u - r, x * u + r, y * u + r], fill=color)

    elif kind == "fullpage":
        d.rounded_rectangle(box(5, 2, 19, 22), radius=2 * u, outline=color, width=w)
        d.line(box(12, 6.5, 12, 15.5), fill=color, width=w)
        d.polygon([(8.2 * u, 13 * u), (15.8 * u, 13 * u), (12 * u, 18 * u)], fill=color)

    elif kind == "compare":
        d.rounded_rectangle(box(2, 4, 11, 20), radius=1.5 * u, outline=color, width=w)
        d.rounded_rectangle(box(13, 4, 22, 20), radius=1.5 * u, outline=color, width=w)
        d.ellipse(box(15, 9, 20, 14), fill=accent)
        d.line(box(4.5, 9, 8.5, 9), fill=color, width=w)
        d.line(box(4.5, 13, 8.5, 13), fill=color, width=w)

    elif kind == "ocr":
        L = 5
        for (x, y, dx, dy) in [(3, 3, 1, 1), (21, 3, -1, 1), (3, 21, 1, -1), (21, 21, -1, -1)]:
            d.line([x * u, y * u, (x + dx * L) * u, y * u], fill=color, width=w)
            d.line([x * u, y * u, x * u, (y + dy * L) * u], fill=color, width=w)
        d.line(box(7.5, 9, 16.5, 9), fill=color, width=w)
        d.line(box(7.5, 12.5, 16.5, 12.5), fill=color, width=w)
        d.line(box(7.5, 16, 13, 16), fill=color, width=w)

    elif kind == "blur":
        for i, x in enumerate((5, 10, 15, 20)):
            for j, y in enumerate((5, 10, 15, 20)):
                r = (2.4 - 0.45 * abs(i - j)) * u
                if r > 0.3 * u:
                    d.ellipse([x * u - r, y * u - r, x * u + r, y * u + r], fill=color)

    elif kind == "folder":
        d.rounded_rectangle(box(2, 5, 11, 9), radius=1.5 * u, fill=color)
        d.rounded_rectangle(box(2, 7, 22, 19), radius=2 * u, fill=color)

    elif kind == "doc":
        d.rounded_rectangle(box(5, 2, 19, 22), radius=2 * u, outline=color, width=w)
        for y in (8, 12, 16):
            d.line(box(8.5, y, 15.5, y), fill=color, width=w)

    elif kind == "mic":
        d.rounded_rectangle(box(9, 2, 15, 14), radius=3 * u, fill=color)
        d.arc(box(5.5, 7, 18.5, 18), start=0, end=180, fill=color, width=w)
        d.line(box(12, 18, 12, 21.5), fill=color, width=w)

    elif kind == "speaker":
        d.polygon([(3 * u, 9 * u), (7 * u, 9 * u), (12 * u, 4.5 * u),
                   (12 * u, 19.5 * u), (7 * u, 15 * u), (3 * u, 15 * u)], fill=color)
        d.arc(box(10, 8, 17, 16), start=-60, end=60, fill=color, width=w)
        d.arc(box(10, 4.5, 21, 19.5), start=-60, end=60, fill=color, width=w)

    elif kind == "pause":
        d.rounded_rectangle(box(6, 5, 10, 19), radius=u, fill=color)
        d.rounded_rectangle(box(14, 5, 18, 19), radius=u, fill=color)

    elif kind == "play":
        d.polygon([(7 * u, 4.5 * u), (19.5 * u, 12 * u), (7 * u, 19.5 * u)], fill=color)

    elif kind == "stop":
        d.rounded_rectangle(box(6, 6, 18, 18), radius=2 * u, fill=color)

    elif kind == "dot":
        d.ellipse(box(4, 4, 20, 20), fill=color)

    elif kind == "logo":
        d.rounded_rectangle(box(1, 1, 23, 23), radius=6 * u, fill=ACCENT)
        d.rounded_rectangle(box(5, 8, 19, 18.5), radius=2.5 * u, fill="white")
        d.rounded_rectangle(box(8.5, 6, 13, 9), radius=1 * u, fill="white")
        d.ellipse(box(9, 10, 15, 16), fill=ACCENT)
        d.ellipse(box(10.8, 11.8, 13.2, 14.2), fill="white")
        d.ellipse(box(16, 9.3, 18, 11.3), fill=RED)

    return ImageTk.PhotoImage(_finish(img, size))


def icon(kind, color=TEXT, size=20, accent=RED):
    key = (kind, color, size, accent)
    if key not in ICONS:
        ICONS[key] = draw_icon(kind, color, max(8, min(256, S(size))), accent)
    return ICONS[key]


def switch_image(on, enabled=True):
    """Pill-shaped on/off switch."""
    key = ("switch", on, enabled)
    if key not in ICONS:
        w, h = S(34) * 4, S(18) * 4
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        track = (ACCENT if on else "#9AA3B2") if enabled else ("#B9BCF0" if on else "#D5D9E0")
        d.rounded_rectangle([0, 0, w - 1, h - 1], radius=h // 2, fill=track)
        pad = h // 7
        knob = h - 2 * pad
        x = w - pad - knob if on else pad
        d.ellipse([x, pad, x + knob, pad + knob], fill="white")
        ICONS[key] = ImageTk.PhotoImage(img.resize((w // 4, h // 4), Image.LANCZOS))
    return ICONS[key]


# ============================================================
# Styled Widgets
# ============================================================

class _LabelButton(tk.Label):
    """A button made from a Label. macOS draws tk.Button natively and
    ignores its colours, so the flat coloured buttons use this there."""

    def __init__(self, parent, command=None, **kw):
        kw.pop("relief", None)
        self._command = command
        super().__init__(parent, **kw)
        self.bind("<ButtonRelease-1>", self._release)

    def _release(self, event):
        inside = 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height()
        if inside and str(self["state"]) != "disabled":
            self.invoke()

    def invoke(self):
        if self._command is not None and str(self["state"]) != "disabled":
            return self._command()

    def configure(self, cnf=None, **kw):
        if "command" in kw:
            self._command = kw.pop("command")
        kw.pop("relief", None)
        return tk.Label.configure(self, cnf, **kw)

    config = configure


_BUTTON_BASE = _LabelButton if IS_MAC else tk.Button


class FlatButton(_BUTTON_BASE):
    """Flat button with hover colour. style: primary | secondary | danger | dark"""

    STYLES = {
        "primary":   dict(bg=ACCENT, fg="white", hover=ACCENT_HOVER, border=ACCENT),
        "secondary": dict(bg=CARD, fg=TEXT, hover="#EEF0F4", border=FIELD_BORDER),
        "danger":    dict(bg=RED, fg="white", hover="#D13438", border=RED),
        "dark":      dict(bg="#2A3040", fg="white", hover="#394157", border="#2A3040"),
    }

    def __init__(self, parent, text, command, style="secondary", icons=None, **kw):
        st = self.STYLES[style]
        self._bg, self._hover = st["bg"], st["hover"]
        self._icons = icons or {}
        options = dict(
            text=text, command=command, bg=st["bg"], fg=st["fg"],
            activebackground=st["hover"], activeforeground=st["fg"],
            disabledforeground=FAINT, relief="flat", bd=0, cursor="hand2",
            highlightthickness=1, highlightbackground=st["border"],
            highlightcolor=st["border"], font=F_BODY_BOLD,
            padx=S(12), pady=S(5),
        )
        if text in self._icons:
            options.update(image=self._icons[text], compound="left")
        options.update(kw)
        super().__init__(parent, **options)
        self.bind("<Enter>", lambda e: self._hover_on(True))
        self.bind("<Leave>", lambda e: self._hover_on(False))

    def _hover_on(self, on):
        if str(self["state"]) != "disabled":
            _BUTTON_BASE.configure(self, bg=self._hover if on else self._bg)

    def configure(self, cnf=None, **kw):
        # Swap the icon together with the text (e.g. Pause <-> Resume)
        if "text" in kw and kw["text"] in self._icons:
            kw["image"] = self._icons[kw["text"]]
        if kw.get("state") == "disabled":
            kw.setdefault("cursor", "arrow")
            kw.setdefault("bg", self._bg)
        elif kw.get("state") == "normal":
            kw.setdefault("cursor", "hand2")
        return _BUTTON_BASE.configure(self, cnf, **kw)

    config = configure


class Tile(tk.Frame):
    """Large clickable action card: icon, title and a short description."""

    def __init__(self, parent, image, title, subtitle, command, hover_border=ACCENT):
        super().__init__(parent, bg=CARD, highlightthickness=1,
                         highlightbackground=BORDER, cursor="hand2")
        self._command = command
        self._state = "normal"
        self._hover_border = hover_border

        self.icon_box = tk.Label(self, image=image, bg=ACCENT_SOFT if hover_border == ACCENT else RED_SOFT,
                                 width=S(38), height=S(38))
        self.icon_box.grid(row=0, column=0, rowspan=2, padx=(S(11), S(10)), pady=S(8))

        self.title = tk.Label(self, text=title, font=F_TILE, fg=TEXT, bg=CARD, anchor="w")
        self.title.grid(row=0, column=1, sticky="sw", pady=(S(8), 0), padx=(0, S(10)))
        self.sub = tk.Label(self, text=subtitle, font=F_TILE_SUB, fg=MUTED, bg=CARD, anchor="w",
                            justify="left")
        self.sub.grid(row=1, column=1, sticky="nw", pady=(S(1), S(8)), padx=(0, S(10)))

        for widget in (self, self.icon_box, self.title, self.sub):
            widget.bind("<Enter>", lambda e: self._hover(True))
            widget.bind("<Leave>", lambda e: self._hover(False))
            widget.bind("<Button-1>", self._click)

    def _hover(self, on):
        if self._state == "disabled":
            return
        bg = CARD_HOVER if on else CARD
        tk.Frame.configure(self, bg=bg,
                           highlightbackground=self._hover_border if on else BORDER)
        for widget in (self.title, self.sub):
            widget.configure(bg=bg)

    def _click(self, event=None):
        if self._state != "disabled" and self._command:
            self._hover(False)
            self._command()

    def configure(self, cnf=None, **kw):
        state = kw.pop("state", None)
        if state is not None:
            self._state = state
            disabled = state == "disabled"
            self.title.configure(fg=FAINT if disabled else TEXT)
            self.sub.configure(fg=FAINT if disabled else MUTED)
            for widget in (self, self.icon_box, self.title, self.sub):
                widget.configure(cursor="arrow" if disabled else "hand2")
            self._hover(False)
        if kw or cnf:
            return tk.Frame.configure(self, cnf, **kw)

    config = configure


class Toggle(tk.Frame):
    """On/off switch bound to a BooleanVar, with an icon and a label."""

    def __init__(self, parent, text, variable, image=None, bg=CARD):
        super().__init__(parent, bg=bg, cursor="hand2")
        self.var = variable
        self._state = "normal"
        self.switch = tk.Label(self, bg=bg)
        self.switch.pack(side="left")
        parts = [self.switch]
        if image is not None:
            self.img = tk.Label(self, image=image, bg=bg)
            self.img.pack(side="left", padx=(S(8), 0))
            parts.append(self.img)
        self.label = tk.Label(self, text=text, font=F_BODY, fg=TEXT, bg=bg)
        self.label.pack(side="left", padx=(S(5), 0))
        parts.append(self.label)
        for widget in [self] + parts:
            widget.bind("<Button-1>", self._toggle)
        variable.trace_add("write", lambda *a: self._draw())
        self._draw()

    def _toggle(self, event=None):
        if self._state != "disabled":
            self.var.set(not self.var.get())

    def _draw(self):
        enabled = self._state != "disabled"
        self.switch.configure(image=switch_image(bool(self.var.get()), enabled))
        self.label.configure(fg=TEXT if enabled else FAINT)

    def configure(self, cnf=None, **kw):
        state = kw.pop("state", None)
        if state is not None:
            self._state = state
            self._draw()
            for widget in self.winfo_children() + [self]:
                widget.configure(cursor="arrow" if state == "disabled" else "hand2")
        if kw or cnf:
            return tk.Frame.configure(self, cnf, **kw)

    config = configure


class Card(tk.Frame):
    """White section card with a clear heading and optional badge."""

    def __init__(self, parent, title, badge=None, required=False):
        super().__init__(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        head = tk.Frame(self, bg=CARD)
        head.pack(fill="x", padx=S(16), pady=(S(8), S(6)))
        tk.Frame(head, bg=ACCENT, width=S(4), height=S(16)).pack(side="left", padx=(0, S(8)))
        tk.Label(head, text=title.upper(), font=F_SECTION, fg=TEXT, bg=CARD).pack(side="left")
        if required:   # red asterisk = must be filled in
            tk.Label(head, text="*", font=font(12, "Segoe UI", "bold"),
                     fg=RED, bg=CARD).pack(side="left", padx=(S(3), 0))
        self.badge = tk.Label(head, font=F_SECTION, bg=ACCENT_SOFT, fg=ACCENT, padx=S(8), pady=S(1))
        if badge:
            self.set_badge(badge)
        # divider between heading and content
        tk.Frame(self, bg="#E3E6EC", height=1).pack(fill="x", padx=S(16))
        self.body = tk.Frame(self, bg=CARD)
        self.body.pack(fill="x", padx=S(16), pady=(S(8), S(11)))

    def set_badge(self, text, warn=False):
        self.badge.configure(text=text, bg=RED_SOFT if warn else ACCENT_SOFT,
                             fg=RED if warn else ACCENT)
        self.badge.pack(side="right")


def styled_entry(parent, variable, readonly=False):
    entry = tk.Entry(parent, textvariable=variable, font=F_BODY, fg=TEXT, bg=CARD,
                     relief="flat", bd=0, highlightthickness=S(2) if S(2) > 1 else 1,
                     highlightbackground=FIELD_BORDER, highlightcolor=ACCENT,
                     insertbackground=TEXT, readonlybackground=FIELD_READONLY,
                     disabledforeground=MUTED)
    if readonly:
        entry.configure(state="readonly")
    return entry


# Window icon (title bar / taskbar when running as a script)
try:
    root.iconphoto(True, icon("logo", size=32))
except Exception:
    pass


# Everything in the normal window lives in main_frame, so it can be
# swapped for the small control bar while recording.
main_frame = tk.Frame(root, bg=BG)
main_frame.pack()


# ============================================================
# Header
# ============================================================

header = tk.Frame(main_frame, bg=HEADER_BG)
header.pack(fill="x")

header_inner = tk.Frame(header, bg=HEADER_BG)
header_inner.pack(fill="x", padx=S(18), pady=S(10))

tk.Label(header_inner, image=icon("logo", size=34), bg=HEADER_BG).pack(side="left")

header_text = tk.Frame(header_inner, bg=HEADER_BG)
header_text.pack(side="left", padx=(S(12), 0))
tk.Label(header_text, text="Screen Capture Utility", font=F_TITLE,
         fg="white", bg=HEADER_BG).pack(anchor="w")
header_subtitle = tk.Label(header_text,
                           text="Screenshots, recordings with sound, and Word / Excel / PowerPoint reports",
                           font=F_SUB, fg=HEADER_SUB, bg=HEADER_BG)
header_subtitle.pack(anchor="w")

# thin accent line under the header
tk.Frame(main_frame, bg=ACCENT, height=S(3)).pack(fill="x")

content = tk.Frame(main_frame, bg=BG)
content.pack(fill="both", padx=S(16), pady=(S(12), S(4)))


# ============================================================
# Save Location Card
# ============================================================

folder_card = Card(content, "Save location", required=True)
folder_card.pack(fill="x", pady=(0, S(9)))

entry_row = tk.Frame(folder_card.body, bg=CARD)
entry_row.pack(fill="x")

tk.Label(entry_row, image=icon("folder", ACCENT, 18), bg=CARD).pack(side="left", padx=(0, S(8)))

folder_entry = styled_entry(entry_row, folder_var)
folder_entry.configure(width=40)
folder_entry.pack(side="left", fill="x", expand=True, ipady=S(5), ipadx=S(4))

browse_button = FlatButton(entry_row, "Browse", browse_folder, style="primary")
browse_button.pack(side="left", padx=(S(8), 0))


# ============================================================
# Document Card (Word / Excel / PowerPoint)
# ============================================================

word_frame = Card(content, "Document  (Word / Excel / PowerPoint)", badge="Optional")
word_frame.pack(fill="x", pady=(0, S(9)))

word_hint_label = tk.Label(
    word_frame.body,
    text="Screenshots are added to this Word, Excel or PowerPoint file in the order you take them.",
    font=F_TILE_SUB, fg=MUTED, bg=CARD)
word_hint_label.pack(anchor="w", pady=(0, S(6)))

doc_row = tk.Frame(word_frame.body, bg=CARD)
doc_row.pack(fill="x")

tk.Label(doc_row, image=icon("doc", ACCENT, 18), bg=CARD).pack(side="left", padx=(0, S(8)))

doc_entry = styled_entry(doc_row, doc_var, readonly=True)
doc_entry.pack(side="left", fill="x", expand=True, ipady=S(5), ipadx=S(4))

select_doc_button = FlatButton(doc_row, "Select", select_word_document)
select_doc_button.pack(side="left", padx=(S(8), 0))

new_doc_button = FlatButton(doc_row, "New...", new_document_menu)
new_doc_button.pack(side="left", padx=(S(6), 0))

clear_doc_button = FlatButton(doc_row, "Clear", clear_word_document)
clear_doc_button.pack(side="left", padx=(S(6), 0))

doc_options_row = tk.Frame(word_frame.body, bg=CARD)
doc_options_row.pack(fill="x", pady=(S(8), 0))

keep_png_check = Toggle(doc_options_row, "Also save separate .png files", keep_png_var)
keep_png_check.pack(side="left")

open_doc_button = FlatButton(doc_options_row, "Open Document", open_word_document)
open_doc_button.pack(side="right")

save_pending_button = FlatButton(doc_options_row, "Save to Word now",
                                 lambda: flush_pending(interactive=True))
save_pending_button.pack(side="right", padx=(0, S(6)))

if not OFFICE_AVAILABLE:
    for widget in (select_doc_button, new_doc_button):
        widget.config(state="disabled")
    word_frame.set_badge("Unavailable - see console", warn=True)


# ============================================================
# Screenshot Card
# ============================================================

shot_card = Card(content, "Screenshot & text", badge="PNG / Office / Text")
shot_card.pack(fill="x", pady=(0, S(9)))

shot_tiles = tk.Frame(shot_card.body, bg=CARD)
shot_tiles.pack(fill="x")
# keeps the window as wide as before the tiles moved to two rows
tk.Frame(shot_card.body, bg=CARD, width=S(700), height=1).pack()
shot_tiles.columnconfigure((0, 1, 2, 3, 4, 5), weight=1, uniform="tiles")

capture_button = Tile(shot_tiles, icon("screen", ACCENT, 24), "Capture Screen",
                      f"Entire desktop\n{KEY}+Shift+F", capture_screen)
capture_button.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=(0, S(5)))

region_button = Tile(shot_tiles, icon("region", ACCENT, 24), "Capture Region",
                     f"Drag to select\n{KEY}+Shift+R", select_region)
region_button.grid(row=0, column=2, columnspan=2, sticky="nsew", padx=(S(5), S(5)))

fullpage_button = Tile(shot_tiles, icon("fullpage", ACCENT, 24), "Full Page",
                       f"Scrolls long pages\n{KEY}+Shift+S", full_page_capture)
fullpage_button.grid(row=0, column=4, columnspan=2, sticky="nsew", padx=(S(5), 0))

ocr_button = Tile(shot_tiles, icon("ocr", ACCENT, 24), "Extract Text",
                  f"Copy text (OCR)\n{KEY}+Shift+T", extract_text_region)
ocr_button.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=(0, S(5)), pady=(S(8), 0))

compare_button = Tile(shot_tiles, icon("compare", ACCENT, 24), "Compare Images",
                      f"Highlight differences\n{KEY}+Shift+C", open_compare_window)
compare_button.grid(row=1, column=3, columnspan=3, sticky="nsew", padx=(S(5), 0), pady=(S(8), 0))

if not OCR_AVAILABLE:
    ocr_button.sub.configure(text="Needs Windows OCR\n(see console)")



# ============================================================
# Recording Card
# ============================================================

record_frame = Card(content, "Screen recording", badge="MP4")
record_frame.pack(fill="x", pady=(0, S(9)))

record_tiles = tk.Frame(record_frame.body, bg=CARD)
record_tiles.pack(fill="x")
record_tiles.columnconfigure((0, 1), weight=1, uniform="tiles")

record_button = Tile(record_tiles, icon("record", TEXT, 24), "Record Screen",
                     f"Full screen  ·  {KEY}+Shift+V", toggle_screen_recording,
                     hover_border=RED)
record_button.grid(row=0, column=0, sticky="nsew", padx=(0, S(6)))

record_region_button = Tile(record_tiles, icon("record_region", TEXT, 24), "Record Region",
                            "Drag to select an area", record_region,
                            hover_border=RED)
record_region_button.grid(row=0, column=1, sticky="nsew", padx=(S(6), 0))

# Sound options
mic_var = tk.BooleanVar(value=AUDIO_AVAILABLE)
system_audio_var = tk.BooleanVar(value=SYSTEM_AUDIO_AVAILABLE)

audio_row = tk.Frame(record_frame.body, bg=CARD)
audio_row.pack(fill="x", pady=(S(8), 0))

tk.Label(audio_row, text="Sound", font=F_BODY_BOLD, fg=TEXT, bg=CARD).pack(side="left", padx=(0, S(12)))

mic_check = Toggle(audio_row, "Microphone", mic_var, image=icon("mic", MUTED, 16))
mic_check.pack(side="left", padx=(0, S(16)))

system_audio_check = Toggle(audio_row, "Computer sound", system_audio_var,
                            image=icon("speaker", MUTED, 16))
system_audio_check.pack(side="left")

if not AUDIO_AVAILABLE:
    mic_check.config(state="disabled")
    system_audio_check.config(state="disabled")
    tk.Label(audio_row, text="unavailable", font=F_TILE_SUB, fg=FAINT, bg=CARD).pack(side="left", padx=(S(10), 0))
elif not SYSTEM_AUDIO_AVAILABLE:
    system_audio_check.config(state="disabled")
    tk.Label(audio_row, text="not on Mac" if IS_MAC else "unavailable", font=F_TILE_SUB,
             fg=FAINT, bg=CARD).pack(side="left", padx=(S(10), 0))

if not RECORDING_AVAILABLE:
    record_button.config(state="disabled")
    record_region_button.config(state="disabled")
    record_frame.set_badge("Unavailable - see console", warn=True)


# ============================================================
# Status Bar
# ============================================================

status_bar = tk.Frame(main_frame, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
status_bar.pack(fill="x", side="bottom")

status_dot = tk.Label(status_bar, bg=CARD, fg=RED, bd=0, padx=0, pady=0,
                      font=font(13, "Segoe UI", "bold"))
status_dot.pack(side="left", anchor="n", padx=(S(16), S(8)), pady=(S(9), 0))

status_label = tk.Label(
    status_bar,
    textvariable=status_var,
    font=F_BODY,
    fg=TEXT,
    bg=CARD,
    height=2,
    justify="left",
    anchor="w"
)
status_label.pack(side="left", fill="x", expand=True, padx=(0, S(16)), pady=S(4))


def update_status_dot(*args):
    """Colour the status dot by message type: green = done,
    amber = busy or needs attention, red = problem, indigo = info."""
    text = status_var.get().strip().lower()
    if text.startswith(("choose a folder", "no save folder")):
        # the save folder is mandatory - mark it with a red asterisk
        status_dot.configure(image="", text="*")
        status_dot.pack_configure(anchor="center", pady=0)
        return
    status_dot.configure(text="")
    status_dot.pack_configure(anchor="n", pady=(S(9), 0))
    if (text.startswith(("recording failed", "could not", "error"))
            or "cancelled" in text or "too small" in text):
        color = RED
    elif text.startswith(("saving recording", "finishing")) or "waiting" in text:
        color = AMBER
    elif text.startswith(("saved:", "added ")):
        color = GREEN
    else:
        color = ACCENT
    status_dot.configure(image=icon("dot", color, 10))


status_var.trace_add("write", update_status_dot)
update_status_dot()


def fit_status_wrap(event=None):
    status_label.config(wraplength=max(S(200), root.winfo_width() - S(70)))


root.after(100, fit_status_wrap)


def fit_to_screen():
    """On small or highly scaled screens, drop two decorative lines so the
    whole window fits above the taskbar."""
    root.update_idletasks()
    available = screen_work_area()[3] - S(40)   # minus the title bar
    if root.winfo_reqheight() > available:
        word_hint_label.pack_forget()
        header_subtitle.pack_forget()


fit_to_screen()


def center_window():
    """Show the window in the middle of the screen (monitor, laptop,
    any size or display scaling)."""
    root.update_idletasks()
    left, top, width, height = screen_work_area()

    # Guess the title bar / border size first, show the window invisibly,
    # then correct the position using the real frame size.
    frame_w, frame_h = S(16), S(39)

    def place():
        win_w = root.winfo_reqwidth() + frame_w
        win_h = root.winfo_reqheight() + frame_h
        x = left + max(0, (width - win_w) // 2)
        y = top + max(0, (height - win_h) // 2)
        root.geometry(f"+{x}+{y}")

    try:
        root.attributes("-alpha", 0.0)
    except Exception:
        pass

    place()
    root.deiconify()
    root.update()

    # Real frame size = outer window minus the inner (client) area
    outer_x, outer_y = root.winfo_x(), root.winfo_y()
    measured_w = 2 * (root.winfo_rootx() - outer_x)
    measured_h = (root.winfo_rooty() - outer_y) + (root.winfo_rootx() - outer_x)
    if measured_h > 0:
        frame_w, frame_h = measured_w, measured_h
        place()
        root.update_idletasks()

    try:
        root.attributes("-alpha", 1.0)
    except Exception:
        pass
    root.lift()
    root.focus_force()


# ============================================================
# Recording Control Bar (shown only while recording)
# ============================================================

rec_bar = tk.Frame(root, bg=BAR_BG, padx=S(12), pady=S(10))

rec_dot = tk.Canvas(rec_bar, width=S(14), height=S(14), bg=BAR_BG,
                    highlightthickness=0)
rec_dot_id = rec_dot.create_oval(S(2), S(2), S(13), S(13), fill=DOT_ON, outline="")
rec_dot.pack(side="left", padx=(0, S(8)))

rec_time_label = tk.Label(rec_bar, text="00:00", fg="white", bg=BAR_BG,
                          font=F_TIMER, width=8, anchor="w")
rec_time_label.pack(side="left", padx=(0, S(10)))

rec_pause_button = FlatButton(
    rec_bar, "Pause", toggle_pause, style="dark",
    icons={"Pause": icon("pause", "white", 14), "Resume": icon("play", "white", 14)},
    width=S(88), padx=S(10), pady=S(5),
)
rec_pause_button.pack(side="left", padx=(0, S(8)))

rec_stop_button = FlatButton(
    rec_bar, "Stop", stop_recording, style="danger",
    icons={"Stop": icon("stop", "white", 14)},
    width=S(76), padx=S(10), pady=S(5),
)
rec_stop_button.pack(side="left")


# Full-page capture progress bar (shown only while scrolling/capturing)
scroll_bar = tk.Frame(root, bg=BAR_BG, padx=S(12), pady=S(10))
tk.Label(scroll_bar, image=icon("fullpage", "white", 18), bg=BAR_BG).pack(side="left", padx=(0, S(8)))
scroll_label = tk.Label(scroll_bar, text="Scrolling to the top...", fg="white", bg=BAR_BG,
                        font=F_BODY_BOLD, width=38, anchor="w")
scroll_label.pack(side="left", padx=(0, S(8)))
tk.Label(scroll_bar, text="Esc also stops", fg=HEADER_SUB, bg=BAR_BG,
         font=F_TILE_SUB).pack(side="left", padx=(0, S(10)))
scroll_stop_button = FlatButton(scroll_bar, "Stop & Save", lambda: None, style="danger",
                                icons={"Stop & Save": icon("stop", "white", 14)},
                                padx=S(10), pady=S(5))
scroll_stop_button.pack(side="left")

update_word_controls()


# ============================================================
# Keyboard Shortcuts (work while this window has focus)
# ============================================================

root.bind("<Control-Shift-F>",
          lambda event: capture_screen() if recorder is None else None)
root.bind("<Control-Shift-R>",
          lambda event: select_region() if recorder is None else None)
root.bind("<Control-Shift-S>",
          lambda event: full_page_capture() if recorder is None else None)
root.bind("<Control-Shift-C>", lambda event: open_compare_window())
root.bind("<Control-Shift-T>",
          lambda event: extract_text_region() if recorder is None else None)
root.bind("<Control-Shift-V>", lambda event: toggle_screen_recording())
root.bind("<Control-Shift-P>", lambda event: toggle_pause())
mirror_control_keys(root)                    # Cmd+Shift+... on a Mac


# ============================================================
# Start Application
# ============================================================

startup_step("window built")
center_window()
_started.set()
faulthandler.cancel_dump_traceback_later()
startup_step(f"window shown: {root.winfo_width()} x {root.winfo_height()}, {len(ICONS)} icons")

if IS_MAC:
    # Quit from the Dock, the app menu or Cmd+Q goes through the same checks
    # as closing the window (recording still saving, screenshots pending...)
    root.createcommand("tk::mac::Quit", on_close)
    try:      # bring the app to the front (opened from the Finder it may start behind)
        mac["AppKit"].NSApplication.sharedApplication().activateIgnoringOtherApps_(True)
    except Exception:
        pass


def check_mac_permissions():
    """macOS: ask once for Screen Recording permission, without which
    screenshots show only the desktop background."""
    if IS_MAC and not mac_screen_capture_allowed(ask=True):
        status_var.set("Allow Screen Recording for Screen Capture Utility in System Settings > "
                       "Privacy & Security, then quit and reopen the app.")


if SELFTEST:
    def _selftest_done():
        print(f"SELFTEST OK: window {root.winfo_width()} x {root.winfo_height()}, "
              f"peak memory {peak_memory_mb():.0f} MB, start-up {time.time() - _T0:.1f} s")
        root.destroy()
    root.after(1500, _selftest_done)
else:
    root.after(800, check_mac_permissions)

print("Screen Capture Utility started.")
print("Choose a save folder in the window (type a path or click Browse...).")
print("Optionally select a Word, Excel or PowerPoint file to collect screenshots in.")

root.mainloop()