import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageGrab, ImageOps, ImageFilter, ImageDraw, ImageChops
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass
from io import BytesIO
import threading
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


# ============================================================
# Optional packages (installed automatically if missing)
# ============================================================
# Screen recording uses ffmpeg (imageio-ffmpeg) and, for sound,
# PyAudioWPatch. OpenCV + NumPy are a fallback for video without sound.
# Adding screenshots to a Word document needs python-docx.
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

RECORDING_AVAILABLE = FFMPEG_EXE is not None or (np is not None and cv2 is not None)
AUDIO_AVAILABLE = FFMPEG_EXE is not None and pyaudio is not None
DOCX_AVAILABLE = docx is not None

if DOCX_AVAILABLE:
    from docx import Document
    from docx.shared import Emu, Inches, Pt

if not RECORDING_AVAILABLE:
    print("Screen recording is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install imageio-ffmpeg')

if not AUDIO_AVAILABLE:
    print("Audio recording is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install imageio-ffmpeg PyAudioWPatch')

if not DOCX_AVAILABLE:
    print("Saving to Word documents is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install python-docx')

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

if not OCR_AVAILABLE:
    print("Text extraction (OCR) is disabled (needs Windows 10/11 and PyWinRT).")

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
# Word Document Output
# ============================================================
# When a Word document is selected, every screenshot is appended to the
# end of that document (with a caption showing the capture time) and the
# document is saved straight away. Screenshots are always added in the
# order they were captured, so the earliest one is at the top.
#
# If the document is open in Word, Windows locks the file and it can't be
# written. In that case captures wait in a queue (in capture order) and
# are written automatically as soon as the document is closed in Word.

@dataclass
class Capture:
    taken_at: datetime
    kind: str          # "Full screen" or "Region"
    png: bytes         # the image, PNG-encoded
    size: tuple        # (width, height) in pixels


word_doc_path = None     # Path of the selected .docx, or None
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


def is_file_locked(path):
    """True if another program (usually Word) has the file open."""
    if not path.exists():
        return False
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def flush_pending(interactive=False):
    """
    Write any waiting captures to the Word document.
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
        append_captures_to_docx(word_doc_path, pending_captures)

    except PermissionError:
        report_document_locked()
        return False

    except Exception as e:
        print(f"Error updating Word document: {e}")
        status_var.set(f"Could not update {name} (see console)")
        update_word_controls()
        if interactive:
            messagebox.showerror(
                "Could not update Word document",
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
        f"{word_doc_path.name} is open in Word - {n} {noun} waiting. "
        f"Close the document and they will be added automatically."
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
    """Route a new screenshot to the Word document or to a .png file."""
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
        "Close that document in Word so they can be saved, then try again.",
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


def select_word_document():
    if not DOCX_AVAILABLE or not can_change_document():
        return

    path = filedialog.askopenfilename(
        parent=root,
        title="Select the Word document to add screenshots to",
        initialdir=document_start_folder(),
        filetypes=[("Word documents", "*.docx")]
    )
    if not path:
        return

    path = Path(path)

    if path.name.startswith("~$"):
        messagebox.showerror(
            "Not a document",
            "That is a temporary lock file Word creates while a document "
            "is open. Please select the real .docx file.",
            parent=root
        )
        return

    # Make sure python-docx can actually read it
    try:
        Document(str(path))
    except Exception as e:
        messagebox.showerror(
            "Can't use this document",
            f"{path}\n\n{e}\n\n"
            "Only Word .docx files are supported (not older .doc files). "
            "Open it in Word and use Save As > Word Document (*.docx).",
            parent=root
        )
        return

    set_word_document(path)


def new_word_document():
    if not DOCX_AVAILABLE or not can_change_document():
        return

    path = filedialog.asksaveasfilename(
        parent=root,
        title="Create a new Word document for screenshots",
        initialdir=document_start_folder(),
        initialfile=f"Screen Captures {datetime.now():%Y-%m-%d}.docx",
        defaultextension=".docx",
        filetypes=[("Word documents", "*.docx")]
    )
    if not path:
        return

    path = Path(path)

    try:
        doc = Document()   # empty document - images only, no heading
        doc.save(str(path))
    except Exception as e:
        messagebox.showerror(
            "Could not create document",
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
    print(f"Screenshots will be added to Word document: {path}")
    update_word_controls()


def clear_word_document():
    global word_doc_path

    if word_doc_path is None or not can_change_document():
        return

    word_doc_path = None
    doc_var.set("")
    status_var.set("No Word document - screenshots will be saved as .png files")
    update_word_controls()


def open_word_document():
    if word_doc_path is None:
        return
    try:
        os.startfile(str(word_doc_path))  # Windows: open in Word
    except Exception as e:
        messagebox.showerror("Could not open document", str(e), parent=root)


def update_word_controls():
    """Enable/disable the Word buttons to match the current state."""
    has_doc = word_doc_path is not None
    state = "normal" if has_doc else "disabled"

    clear_doc_button.config(state=state)
    open_doc_button.config(state=state)
    keep_png_check.config(state=state)

    n = len(pending_captures)
    if n:
        save_pending_button.config(text=f"Save to Word now ({n} waiting)",
                                   state="normal")
    else:
        save_pending_button.config(text="Save to Word now", state="disabled")


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


def _prepare_for_ocr(image, max_dimension):
    """RGBA image sized for best OCR results: small captures are enlarged
    (Windows OCR reads small text better that way), huge ones reduced to
    the engine's limit, and a white margin is added around the text."""
    image = image.convert("RGB")
    margin = 12
    image = ImageOps.expand(image, border=margin, fill=(255, 255, 255))

    longest = max(image.size)
    limit = max_dimension or 4000
    scale = 1.0
    if longest < 1600:
        scale = min(2.0, 1600 / longest)
    scale = min(scale, limit / longest)
    if abs(scale - 1.0) > 0.01:
        image = image.resize((max(1, int(image.width * scale)),
                              max(1, int(image.height * scale))),
                             Image.LANCZOS)
    return image.convert("RGBA")


def _lines_to_text(lines):
    """Join OCR lines, adding a blank line where there's a clear vertical
    gap (so paragraphs stay separated)."""
    rows = []
    for line in lines:
        text = line.text
        try:
            rects = [w.bounding_rect for w in line.words]
            top = min(r.y for r in rects)
            bottom = max(r.y + r.height for r in rects)
        except Exception:
            top = bottom = None
        rows.append((text, top, bottom))

    heights = [b - t for _, t, b in rows if t is not None and b is not None and b > t]
    typical = sorted(heights)[len(heights) // 2] if heights else None

    output = []
    previous_bottom = None
    for text, top, bottom in rows:
        if (typical and previous_bottom is not None and top is not None
                and top - previous_bottom > typical * 1.2):
            output.append("")
        output.append(text)
        if bottom is not None:
            previous_bottom = bottom
    return "\n".join(output).strip()


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

    writer = DataWriter()
    writer.write_bytes(prepared.tobytes())
    bitmap = SoftwareBitmap.create_copy_from_buffer(
        writer.detach_buffer(), BitmapPixelFormat.RGBA8,
        prepared.width, prepared.height)

    async def recognize():
        return await engine.recognize_async(bitmap)

    result = asyncio.run(recognize())

    return {
        "text": _lines_to_text(list(result.lines)),
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
            "Text extraction uses the OCR engine built into Windows 10/11.\n\n"
            "Install its Python bindings with:\n"
            f'  "{sys.executable}" -m pip install {" ".join(OCR_PACKAGES)}',
            parent=root)
        return
    start_region_selection(ocr_selected_region)


def ocr_selected_region(x1, y1, x2, y2):
    try:
        image = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
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

        open_button = FlatButton(lang_row, "Open image...", self.open_image)
        open_button.pack(side="right")

        # --- text area -----------------------------------------------
        text_frame = tk.Frame(body, bg=CARD, highlightthickness=S(2) if S(2) > 1 else 1,
                              highlightbackground=FIELD_BORDER)
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(text_frame, wrap="word", font=("Segoe UI", 11),
                            fg=TEXT, bg=CARD, relief="flat", bd=0,
                            padx=S(12), pady=S(10), undo=True,
                            width=72, height=18, insertbackground=TEXT,
                            selectbackground="#C7CBF9", selectforeground=TEXT)
        scroll = tk.Scrollbar(text_frame, command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
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
        self.future = OCR_WORKER.submit(run_windows_ocr, self.image, self.language_tag)
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
# After a screenshot is taken (when "Review & blur before saving" is on),
# the Blur editor opens: drag boxes over anything sensitive and each box
# is blurred with a Gaussian blur. Saving sends the result on to the Word
# document or .png file as usual. Existing images can be blurred too.

BLUR_DEFAULT_RADIUS = 12


def blur_region(image, box, radius):
    """Blur one rectangular area of the image in place."""
    blurred = image.crop(box)
    blurred = blurred.filter(ImageFilter.GaussianBlur(radius=radius))
    image.paste(blurred, box)


def deliver_capture(image, kind, taken_at):
    """Save a new screenshot - via the Blur editor if review is switched on."""
    if blur_review_var.get():
        # open after the main window has been restored, so the editor is on top
        root.after(50, lambda: BlurEditor(
            image,
            on_save=lambda result: annotate_then_save(result, kind, taken_at),
            title=f"{kind} capture",
        ))
    else:
        annotate_then_save(image, kind, taken_at)


def blur_existing_image():
    """Pick a saved image, blur parts of it and save a copy."""
    folder = save_folder or (word_doc_path.parent if word_doc_path else Path.home())
    path = filedialog.askopenfilename(
        parent=root,
        title="Choose an image to blur",
        initialdir=str(folder),
        filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"),
                   ("All files", "*.*")])
    if not path:
        return
    path = Path(path)
    try:
        with Image.open(path) as img:
            image = img.convert("RGB")
    except Exception as e:
        messagebox.showerror("Could not open image", str(e), parent=root)
        return

    def save_copy(result):
        target = filedialog.asksaveasfilename(
            parent=root,
            title="Save blurred image",
            initialdir=str(path.parent),
            initialfile=f"{path.stem}_blurred.png",
            defaultextension=".png",
            filetypes=[("PNG image", "*.png"), ("JPEG image", "*.jpg *.jpeg")])
        if not target:
            return False
        try:
            result.save(target)
            status_var.set(f"Saved: {Path(target).name}")
            return True
        except Exception as e:
            messagebox.showerror("Could not save", str(e), parent=root)
            return False

    BlurEditor(image, on_save=save_copy, title=path.name, save_label="Save copy")


class BlurEditor:
    """Window where the user drags boxes over sensitive areas to blur them."""

    def __init__(self, image, on_save, title="", save_label="Save"):
        self.original = image.convert("RGB")
        self.on_save = on_save
        self.boxes = []                 # blurred areas, in image pixels
        self.drag_start = None
        self.drag_rect = None
        self.photo = None

        self.win = tk.Toplevel(root)
        self.win.title("Blur sensitive information")
        self.win.configure(bg=BG)
        self.win.transient(root)
        self.win.protocol("WM_DELETE_WINDOW", self.cancel)

        # --- header ---------------------------------------------------
        head = tk.Frame(self.win, bg=HEADER_BG)
        head.pack(fill="x")
        inner = tk.Frame(head, bg=HEADER_BG)
        inner.pack(fill="x", padx=S(20), pady=S(12))
        tk.Label(inner, image=icon("blur", "white", 22), bg=HEADER_BG).pack(side="left")
        titles = tk.Frame(inner, bg=HEADER_BG)
        titles.pack(side="left", padx=(S(10), 0))
        tk.Label(titles, text="Blur sensitive information", font=F_TILE,
                 fg="white", bg=HEADER_BG).pack(anchor="w")
        tk.Label(titles,
                 text="Drag over anything you want to hide.  Ctrl+Z undoes the last box.",
                 font=F_TILE_SUB, fg=HEADER_SUB, bg=HEADER_BG).pack(anchor="w")
        self.info = tk.Label(inner, text="", font=F_TILE_SUB, fg=HEADER_SUB, bg=HEADER_BG)
        self.info.pack(side="right")
        tk.Frame(self.win, bg=ACCENT, height=S(3)).pack(fill="x")

        body = tk.Frame(self.win, bg=BG)
        body.pack(fill="both", expand=True, padx=S(16), pady=S(12))

        # --- image canvas (scaled to fit the screen) -------------------
        left, top, width, height = screen_work_area()
        max_w, max_h = int(width * 0.82), int(height * 0.62)
        iw, ih = self.original.size
        width_scale = min(1.0, max_w / iw)
        # Tall images (full-page captures) are shown at a readable width
        # with a scrollbar instead of being squeezed to fit the screen
        self.tall = ih * width_scale > max_h * 1.3
        self.scale = width_scale if self.tall else min(1.0, max_w / iw, max_h / ih)
        self.view_w = max(1, int(iw * self.scale))
        self.view_h = max(1, int(ih * self.scale))

        frame = tk.Frame(body, bg=CARD, highlightthickness=1, highlightbackground=FIELD_BORDER)
        frame.pack()
        self.canvas = tk.Canvas(frame, width=self.view_w, height=min(self.view_h, max_h),
                                bg="#2A2F3A", highlightthickness=0, cursor="crosshair",
                                scrollregion=(0, 0, self.view_w, self.view_h))
        if self.tall:
            scroll = tk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=scroll.set)
            scroll.pack(side="right", fill="y")
            self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(
                int(-e.delta / 120) * 3, "units"))
            self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
            self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(3, "units"))
        self.canvas.pack(side="left")
        self.image_item = self.canvas.create_image(0, 0, anchor="nw")

        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._drag)
        self.canvas.bind("<ButtonRelease-1>", self._release)

        # --- controls ------------------------------------------------
        controls = tk.Frame(body, bg=BG)
        controls.pack(fill="x", pady=(S(12), 0))

        tk.Label(controls, text="Blur strength", font=F_BODY_BOLD, fg=TEXT,
                 bg=BG).pack(side="left")
        self.radius_var = tk.IntVar(value=BLUR_DEFAULT_RADIUS)
        style = ttk.Style(self.win)
        style.configure("Blur.Horizontal.TScale", background=BG)
        self.slider = ttk.Scale(controls, from_=4, to=40, orient="horizontal",
                                length=S(170), style="Blur.Horizontal.TScale",
                                value=BLUR_DEFAULT_RADIUS,
                                command=lambda v: self._slider_moved(v))
        self.slider.pack(side="left", padx=(S(8), S(6)))
        # re-blur once when the slider is released (smooth while dragging)
        self.slider.bind("<ButtonRelease-1>", lambda e: self._strength_changed())
        self.radius_label = tk.Label(controls, text=str(BLUR_DEFAULT_RADIUS),
                                     font=F_BODY, fg=MUTED, bg=BG, width=3, anchor="w")
        self.radius_label.pack(side="left")

        self.undo_button = FlatButton(controls, "Undo", self.undo)
        self.undo_button.pack(side="left", padx=(S(12), 0))
        self.clear_button = FlatButton(controls, "Clear all", self.clear)
        self.clear_button.pack(side="left", padx=(S(6), 0))

        FlatButton(controls, save_label, self.save, style="primary").pack(side="right")
        FlatButton(controls, "Cancel", self.cancel).pack(side="right", padx=(0, S(6)))

        self.win.bind("<Control-z>", lambda e: self.undo())
        self.win.bind("<Control-Z>", lambda e: self.undo())
        self.win.bind("<Return>", lambda e: self.save())
        self.win.bind("<Escape>", lambda e: self.cancel())

        self.working = self.original.copy()
        self.redraw()

        self.win.update_idletasks()
        w = self.win.winfo_reqwidth()
        h = self.win.winfo_reqheight() + S(39)
        self.win.geometry(f"+{left + max(0, (width - w) // 2)}+{top + max(0, (height - h) // 2)}")
        self.win.lift()
        self.win.focus_force()
        self.win.grab_set()   # finish (save/cancel) before doing anything else

    # ---- drawing ----------------------------------------------------

    def rebuild(self):
        """Re-apply every blur box to a fresh copy of the original."""
        self.working = self.original.copy()
        radius = self.radius_var.get()
        for box in self.boxes:
            blur_region(self.working, box, radius)

    def redraw(self):
        view = self.working
        if self.scale < 1.0:
            view = view.resize((self.view_w, self.view_h),
                               Image.BILINEAR if self.tall else Image.LANCZOS)
        self.photo = ImageTk.PhotoImage(view)
        self.canvas.itemconfigure(self.image_item, image=self.photo)

        # thin outline around each blurred area (not saved in the image)
        self.canvas.delete("box")
        for (x0, y0, x1, y1) in self.boxes:
            self.canvas.create_rectangle(
                x0 * self.scale, y0 * self.scale, x1 * self.scale, y1 * self.scale,
                outline=ACCENT, width=1, dash=(4, 3), tags="box")

        n = len(self.boxes)
        self.info.configure(text=f"{n} area{'s' if n != 1 else ''} blurred")
        state = "normal" if n else "disabled"
        self.undo_button.config(state=state)
        self.clear_button.config(state=state)

    # ---- mouse ------------------------------------------------------

    def _clamp(self, x, y):
        return (min(max(x, 0), self.view_w), min(max(y, 0), self.view_h))

    def _point(self, event):
        """Mouse position in image-view coordinates (accounts for scrolling)."""
        return self._clamp(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))

    def _press(self, event):
        self.drag_start = self._point(event)
        self.drag_rect = self.canvas.create_rectangle(
            *self.drag_start, *self.drag_start, outline=RED, width=2, dash=(5, 3))

    def _drag(self, event):
        if self.drag_rect is not None:
            # scroll while dragging near the top/bottom edge of a tall image
            if self.tall:
                if event.y < S(24):
                    self.canvas.yview_scroll(-1, "units")
                elif event.y > self.canvas.winfo_height() - S(24):
                    self.canvas.yview_scroll(1, "units")
            x, y = self._point(event)
            self.canvas.coords(self.drag_rect, *self.drag_start, x, y)

    def _release(self, event):
        if self.drag_rect is None:
            return
        self.canvas.delete(self.drag_rect)
        self.drag_rect = None
        x0, y0 = self.drag_start
        x1, y1 = self._point(event)

        # canvas -> image coordinates
        iw, ih = self.original.size
        box = (int(min(x0, x1) / self.scale), int(min(y0, y1) / self.scale),
               int(round(max(x0, x1) / self.scale)), int(round(max(y0, y1) / self.scale)))
        box = (max(0, box[0]), max(0, box[1]), min(iw, box[2]), min(ih, box[3]))
        if box[2] - box[0] < 3 or box[3] - box[1] < 3:
            return

        self.boxes.append(box)
        blur_region(self.working, box, self.radius_var.get())
        self.redraw()

    # ---- controls ---------------------------------------------------

    def _slider_moved(self, value):
        radius = int(round(float(value)))
        if radius != self.radius_var.get():
            self.radius_var.set(radius)
            self.radius_label.configure(text=str(radius))
            # live preview for normal-size screenshots
            if self.boxes and self.original.width * self.original.height <= 2_500_000:
                self.rebuild()
                self.redraw()

    def _strength_changed(self):
        self.radius_label.configure(text=str(self.radius_var.get()))
        if self.boxes:
            self.rebuild()
            self.redraw()

    def undo(self):
        if self.boxes:
            self.boxes.pop()
            self.rebuild()
            self.redraw()

    def clear(self):
        if self.boxes:
            self.boxes.clear()
            self.rebuild()
            self.redraw()

    def save(self):
        result = self.working.copy()
        self.win.grab_release()
        if self.on_save(result) is False:     # e.g. user cancelled "Save as"
            self.win.grab_set()
            return
        n = len(self.boxes)
        print(f"Saved with {n} blurred area(s)")
        self.win.destroy()

    def cancel(self):
        if self.boxes and not messagebox.askyesno(
                "Discard screenshot?",
                "Close without saving? The screenshot will be discarded.",
                parent=self.win):
            return
        self.win.grab_release()
        self.win.destroy()
        status_var.set("Screenshot discarded")


# ============================================================
# Review & Annotate window (opens after every capture)
# ============================================================
# After each screenshot (Capture Screen, Capture Region, Full Page) the
# captured image is shown in the Review window - after the Blur window, if
# "Review & blur before saving" is on. There the user can highlight, draw
# and add shapes, then:
#   Copy   - put the image on the clipboard (paste into Teams, Outlook ...)
#   Save   - save it to the folder / Word document
#   Cancel - discard it

def annotate_then_save(image, kind, taken_at):
    """Show the captured image in the Review window; save only on Save."""
    root.after(50, lambda: AnnotateEditor(
        image,
        on_save=lambda result: handle_capture(result, kind, taken_at),
        title=f"{kind} capture",
        save_label="Save to Word" if word_doc_path is not None else "Save",
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
      Text (T)  Crop (C)
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
                 text="Annotate if you like, then Copy, Save or Cancel.  Ctrl+C copy, Ctrl+Z undo",
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
            if key != "eraser":
                self.color_frame.pack(side="left", padx=S(10), pady=S(7))
            self.size_frame.pack(side="left", padx=S(14), pady=S(7))
            if key == "text":
                self.text_box_toggle.pack(side="left", padx=(S(4), S(10)), pady=S(7))
        self.size_label.configure(text="Size" if key == "text" else "Thickness")
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
        elif kind == "text":
            draw_text(self.shape_layer, op["pos"], op["text"], op["color"],
                      op["size"], op["box"])
        else:
            draw_shape(self.shape_layer, kind, op["box"], op["color"], op["width"])

    def _compose(self):
        """Screenshot + highlight (multiply) + ink + shapes -> self.working."""
        image = self.original.copy()
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
            self.canvas.yview_scroll(int(-event.delta / 120) * 3, "units")

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
        return ("Segoe UI", -self._screen_width(), "bold")

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
        status_var.set("Screenshot copied to the clipboard - paste it with Ctrl+V")

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
        return ImageGrab.grab(bbox=self.bbox, all_screens=True).convert("RGB")

    def _escape_pressed(self):
        """Esc is down now, or was pressed since the last check (works
        whichever window has the focus)."""
        try:
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
            from ctypes import wintypes
            point = wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(point))
            self.saved_cursor = (point.x, point.y)
        except Exception:
            self.saved_cursor = None

    def _restore_cursor(self):
        try:
            if self.saved_cursor:
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
    if sys.platform != "win32":
        messagebox.showinfo("Not available", "Full-page capture needs Windows.", parent=root)
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
        if not blur_review_var.get():
            root.after(300, lambda: status_var.set(
                f"{status_var.get()}\nFull page: {message}"))

    capture = ScrollingCapture((x1, y1, x2, y2), done, progress)
    scroll_stop_button.config(command=capture.stop, state="normal")
    root.bind("<Escape>", lambda e: capture.stop())
    enter_compact_mode(scroll_bar, area, "Capturing full page")
    capture.start()


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
        screenshot = ImageGrab.grab(all_screens=True)
        deliver_capture(screenshot, "Full screen", taken_at)

    except Exception as e:
        print(f"Error capturing screen: {e}")

    finally:
        root.deiconify()
        root.lift()


# ============================================================
# Region Selection Overlay (shared by screenshot and recording)
# ============================================================

def start_region_selection(on_selected, hint=None):
    """
    Hide the main window and show the selection overlay.
    on_selected(left, top, right, bottom) is called with screen
    coordinates after the user drags out a rectangle.
    """
    root.withdraw()
    root.after(200, lambda: create_selection_overlay(on_selected, hint))


def create_selection_overlay(on_selected, hint=None):
    overlay = tk.Toplevel(root)
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
                           fill="white", font=("Segoe UI Semibold", 14))

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
            return

        # Wait for overlay to disappear
        root.after(200, lambda: on_selected(x1, y1, x2, y2))

    def cancel_selection(event):
        overlay.destroy()
        root.deiconify()
        root.lift()
        status_var.set("Selection cancelled")

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
        screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
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
                shot = ImageGrab.grab(bbox=bbox, all_screens=True)

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
        print("Note: this version of Windows can't hide the recording "
              "controls from capture.")


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
        recorder = ScreenRecorder(
            filename, area,
            record_mic=mic_var.get() and AUDIO_AVAILABLE,
            record_system=system_audio_var.get() and AUDIO_AVAILABLE,
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
        message += ("\nNote: the controls may appear in the video "
                    "(hiding them needs Windows 10 version 2004 or later).")
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

    # Make sure every screenshot ends up in the Word document
    while pending_captures and not flush_pending(interactive=True):
        retry = messagebox.askretrycancel(
            "Word document is open",
            f"{len(pending_captures)} capture(s) haven't been added to:\n\n"
            f"{word_doc_path}\n\n"
            "The document is probably open in Word. Close it there, "
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

root = tk.Tk()
root.withdraw()   # shown, centred, once the layout is built
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

# Fonts (Segoe UI on Windows; Tk falls back elsewhere)
F_TITLE = ("Segoe UI Semibold", 16)
F_SUB = ("Segoe UI", 10)
F_SECTION = ("Segoe UI Semibold", 9)
F_BODY = ("Segoe UI", 10)
F_BODY_BOLD = ("Segoe UI Semibold", 10)
F_TILE = ("Segoe UI Semibold", 11)
F_TILE_SUB = ("Segoe UI", 9)
F_TIMER = ("Consolas", 13, "bold")

# Scale pixel sizes with Windows display scaling (fonts scale by themselves)
UI_SCALE = max(1.0, root.winfo_fpixels("1i") / 96.0)


def S(pixels):
    return int(round(pixels * UI_SCALE))


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
        ICONS[key] = draw_icon(kind, color, S(size), accent)
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

class FlatButton(tk.Button):
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
            padx=S(14), pady=S(6),
        )
        if text in self._icons:
            options.update(image=self._icons[text], compound="left")
        options.update(kw)
        super().__init__(parent, **options)
        self.bind("<Enter>", lambda e: self._hover_on(True))
        self.bind("<Leave>", lambda e: self._hover_on(False))

    def _hover_on(self, on):
        if str(self["state"]) != "disabled":
            tk.Button.configure(self, bg=self._hover if on else self._bg)

    def configure(self, cnf=None, **kw):
        # Swap the icon together with the text (e.g. Pause <-> Resume)
        if "text" in kw and kw["text"] in self._icons:
            kw["image"] = self._icons[kw["text"]]
        if kw.get("state") == "disabled":
            kw.setdefault("cursor", "arrow")
            kw.setdefault("bg", self._bg)
        elif kw.get("state") == "normal":
            kw.setdefault("cursor", "hand2")
        return tk.Button.configure(self, cnf, **kw)

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
                                 width=S(44), height=S(44))
        self.icon_box.grid(row=0, column=0, rowspan=2, padx=(S(14), S(12)), pady=S(11))

        self.title = tk.Label(self, text=title, font=F_TILE, fg=TEXT, bg=CARD, anchor="w")
        self.title.grid(row=0, column=1, sticky="sw", pady=(S(11), 0), padx=(0, S(14)))
        self.sub = tk.Label(self, text=subtitle, font=F_TILE_SUB, fg=MUTED, bg=CARD, anchor="w",
                            justify="left")
        self.sub.grid(row=1, column=1, sticky="nw", pady=(S(1), S(11)), padx=(0, S(14)))

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
        head.pack(fill="x", padx=S(20), pady=(S(11), S(8)))
        tk.Frame(head, bg=ACCENT, width=S(4), height=S(16)).pack(side="left", padx=(0, S(8)))
        tk.Label(head, text=title.upper(), font=F_SECTION, fg=TEXT, bg=CARD).pack(side="left")
        if required:   # red asterisk = must be filled in
            tk.Label(head, text="*", font=(F_SECTION[0], F_SECTION[1] + 3, "bold"),
                     fg=RED, bg=CARD).pack(side="left", padx=(S(3), 0))
        self.badge = tk.Label(head, font=F_SECTION, bg=ACCENT_SOFT, fg=ACCENT, padx=S(8), pady=S(1))
        if badge:
            self.set_badge(badge)
        # divider between heading and content
        tk.Frame(self, bg="#E3E6EC", height=1).pack(fill="x", padx=S(20))
        self.body = tk.Frame(self, bg=CARD)
        self.body.pack(fill="x", padx=S(20), pady=(S(11), S(14)))

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
header_inner.pack(fill="x", padx=S(24), pady=S(14))

tk.Label(header_inner, image=icon("logo", size=40), bg=HEADER_BG).pack(side="left")

header_text = tk.Frame(header_inner, bg=HEADER_BG)
header_text.pack(side="left", padx=(S(12), 0))
tk.Label(header_text, text="Screen Capture Utility", font=F_TITLE,
         fg="white", bg=HEADER_BG).pack(anchor="w")
header_subtitle = tk.Label(header_text,
                           text="Screenshots, recordings with sound, and Word reports",
                           font=F_SUB, fg=HEADER_SUB, bg=HEADER_BG)
header_subtitle.pack(anchor="w")

# thin accent line under the header
tk.Frame(main_frame, bg=ACCENT, height=S(3)).pack(fill="x")

content = tk.Frame(main_frame, bg=BG)
content.pack(fill="both", padx=S(24), pady=(S(16), S(4)))


# ============================================================
# Save Location Card
# ============================================================

folder_card = Card(content, "Save location", required=True)
folder_card.pack(fill="x", pady=(0, S(12)))

entry_row = tk.Frame(folder_card.body, bg=CARD)
entry_row.pack(fill="x")

tk.Label(entry_row, image=icon("folder", ACCENT, 18), bg=CARD).pack(side="left", padx=(0, S(8)))

folder_entry = styled_entry(entry_row, folder_var)
folder_entry.configure(width=62)
folder_entry.pack(side="left", fill="x", expand=True, ipady=S(5), ipadx=S(4))

browse_button = FlatButton(entry_row, "Browse", browse_folder, style="primary")
browse_button.pack(side="left", padx=(S(8), 0))


# ============================================================
# Word Document Card
# ============================================================

word_frame = Card(content, "Word document", badge="Optional")
word_frame.pack(fill="x", pady=(0, S(12)))

word_hint_label = tk.Label(
    word_frame.body,
    text="Screenshots are added to this document in the order you take them.",
    font=F_TILE_SUB, fg=MUTED, bg=CARD)
word_hint_label.pack(anchor="w", pady=(0, S(6)))

doc_row = tk.Frame(word_frame.body, bg=CARD)
doc_row.pack(fill="x")

tk.Label(doc_row, image=icon("doc", ACCENT, 18), bg=CARD).pack(side="left", padx=(0, S(8)))

doc_entry = styled_entry(doc_row, doc_var, readonly=True)
doc_entry.pack(side="left", fill="x", expand=True, ipady=S(5), ipadx=S(4))

select_doc_button = FlatButton(doc_row, "Select", select_word_document)
select_doc_button.pack(side="left", padx=(S(8), 0))

new_doc_button = FlatButton(doc_row, "New", new_word_document)
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

if not DOCX_AVAILABLE:
    for widget in (select_doc_button, new_doc_button):
        widget.config(state="disabled")
    word_frame.set_badge("Unavailable - python-docx not installed", warn=True)


# ============================================================
# Screenshot Card
# ============================================================

shot_card = Card(content, "Screenshot & text", badge="PNG / Word / Text")
shot_card.pack(fill="x", pady=(0, S(12)))

shot_tiles = tk.Frame(shot_card.body, bg=CARD)
shot_tiles.pack(fill="x")
shot_tiles.columnconfigure((0, 1, 2, 3), weight=1, uniform="tiles")

capture_button = Tile(shot_tiles, icon("screen", ACCENT, 24), "Capture Screen",
                      "Entire desktop\nCtrl+Shift+F", capture_screen)
capture_button.grid(row=0, column=0, sticky="nsew", padx=(0, S(5)))

region_button = Tile(shot_tiles, icon("region", ACCENT, 24), "Capture Region",
                     "Drag to select\nCtrl+Shift+R", select_region)
region_button.grid(row=0, column=1, sticky="nsew", padx=(S(5), S(5)))

fullpage_button = Tile(shot_tiles, icon("fullpage", ACCENT, 24), "Full Page",
                       "Scrolls long pages\nCtrl+Shift+S", full_page_capture)
fullpage_button.grid(row=0, column=2, sticky="nsew", padx=(S(5), S(5)))

ocr_button = Tile(shot_tiles, icon("ocr", ACCENT, 24), "Extract Text",
                  "Copy text (OCR)\nCtrl+Shift+T", extract_text_region)
ocr_button.grid(row=0, column=3, sticky="nsew", padx=(S(5), 0))

if not OCR_AVAILABLE:
    ocr_button.sub.configure(text="Needs Windows OCR\n(see console)")

blur_review_var = tk.BooleanVar(value=False)

blur_row = tk.Frame(shot_card.body, bg=CARD)
blur_row.pack(fill="x", pady=(S(10), 0))

blur_toggle = Toggle(blur_row, "Review & blur before saving", blur_review_var,
                     image=icon("blur", MUTED, 16))
blur_toggle.pack(side="left")

blur_file_button = FlatButton(blur_row, "Blur an image...", blur_existing_image)
blur_file_button.pack(side="right")


# ============================================================
# Recording Card
# ============================================================

record_frame = Card(content, "Screen recording", badge="MP4")
record_frame.pack(fill="x", pady=(0, S(12)))

record_tiles = tk.Frame(record_frame.body, bg=CARD)
record_tiles.pack(fill="x")
record_tiles.columnconfigure((0, 1), weight=1, uniform="tiles")

record_button = Tile(record_tiles, icon("record", TEXT, 24), "Record Screen",
                     "Full screen  ·  Ctrl+Shift+V", toggle_screen_recording,
                     hover_border=RED)
record_button.grid(row=0, column=0, sticky="nsew", padx=(0, S(6)))

record_region_button = Tile(record_tiles, icon("record_region", TEXT, 24), "Record Region",
                            "Drag to select an area", record_region,
                            hover_border=RED)
record_region_button.grid(row=0, column=1, sticky="nsew", padx=(S(6), 0))

# Sound options
mic_var = tk.BooleanVar(value=AUDIO_AVAILABLE)
system_audio_var = tk.BooleanVar(value=AUDIO_AVAILABLE)

audio_row = tk.Frame(record_frame.body, bg=CARD)
audio_row.pack(fill="x", pady=(S(10), 0))

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
                      font=(F_BODY[0], F_BODY[1] + 3, "bold"))
status_dot.pack(side="left", anchor="n", padx=(S(16), S(8)), pady=(S(13), 0))

status_label = tk.Label(
    status_bar,
    textvariable=status_var,
    font=F_BODY,
    fg=TEXT,
    bg=CARD,
    height=3,
    justify="left",
    anchor="w"
)
status_label.pack(side="left", fill="x", expand=True, padx=(0, S(16)), pady=S(6))


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
    status_dot.pack_configure(anchor="n", pady=(S(13), 0))
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


def screen_work_area():
    """(left, top, width, height) of the usable area (without the taskbar)
    of the monitor the mouse is on - the one the user is looking at."""
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
root.bind("<Control-Shift-T>",
          lambda event: extract_text_region() if recorder is None else None)
root.bind("<Control-Shift-V>", lambda event: toggle_screen_recording())
root.bind("<Control-Shift-P>", lambda event: toggle_pause())


# ============================================================
# Start Application
# ============================================================

center_window()

print("Screen Capture Utility started.")
print("Choose a save folder in the window (type a path or click Browse...).")
print("Optionally select a Word document to collect screenshots in.")

root.mainloop()