import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageGrab, ImageOps
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


np = ensure_package("numpy", "numpy")
cv2 = ensure_package("cv2", "opencv-python")
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
        if width * h_px / w_px > max_h:
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
        handle_capture(screenshot, "Full screen", taken_at)

    except Exception as e:
        print(f"Error capturing screen: {e}")

    finally:
        root.deiconify()
        root.lift()


# ============================================================
# Region Selection Overlay (shared by screenshot and recording)
# ============================================================

def start_region_selection(on_selected):
    """
    Hide the main window and show the selection overlay.
    on_selected(left, top, right, bottom) is called with screen
    coordinates after the user drags out a rectangle.
    """
    root.withdraw()
    root.after(200, lambda: create_selection_overlay(on_selected))


def create_selection_overlay(on_selected):
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
        handle_capture(screenshot, "Region", taken_at)

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


def bar_position_for(area):
    """Top-left corner for the control bar: just below the recorded
    area, else just above it, else inside its top edge."""
    root.update_idletasks()
    w = rec_bar.winfo_reqwidth() + 20      # + window border
    h = rec_bar.winfo_reqheight() + 45     # + title bar
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

    elif kind == "ocr":
        L = 5
        for (x, y, dx, dy) in [(3, 3, 1, 1), (21, 3, -1, 1), (3, 21, 1, -1), (21, 21, -1, -1)]:
            d.line([x * u, y * u, (x + dx * L) * u, y * u], fill=color, width=w)
            d.line([x * u, y * u, x * u, (y + dy * L) * u], fill=color, width=w)
        d.line(box(7.5, 9, 16.5, 9), fill=color, width=w)
        d.line(box(7.5, 12.5, 16.5, 12.5), fill=color, width=w)
        d.line(box(7.5, 16, 13, 16), fill=color, width=w)

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

    def __init__(self, parent, title, badge=None):
        super().__init__(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        head = tk.Frame(self, bg=CARD)
        head.pack(fill="x", padx=S(20), pady=(S(11), S(8)))
        tk.Frame(head, bg=ACCENT, width=S(4), height=S(16)).pack(side="left", padx=(0, S(8)))
        tk.Label(head, text=title.upper(), font=F_SECTION, fg=TEXT, bg=CARD).pack(side="left")
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

folder_card = Card(content, "Save location")
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
shot_tiles.columnconfigure((0, 1, 2), weight=1, uniform="tiles")

capture_button = Tile(shot_tiles, icon("screen", ACCENT, 24), "Capture Screen",
                      "Entire desktop\nCtrl+Shift+F", capture_screen)
capture_button.grid(row=0, column=0, sticky="nsew", padx=(0, S(6)))

region_button = Tile(shot_tiles, icon("region", ACCENT, 24), "Capture Region",
                     "Drag to select\nCtrl+Shift+R", select_region)
region_button.grid(row=0, column=1, sticky="nsew", padx=(S(6), S(6)))

ocr_button = Tile(shot_tiles, icon("ocr", ACCENT, 24), "Extract Text",
                  "Copy text from screen\nCtrl+Shift+T", extract_text_region)
ocr_button.grid(row=0, column=2, sticky="nsew", padx=(S(6), 0))

if not OCR_AVAILABLE:
    ocr_button.sub.configure(text="Needs Windows OCR\n(see console)")


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

status_dot = tk.Label(status_bar, bg=CARD)
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

update_word_controls()


# ============================================================
# Keyboard Shortcuts (work while this window has focus)
# ============================================================

root.bind("<Control-Shift-F>",
          lambda event: capture_screen() if recorder is None else None)
root.bind("<Control-Shift-R>",
          lambda event: select_region() if recorder is None else None)
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