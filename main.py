import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import ImageGrab
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


# ============================================================
# Optional packages (installed automatically if missing)
# ============================================================
# Screen recording needs OpenCV (to write .mp4) and NumPy.
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

RECORDING_AVAILABLE = np is not None and cv2 is not None
DOCX_AVAILABLE = docx is not None

if DOCX_AVAILABLE:
    from docx import Document
    from docx.shared import Emu, Inches, Pt

if not RECORDING_AVAILABLE:
    print("Screen recording is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install opencv-python numpy')

if not DOCX_AVAILABLE:
    print("Saving to Word documents is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install python-docx')


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
        # Fit the image to the page width (and height), keeping its
        # proportions, but never enlarge it beyond its real size.
        w_px, h_px = capture.size
        width = min(usable_w, w_px * EMU_PER_PIXEL)
        if width * h_px / w_px > max_h:
            width = max_h * w_px / h_px

        doc.add_picture(BytesIO(capture.png), width=Emu(int(width)))

        # A little space below each image so they don't touch
        doc.paragraphs[-1].paragraph_format.space_after = Pt(12)

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
        doc = Document()
        doc.add_heading("Screen Captures", level=1)
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

class ScreenRecorder:
    """
    Records an area of the screen to an .mp4 file in a background
    thread. Frames are written on a fixed clock so the video plays
    back at real-time speed even if capturing is sometimes slow.
    Recording can be paused and resumed; paused time is left out of
    the video.
    """

    def __init__(self, filename, area, fps=RECORDING_FPS):
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
        self.frames_written = 0
        self.start_time = None
        self.paused = False
        self.pause_started = None
        self.paused_total = 0.0

    def start(self):
        self.start_time = time.perf_counter()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self):
        self.stop_event.set()
        if self.thread is not None:
            self.thread.join(timeout=10)

    def pause(self):
        with self.lock:
            if not self.paused:
                self.paused = True
                self.pause_started = time.perf_counter()

    def resume(self):
        with self.lock:
            if self.paused:
                self.paused_total += time.perf_counter() - self.pause_started
                self.paused = False
                self.pause_started = None

    def is_running(self):
        return self.thread is not None and self.thread.is_alive()

    def elapsed(self):
        """Recorded time in seconds (not counting paused time)."""
        with self.lock:
            if self.start_time is None:
                return 0
            now = self.pause_started if self.paused else time.perf_counter()
            return now - self.start_time - self.paused_total

    def _run(self):
        width = self.area["width"]
        height = self.area["height"]

        writer = cv2.VideoWriter(
            str(self.filename),
            cv2.VideoWriter_fourcc(*"mp4v"),
            self.fps,
            (width, height)
        )

        if not writer.isOpened():
            self.error = "Could not open the video file for writing."
            return

        frame_interval = 1.0 / self.fps

        bbox = (
            self.area["left"],
            self.area["top"],
            self.area["left"] + width,
            self.area["top"] + height,
        )

        try:
            with self.lock:
                self.start_time = time.perf_counter()

            while not self.stop_event.is_set():
                if self.paused:
                    self.stop_event.wait(0.05)
                    continue

                shot = ImageGrab.grab(bbox=bbox, all_screens=True)

                # Paused while grabbing? Drop this frame.
                if self.paused:
                    continue

                frame = np.array(shot.convert("RGB"))
                frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)

                # Make sure the frame is exactly the video size
                frame = frame[:height, :width]
                if frame.shape[0] != height or frame.shape[1] != width:
                    frame = cv2.resize(frame, (width, height))

                # Write as many copies as needed to keep real-time speed
                target_frames = int(self.elapsed() * self.fps) + 1
                while self.frames_written < target_frames:
                    writer.write(frame)
                    self.frames_written += 1

                # Sleep until the next frame is due
                delay = self.frames_written * frame_interval - self.elapsed()
                if delay > 0:
                    self.stop_event.wait(delay)

        except Exception as e:
            self.error = str(e)

        finally:
            writer.release()


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

BAR_BG = "#202124"
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

    if not RECORDING_AVAILABLE:
        return

    if not ensure_save_folder():
        status_var.set("No save folder selected - recording cancelled")
        return

    area = primary_monitor_area()

    # Hide the full window first, then show the bar and start
    root.withdraw()
    root.after(300, lambda: start_recording(area))


def record_region():
    if recorder is not None or not RECORDING_AVAILABLE:
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
        recorder = ScreenRecorder(filename, area)
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


def stop_recording():
    global recorder

    if recorder is None:
        return

    current = recorder
    recorder = None

    rec_time_label.config(text="Saving...")
    root.update_idletasks()

    current.stop()

    exit_recording_mode()

    if current.error:
        print(f"Error while recording: {current.error}")
        status_var.set(f"Recording failed: {current.error}")
        messagebox.showerror("Recording failed", current.error, parent=root)
        return

    if current.frames_written == 0:
        status_var.set("Recording failed: no frames were captured.")
        messagebox.showerror(
            "Recording failed",
            "No frames were captured, so the video is empty.",
            parent=root
        )
        return

    duration = current.frames_written / current.fps
    print("Recording saved successfully:")
    print(current.filename)
    message = f"Saved: {current.filename.name}  ({duration:.1f} s)"
    if not bar_excluded:
        message += ("\nNote: the recording controls may be visible in the "
                    "video (hiding them needs Windows 10 version 2004 or later).")
    status_var.set(message)


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
root.title("Screen Capture Utility")
root.resizable(False, False)
root.protocol("WM_DELETE_WINDOW", on_close)

folder_var = tk.StringVar()
doc_var = tk.StringVar()
keep_png_var = tk.BooleanVar(value=False)
status_var = tk.StringVar(value="Choose a folder to save screenshots and recordings.")

# Everything in the normal window lives in main_frame, so it can be
# swapped for the small control bar while recording.
main_frame = tk.Frame(root)
main_frame.pack()


# ============================================================
# Recording Control Bar (shown only while recording)
# ============================================================

rec_bar = tk.Frame(root, bg=BAR_BG, padx=10, pady=8)

rec_dot = tk.Canvas(rec_bar, width=14, height=14, bg=BAR_BG,
                    highlightthickness=0)
rec_dot_id = rec_dot.create_oval(2, 2, 13, 13, fill=DOT_ON, outline="")
rec_dot.pack(side="left", padx=(0, 6))

rec_time_label = tk.Label(rec_bar, text="00:00", fg="white", bg=BAR_BG,
                          font=("Segoe UI", 11, "bold"), width=7, anchor="w")
rec_time_label.pack(side="left", padx=(0, 8))

rec_pause_button = tk.Button(
    rec_bar, text="Pause", command=toggle_pause,
    bg="#3c4043", fg="white", activebackground="#5f6368",
    activeforeground="white", relief="flat", bd=0, padx=10, pady=3,
    font=("Segoe UI", 9, "bold"), cursor="hand2", width=9
)
rec_pause_button.pack(side="left", padx=(0, 6))

rec_stop_button = tk.Button(
    rec_bar, text="Stop", command=stop_recording,
    bg="#d93025", fg="white", activebackground="#b3261e",
    activeforeground="white", relief="flat", bd=0, padx=10, pady=3,
    font=("Segoe UI", 9, "bold"), cursor="hand2", width=7
)
rec_stop_button.pack(side="left")


# ============================================================
# Save Folder Row
# ============================================================

folder_frame = tk.Frame(main_frame)
folder_frame.pack(pady=(15, 5), padx=15, fill="x")

tk.Label(folder_frame, text="Save folder:").pack(anchor="w")

entry_row = tk.Frame(folder_frame)
entry_row.pack(fill="x")

folder_entry = tk.Entry(entry_row, textvariable=folder_var, width=50)
folder_entry.pack(side="left", fill="x", expand=True, ipady=3)

browse_button = tk.Button(
    entry_row,
    text="Browse...",
    command=browse_folder
)
browse_button.pack(side="left", padx=(5, 0))


# ============================================================
# Word Document Row
# ============================================================

word_frame = tk.LabelFrame(
    main_frame,
    text="Word document (optional) - screenshots are added in capture order",
    padx=10,
    pady=8
)
word_frame.pack(pady=(10, 5), padx=15, fill="x")

doc_row = tk.Frame(word_frame)
doc_row.pack(fill="x")

doc_entry = tk.Entry(doc_row, textvariable=doc_var, state="readonly")
doc_entry.pack(side="left", fill="x", expand=True, ipady=3)

select_doc_button = tk.Button(doc_row, text="Select...",
                              command=select_word_document)
select_doc_button.pack(side="left", padx=(5, 0))

new_doc_button = tk.Button(doc_row, text="New...",
                           command=new_word_document)
new_doc_button.pack(side="left", padx=(5, 0))

clear_doc_button = tk.Button(doc_row, text="Clear",
                             command=clear_word_document)
clear_doc_button.pack(side="left", padx=(5, 0))

doc_options_row = tk.Frame(word_frame)
doc_options_row.pack(fill="x", pady=(6, 0))

keep_png_check = tk.Checkbutton(
    doc_options_row,
    text="Also save separate .png files",
    variable=keep_png_var
)
keep_png_check.pack(side="left")

open_doc_button = tk.Button(doc_options_row, text="Open Document",
                            command=open_word_document)
open_doc_button.pack(side="right")

save_pending_button = tk.Button(doc_options_row, text="Save to Word now",
                                command=lambda: flush_pending(interactive=True))
save_pending_button.pack(side="right", padx=(0, 5))

if not DOCX_AVAILABLE:
    for widget in (select_doc_button, new_doc_button):
        widget.config(state="disabled")
    word_frame.config(text="Word document - unavailable (python-docx not installed)")


# ============================================================
# Screenshot Buttons
# ============================================================

screenshot_frame = tk.LabelFrame(main_frame, text="Screenshot", padx=10, pady=8)
screenshot_frame.pack(pady=5, padx=15, fill="x")

capture_button = tk.Button(
    screenshot_frame,
    text="Capture Screen",
    command=capture_screen,
    width=18,
    height=2
)
capture_button.pack(side="left", expand=True)

region_button = tk.Button(
    screenshot_frame,
    text="Capture Region",
    command=select_region,
    width=18,
    height=2
)
region_button.pack(side="left", expand=True)


# ============================================================
# Recording Buttons
# ============================================================

record_frame = tk.LabelFrame(main_frame, text="Screen Recording (.mp4)", padx=10, pady=8)
record_frame.pack(pady=5, padx=15, fill="x")

record_button = tk.Button(
    record_frame,
    text="Record Screen",
    command=toggle_screen_recording,
    width=18,
    height=2
)
record_button.pack(side="left", expand=True)

record_region_button = tk.Button(
    record_frame,
    text="Record Region",
    command=record_region,
    width=18,
    height=2
)
record_region_button.pack(side="left", expand=True)

if not RECORDING_AVAILABLE:
    record_button.config(state="disabled")
    record_region_button.config(state="disabled")
    record_frame.config(text="Screen Recording (.mp4) - unavailable, see console")


# ============================================================
# Status Line
# ============================================================

status_label = tk.Label(
    main_frame,
    textvariable=status_var,
    fg="gray25",
    height=3,
    justify="center"
)
status_label.pack(pady=(5, 10), padx=15, fill="x")


def fit_status_wrap(event=None):
    status_label.config(wraplength=max(200, root.winfo_width() - 40))


root.after(100, fit_status_wrap)

update_word_controls()


# ============================================================
# Keyboard Shortcuts (work while this window has focus)
# ============================================================

root.bind("<Control-Shift-F>",
          lambda event: capture_screen() if recorder is None else None)
root.bind("<Control-Shift-R>",
          lambda event: select_region() if recorder is None else None)
root.bind("<Control-Shift-V>", lambda event: toggle_screen_recording())
root.bind("<Control-Shift-P>", lambda event: toggle_pause())


# ============================================================
# Start Application
# ============================================================

print("Screen Capture Utility started.")
print("Choose a save folder in the window (type a path or click Browse...).")
print("Optionally select a Word document to collect screenshots in.")

root.mainloop()