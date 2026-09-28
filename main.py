import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import ImageGrab
from datetime import datetime
from pathlib import Path
import threading
import time
import ctypes
import sys
import subprocess
import importlib


# ============================================================
# Recording packages (installed automatically if missing)
# ============================================================
# Screen recording needs OpenCV (to write .mp4) and NumPy.
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

RECORDING_AVAILABLE = np is not None and cv2 is not None

if not RECORDING_AVAILABLE:
    print("Screen recording is disabled. To enable it, run:")
    print(f'  "{sys.executable}" -m pip install opencv-python numpy')


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
# Recording Settings
# ============================================================

RECORDING_FPS = 15


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

def timestamp():
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def get_screenshot_filename():
    return save_folder / f"Screenshot_{timestamp()}.png"


def get_recording_filename():
    return save_folder / f"Recording_{timestamp()}.mp4"


# ============================================================
# Save Screenshot
# ============================================================

def save_screenshot(image):
    filename = get_screenshot_filename()

    try:
        image.save(filename, format="PNG")
        print("Screenshot saved successfully:")
        print(filename)
        status_var.set(f"Saved: {filename.name}")

    except Exception as e:
        print(f"Error saving screenshot: {e}")
        status_var.set("Error saving screenshot (see console)")


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
        screenshot = ImageGrab.grab(all_screens=True)
        save_screenshot(screenshot)

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
        screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        save_screenshot(screenshot)

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
        self.thread = None
        self.error = None
        self.frames_written = 0
        self.start_time = None

    def start(self):
        self.start_time = time.perf_counter()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self):
        self.stop_event.set()
        if self.thread is not None:
            self.thread.join(timeout=10)

    def is_running(self):
        return self.thread is not None and self.thread.is_alive()

    def elapsed(self):
        if self.start_time is None:
            return 0
        return time.perf_counter() - self.start_time

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
            self.start_time = time.perf_counter()

            while not self.stop_event.is_set():
                shot = ImageGrab.grab(bbox=bbox)

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
                next_due = self.start_time + self.frames_written * frame_interval
                delay = next_due - time.perf_counter()
                if delay > 0:
                    self.stop_event.wait(delay)

        except Exception as e:
            self.error = str(e)

        finally:
            writer.release()


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

    start_recording(primary_monitor_area())


def record_region():
    if recorder is not None:
        return

    if not ensure_save_folder():
        status_var.set("No save folder selected - recording cancelled")
        return

    start_region_selection(start_region_recording)


def start_region_recording(x1, y1, x2, y2):
    root.deiconify()
    root.lift()

    start_recording({
        "left": x1,
        "top": y1,
        "width": x2 - x1,
        "height": y2 - y1,
    })


def start_recording(area):
    global recorder

    filename = get_recording_filename()

    recorder = ScreenRecorder(filename, area)
    recorder.start()

    print(f"Recording started: {filename}")

    record_button.config(text="Stop Recording", fg="white", bg="#c62828",
                         activebackground="#b71c1c", activeforeground="white")
    record_region_button.config(state="disabled")

    update_recording_status()


def stop_recording():
    global recorder

    if recorder is None:
        return

    current = recorder
    recorder = None

    status_var.set("Finishing recording...")
    root.update_idletasks()

    current.stop()

    reset_record_buttons()

    if current.error:
        print(f"Error while recording: {current.error}")
        status_var.set("Recording failed (see console)")
        messagebox.showerror("Recording failed", current.error, parent=root)
        return

    duration = current.frames_written / current.fps
    print("Recording saved successfully:")
    print(current.filename)
    status_var.set(f"Saved: {current.filename.name}  ({duration:.1f} s)")


def reset_record_buttons():
    record_button.config(text="Record Screen", fg="black",
                         bg=default_button_bg,
                         activebackground=default_button_active_bg,
                         activeforeground="black")
    record_region_button.config(state="normal")


def update_recording_status():
    """Refresh the elapsed-time display while recording (runs on the UI thread)."""
    if recorder is None:
        return

    # The recording thread stopped on its own because of an error
    if not recorder.is_running() and recorder.error:
        stop_recording()
        return

    seconds = int(recorder.elapsed())
    minutes, seconds = divmod(seconds, 60)
    status_var.set(f"\u25CF Recording...  {minutes:02d}:{seconds:02d}   "
                   f"(click Stop Recording or press Ctrl+Shift+V)")

    root.after(500, update_recording_status)


# ============================================================
# Close Handling
# ============================================================

def on_close():
    # Finish the video file properly so it isn't left corrupted
    if recorder is not None:
        stop_recording()
    root.destroy()


# ============================================================
# Main Window
# ============================================================

root = tk.Tk()
root.title("Screen Capture Utility")
root.geometry("480x400")
root.resizable(False, False)
root.protocol("WM_DELETE_WINDOW", on_close)

folder_var = tk.StringVar()
status_var = tk.StringVar(value="Choose a folder to save screenshots and recordings.")


# ============================================================
# Save Folder Row
# ============================================================

folder_frame = tk.Frame(root)
folder_frame.pack(pady=(15, 5), padx=15, fill="x")

tk.Label(folder_frame, text="Save folder:").pack(anchor="w")

entry_row = tk.Frame(folder_frame)
entry_row.pack(fill="x")

folder_entry = tk.Entry(entry_row, textvariable=folder_var)
folder_entry.pack(side="left", fill="x", expand=True, ipady=3)

browse_button = tk.Button(
    entry_row,
    text="Browse...",
    command=browse_folder
)
browse_button.pack(side="left", padx=(5, 0))


# ============================================================
# Screenshot Buttons
# ============================================================

screenshot_frame = tk.LabelFrame(root, text="Screenshot (.png)", padx=10, pady=8)
screenshot_frame.pack(pady=(12, 5), padx=15, fill="x")

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

record_frame = tk.LabelFrame(root, text="Screen Recording (.mp4)", padx=10, pady=8)
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

default_button_bg = record_button.cget("bg")
default_button_active_bg = record_button.cget("activebackground")

if not RECORDING_AVAILABLE:
    record_button.config(state="disabled")
    record_region_button.config(state="disabled")
    record_frame.config(text="Screen Recording (.mp4) - unavailable, see console")


# ============================================================
# Status Line
# ============================================================

status_label = tk.Label(
    root,
    textvariable=status_var,
    fg="gray25",
    wraplength=450,
    justify="center"
)
status_label.pack(pady=(10, 0))


# ============================================================
# Keyboard Shortcuts (work while this window has focus)
# ============================================================

root.bind("<Control-Shift-F>", lambda event: capture_screen())
root.bind("<Control-Shift-R>", lambda event: select_region())
root.bind("<Control-Shift-V>", lambda event: toggle_screen_recording())


# ============================================================
# Start Application
# ============================================================

print("Screen Capture Utility started.")
print("Choose a save folder in the window (type a path or click Browse...).")

root.mainloop()