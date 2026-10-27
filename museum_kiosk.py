r"""
Museum Kiosk Launcher (MVP)
---------------------------
Fullscreen menu of the PowerPoint files and videos in the presentations folder, each shown
as a card with a preview (first slide / video frame). Tap one to play it full screen:
    Presentations  play as a PowerPoint slideshow
    Videos         play in Microsoft Edge (kiosk mode, built-in HTML5 player); the media
                   controls appear on touch/click and hide themselves while playing
Esc or the "Back to menu" button (top-right) ends playback and the menu returns.
Esc is watched globally, so it works whatever has focus. The button fades out after a few
idle seconds; any touch, click or key brings it back.
Right-click during a slideshow is left to PowerPoint (its normal context menu).

Presentations folder (first match wins):
    1. The folder saved in settings (chosen on first run, or in the settings panel)
    2. A "Presentations" folder next to this script / the exe
    3. Neither -> the app asks for a folder on startup

Settings, log and caches:  C:\ProgramData\MuseumKiosk\
    settings.json   presentations_folder, loop
    kiosk.log
    thumbs\         previews (rebuilt automatically when a file changes)
    player.html     the video player page (rewritten for each video)
    edge-profile\   Edge's data for the video player

Staff settings panel (presentations folder, loop on/off, Quit Kiosk button):
    Touch:     tap the screen corners in order  top-left > top-right > bottom-right > bottom-left
               (within 5 seconds, on the menu screen)
    Keyboard:  Ctrl+Alt+S   (not during playback)
    Ctrl+Alt+Q quits the kiosk directly, during playback too (playback is closed first).

Loop setting (applies to every deck and video, overrides a deck's own setting):
    on   loop until Esc / Back to menu
    off  play once, then return to the menu
Slide advancing (applied to every deck, in memory only - files are never changed):
    tap / swipe left   next slide, on every slide (timed or not)
    swipe right        previous slide
    timed slides       also advance by themselves (Transitions > Advance Slide > After)

Run from source:  py -m pip install -r requirements.txt   then   py museum_kiosk.py
Build exe:        build.bat   (puts MuseumKiosk-v<VERSION>.exe in the project folder)

Requires PowerPoint installed and licensed (decks + all previews) and Microsoft Edge
(video playback). Provided as-is.
"""

# VERSION - bump on every change:
#   0.0.x  minor changes (fixes, tweaks, text/config/color changes)
#   0.x.0  major revisions (features added/removed, behavior changes) - reset the last digit to 0
# This project stays in alpha: always 0.x.x, never bump to 1.0.0.
VERSION = "0.8.0"

import ctypes
import hashlib
import html
import json
import logging
import os
import subprocess
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import win32com.client   # pywin32 - drives PowerPoint over COM


def app_dir():
    """Folder containing the .exe (when built) or this .py file."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# ---------- CONFIGURATION ----------
APP_DIR = app_dir()
DATA_DIR = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData")) / "MuseumKiosk"
SETTINGS_FILE = DATA_DIR / "settings.json"
LOG_FILE = DATA_DIR / "kiosk.log"
THUMB_DIR = DATA_DIR / "thumbs"
PLAYER_PAGE = DATA_DIR / "player.html"
EDGE_PROFILE_DIR = DATA_DIR / "edge-profile"
DEFAULT_SETTINGS = {"presentations_folder": None, "loop": True}

TITLE_TEXT = "Choose a Presentation"
SUBTITLE_TEXT = "Tap a presentation to begin"
BACK_BUTTON_TEXT = "✕  Back to menu"     # shown top-right over every running show/video
BACK_BUTTON_IDLE_SECONDS = 3             # fades out after this long with no touch/click/key
BACK_BUTTON_ALPHA = 0.75                 # opacity when fully shown
COLUMNS = 3

BG_COLOR = "#1a1a2e"
CARD_COLOR = "#16213e"
ACCENT_COLOR = "#e94560"
TEXT_COLOR = "#ffffff"
MUTED_COLOR = "#aaaaaa"
# -----------------------------------

SIDE_PAD = 60                # space left/right of the card grid
CARD_GAP = 15                # space around each card
CARD_PAD = 12                # space inside each card around the preview
SCROLLBAR_W = 28
TAP_TOLERANCE_PX = 12        # finger drift allowed before a tap becomes a scroll
POPUP_SECONDS = 8            # error popups close themselves after this long
CORNER_PX = 70               # size of each corner hot zone for the settings tap pattern
CORNER_SECONDS = 5           # all four corner taps must happen within this time
DECK_EXTENSIONS = {".pptx", ".ppt", ".ppsx", ".pps"}
VIDEO_EXTENSIONS = {".mp4", ".m4v", ".mov", ".webm"}   # what Edge's player handles
EXTENSIONS = DECK_EXTENSIONS | VIDEO_EXTENSIONS
VIDEO_START_SECONDS = 20     # how long Edge gets to show the video window
VIDEO_PREVIEW_MS = 2000      # video preview frame: this far in (or 1/4 of a short video)
SHOW_POLL_MS = 300           # how often a running show/video is checked
CREATE_NO_WINDOW = 0x08000000
ERROR_ALREADY_EXISTS = 183
WM_CLOSE = 0x0010
SHOW_WINDOW_CLASS = "screenClass"        # PowerPoint's fullscreen slideshow window
EDGE_WINDOW_CLASS = "Chrome_WidgetWin_1" # Edge's top-level browser window
SWP_NOMOVE_NOSIZE_NOACTIVATE = 0x0002 | 0x0001 | 0x0010
SW_HIDE, SW_SHOWNOACTIVATE = 0, 4
FADE_TICK_MS = 40            # back button: activity check / fade step interval
FADE_STEPS = 8               # back button: steps per fade (8 x 40 ms = ~0.3 s)
VK_ESCAPE, VK_CONTROL, VK_MENU, VK_Q, VK_S = 0x1B, 0x11, 0x12, 0x51, 0x53   # VK_MENU = Alt

# PowerPoint COM constants
MSO_TRUE, MSO_FALSE = -1, 0
PP_SHOW_TYPE_SPEAKER = 1     # normal full-screen show (even if the deck is set to "window")
PP_SLIDESHOW_DONE = 5        # View.State once a non-looping show reaches its end
PP_LAYOUT_BLANK = 12
PP_ADVANCE_USE_TIMINGS = 2   # SlideShowSettings.AdvanceMode: use slide timings if present

# Video player page. __TOKEN__ is a unique window title the app uses to find the Edge
# window; the page adds " ended" / " error" to it so the app knows when playback stops.
PLAYER_HTML = """<!doctype html>
<html><head><meta charset="utf-8"><title>__TOKEN__</title>
<style>
  html, body { margin: 0; height: 100%; background: #000; overflow: hidden; }
  video { width: 100vw; height: 100vh; object-fit: contain; background: #000; outline: none; }
</style></head>
<body>
<video id="v" src="__SRC__" autoplay playsinline controls __LOOP__
       controlslist="nodownload nofullscreen noremoteplayback noplaybackrate"
       disablepictureinpicture disableremoteplayback></video>
<script>
  const v = document.getElementById("v");
  const report = (state) => { document.title = "__TOKEN__ " + state; };
  v.addEventListener("ended", () => report("ended"));
  v.addEventListener("error", () => report("error"));
  // If autoplay with sound is blocked, start muted rather than not at all.
  v.play().catch(() => { v.muted = true; v.play().catch(() => {}); });
  document.addEventListener("contextmenu", (e) => e.preventDefault());
</script>
</body></html>
"""


# ---------- setup helpers ----------
def ensure_data_dir():
    """Create the ProgramData folder and, when we create it, give local Users modify rights,
    so settings/log/previews work no matter which account (admin or kiosk user) runs first."""
    try:
        if not DATA_DIR.is_dir():
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            subprocess.run(["icacls", str(DATA_DIR), "/grant", "*S-1-5-32-545:(OI)(CI)M"],
                           capture_output=True, creationflags=CREATE_NO_WINDOW, timeout=15)
        THUMB_DIR.mkdir(exist_ok=True)
    except Exception:
        pass   # logging isn't set up yet; the app still runs without the folder


def setup_logging():
    try:
        logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                            format="%(asctime)s %(levelname)s %(message)s")
    except OSError:
        logging.basicConfig(level=logging.INFO)


def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    try:
        settings.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
    except FileNotFoundError:
        pass
    except (OSError, ValueError):
        logging.exception("Could not read %s - using defaults", SETTINGS_FILE)
    return settings


def save_settings(settings):
    try:
        SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")
        return True
    except OSError:
        logging.exception("Could not save %s", SETTINGS_FILE)
        return False


def resolve_folder(settings):
    """Saved folder, else ./Presentations next to the app, else None (ask the user)."""
    saved = settings.get("presentations_folder")
    if saved:
        return Path(saved)      # used even if missing right now (e.g. drive not ready) - no prompt
    local = APP_DIR / "Presentations"
    return local if local.is_dir() else None


def unblock(path):
    """Remove the 'downloaded from the internet' mark (like Unblock-File) so PowerPoint
    doesn't hold the deck in Protected View."""
    try:
        os.remove(f"{path}:Zone.Identifier")
        logging.info("Unblocked %s", path.name)
    except FileNotFoundError:
        pass
    except OSError:
        logging.warning("Could not unblock %s (no write access?)", path.name)


def is_video(path):
    return path.suffix.lower() in VIDEO_EXTENSIONS


def display_name(path):
    return path.stem.replace("_", " ").replace("-", " ")


def key_down(vk):
    return bool(ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000)


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]


def last_input_tick():
    """Tick count of the last touch/mouse/keyboard input anywhere in the session.
    Works even though PowerPoint/Edge, not this app, receives the touches."""
    info = LASTINPUTINFO(ctypes.sizeof(LASTINPUTINFO), 0)
    ctypes.windll.user32.GetLastInputInfo(ctypes.byref(info))
    return info.dwTime


# ---------- Windows window helpers ----------
def window_exists(hwnd):
    return bool(hwnd) and bool(ctypes.windll.user32.IsWindow(ctypes.c_void_p(hwnd)))


def window_title(hwnd):
    user32 = ctypes.windll.user32
    handle = ctypes.c_void_p(hwnd)
    length = user32.GetWindowTextLengthW(handle)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(handle, buf, length + 1)
    return buf.value


def window_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(ctypes.c_void_p(hwnd), buf, 256)
    return buf.value


def window_pid(hwnd):
    pid = ctypes.c_ulong(0)
    ctypes.windll.user32.GetWindowThreadProcessId(ctypes.c_void_p(hwnd), ctypes.byref(pid))
    return pid.value


def find_video_window(token, pid):
    """The Edge window showing our player page: matched by its title, or failing that by
    being Edge's visible main window in the process we started."""
    user32 = ctypes.windll.user32
    found = []

    def check(hwnd, _lparam):
        if hwnd and user32.IsWindowVisible(ctypes.c_void_p(hwnd)):
            if window_title(hwnd).startswith(token) or (
                    window_pid(hwnd) == pid and window_class(hwnd) == EDGE_WINDOW_CLASS):
                found.append(hwnd)
                return False                 # stop enumerating
        return True

    callback = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(check)
    user32.EnumWindows(callback, None)
    return found[0] if found else None


def find_edge():
    """Path to msedge.exe, or None."""
    import winreg
    key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ | view) as key:
                    path = winreg.QueryValue(key, None).strip('"')
                    if path and os.path.exists(path):
                        return path
            except OSError:
                pass
    for base in (os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                 os.environ.get("ProgramFiles", r"C:\Program Files")):
        path = os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe")
        if os.path.exists(path):
            return path
    return None


# ---------- PowerPoint (COM) ----------
def powerpoint():
    """PowerPoint is single-instance: this starts it hidden, or attaches to the running one."""
    return win32com.client.Dispatch("PowerPoint.Application")


def open_readonly(app, path):
    # Open(FileName, ReadOnly, Untitled, WithWindow): no editor window is ever shown.
    return app.Presentations.Open(str(path), MSO_TRUE, MSO_FALSE, MSO_FALSE)


def close_if_open(app, path):
    """Close any copy of this deck still open in PowerPoint, so the show starts fresh."""
    target = str(path).lower()
    for i in range(app.Presentations.Count, 0, -1):
        pres = app.Presentations.Item(i)
        if str(pres.FullName).lower() == target:
            logging.warning("%s was still open in PowerPoint - closing it first", path.name)
            pres.Close()


def enable_touch_advance(pres):
    """Make every slide advance on tap/swipe ("On Mouse Click"), timed or not.
    Timings are left as they are. Returns (slide count, timed slide count)."""
    count = pres.Slides.Count
    timed = 0
    for i in range(1, count + 1):
        try:
            transition = pres.Slides.Item(i).SlideShowTransition
            transition.AdvanceOnClick = MSO_TRUE
            if transition.AdvanceOnTime:
                timed += 1
        except Exception:
            logging.warning("Could not enable tap/swipe on slide %d", i)
    return count, timed


def quit_powerpoint_if_idle():
    """Quit PowerPoint only if nothing else is open in it."""
    try:
        app = powerpoint()
        if app.Presentations.Count == 0:
            app.Quit()
    except Exception:
        logging.exception("Could not quit PowerPoint")


def export_deck_thumbnail(path, png, width):
    pres = open_readonly(powerpoint(), path)
    try:
        ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
        pres.Slides.Item(1).Export(str(png), "PNG", width, round(width * ratio))
    finally:
        pres.Close()


def export_video_thumbnail(path, png, width):
    """Preview frame of a video: PowerPoint inserts it into a throwaway, windowless
    presentation sized to the video and exports that slide."""
    pres = powerpoint().Presentations.Add(MSO_FALSE)     # WithWindow = False
    try:
        slide = pres.Slides.Add(1, PP_LAYOUT_BLANK)
        # AddMediaObject2(FileName, LinkToFile, SaveWithDocument, Left, Top) - native size
        video = slide.Shapes.AddMediaObject2(str(path), MSO_TRUE, MSO_FALSE, 0, 0)
        w, h = video.Width, video.Height
        pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight = w, h
        video.Left, video.Top, video.Width, video.Height = 0, 0, w, h
        try:                                  # skip a black first frame / fade-in
            length = int(video.MediaFormat.Length)
            video.MediaFormat.SetDisplayPicture(min(VIDEO_PREVIEW_MS, length // 4))
        except Exception:
            pass
        slide.Export(str(png), "PNG", width, round(width * h / w))
    finally:
        pres.Saved = MSO_TRUE
        pres.Close()


# ---------- what's playing ----------
class PowerPointShow:
    """A deck running as a full-screen PowerPoint slideshow."""

    def __init__(self, path, loop):
        self.path = path
        self.hwnd = None
        app = powerpoint()
        close_if_open(app, path)
        pres = open_readonly(app, path)
        try:
            show = pres.SlideShowSettings
            show.ShowType = PP_SHOW_TYPE_SPEAKER     # "browsed at a kiosk" would ignore touch
            show.LoopUntilStopped = MSO_TRUE if loop else MSO_FALSE
            try:
                show.AdvanceMode = PP_ADVANCE_USE_TIMINGS  # timed slides play by themselves
                slides, timed = enable_touch_advance(pres)
                logging.info("%s: %d slides, %d timed - tap/swipe advances all",
                             path.name, slides, timed)
            except Exception:                    # never block playback over this
                logging.exception("Could not adjust slide advancing for %s - playing as-is",
                                  path.name)
            try:
                show.ShowPresenterView = MSO_FALSE   # no presenter view on multi-monitor PCs
            except Exception:
                pass
            pres.Saved = MSO_TRUE   # the settings above aren't real edits - no "save?" prompt
            window = show.Run()
        except Exception:
            try:
                pres.Close()
            except Exception:
                pass
            raise
        self.pres = pres
        try:
            self.hwnd = int(window.HWND)
        except Exception:
            self.hwnd = None                     # fall back to finding the window by class

    def appeared(self):
        return True                              # Run() returns once the show is up

    def is_open(self):
        """Checked through Windows, not PowerPoint, so it can't fail while PowerPoint is
        busy (e.g. its right-click menu is open)."""
        if self.hwnd:
            return window_exists(self.hwnd)
        return bool(ctypes.windll.user32.FindWindowW(SHOW_WINDOW_CLASS, None))

    def poll(self):
        """'ended' once a non-looping show reaches its end, otherwise None."""
        try:
            if not self.pres.Saved:              # keep it "unchanged" so ending the show
                self.pres.Saved = MSO_TRUE       # never asks to save the read-only deck
            if self.pres.SlideShowWindow.View.State == PP_SLIDESHOW_DONE:
                return "ended"
        except Exception:
            pass   # PowerPoint busy (e.g. right-click menu) - not a reason to end the show
        return None

    def stop(self):
        try:
            self.pres.Saved = MSO_TRUE
            self.pres.SlideShowWindow.View.Exit()
        except Exception:
            logging.warning("PowerPoint busy - show not stopped")

    def close(self):
        try:
            self.pres.Close()
        except Exception:
            logging.exception("Could not close presentation")
        quit_powerpoint_if_idle()


class VideoShow:
    """A video playing full screen in Microsoft Edge (kiosk mode) on a local player page.
    Edge's own HTML5 controls appear on touch/click and hide themselves while playing."""

    def __init__(self, path, loop):
        edge = find_edge()
        if edge is None:
            raise RuntimeError("Microsoft Edge (msedge.exe) not found")
        self.path = path
        self.hwnd = None
        self.stopped = False
        self.token = f"MuseumKioskVideo-{os.getpid()}-{time.monotonic_ns()}"
        page = (PLAYER_HTML.replace("__TOKEN__", self.token)
                .replace("__SRC__", html.escape(Path(path).resolve().as_uri(), quote=True))
                .replace("__LOOP__", "loop" if loop else ""))
        PLAYER_PAGE.write_text(page, encoding="utf-8")
        self.proc = subprocess.Popen([
            edge, "--kiosk", PLAYER_PAGE.as_uri(), "--edge-kiosk-type=fullscreen",
            "--inprivate", "--no-first-run", "--no-default-browser-check",
            f"--user-data-dir={EDGE_PROFILE_DIR}",            # separate from any user Edge
            "--autoplay-policy=no-user-gesture-required",     # start with sound, no tap
        ])
        self.started = time.monotonic()

    def appeared(self):
        if self.hwnd is None and not self.stopped:
            self.hwnd = find_video_window(self.token, self.proc.pid)
        return self.hwnd is not None

    def is_open(self):
        if self.appeared():
            return window_exists(self.hwnd)
        if self.stopped:
            return False
        return time.monotonic() - self.started < VIDEO_START_SECONDS   # still starting

    def poll(self):
        """'ended' / 'error' as reported by the player page, otherwise None."""
        title = window_title(self.hwnd) if self.hwnd else ""
        if title.startswith(self.token + " ended"):
            return "ended"
        if title.startswith(self.token + " error"):
            return "error"
        return None

    def stop(self):
        self.stopped = True
        if self.hwnd:
            ctypes.windll.user32.PostMessageW(ctypes.c_void_p(self.hwnd), WM_CLOSE, 0, 0)
        else:
            self.close()                         # stopped before the window even appeared

    def close(self):
        if self.hwnd and self.stopped:           # asked to close: give Edge a moment to exit
            try:
                self.proc.wait(timeout=2)
            except subprocess.TimeoutExpired:
                pass
        if self.proc.poll() is None:             # still running: end Edge and its helpers
            subprocess.run(["taskkill", "/PID", str(self.proc.pid), "/T", "/F"],
                           capture_output=True, creationflags=CREATE_NO_WINDOW, timeout=10)


# ---------- app ----------
class KioskLauncher:
    def __init__(self, root):
        self.root = root
        self.settings = load_settings()
        self.folder = resolve_folder(self.settings)
        self.show = None              # PowerPointShow / VideoShow currently playing
        self.show_active = False      # also True while a show is starting
        self.show_shown = False       # its window has appeared
        self.show_stopping = False    # we asked it to stop
        self.show_failed = False      # it reported an error
        self.back_win = None          # "Back to menu" button shown over the show
        self.back_job = None          # its fade timer
        self.back_alpha = 0.0
        self.back_visible = False
        self.back_last_input = 0      # last Windows input tick seen
        self.back_active_at = 0.0     # when the button was last woken by input
        self.settings_win = None
        self.settings_key_latched = False
        self.esc_latched = False
        self.corner_step = 0          # progress through the 4-corner tap pattern
        self.corner_started = 0.0
        self.popup = None
        self.popup_timer = None
        self.thumb_labels = {}        # path -> preview Label on the current cards
        self.thumb_images = {}        # path -> PhotoImage (Tk needs a live reference)
        self.pending_thumbs = []
        self.thumb_job = None
        self.can_scroll = False

        # Preview size: fill one column of the grid at this screen width, 16:9 placeholder.
        usable = (root.winfo_screenwidth() - 2 * SIDE_PAD - SCROLLBAR_W
                  - COLUMNS * 2 * CARD_GAP)
        self.thumb_w = max(200, usable // COLUMNS - 2 * CARD_PAD)
        self.thumb_h = round(self.thumb_w * 9 / 16)
        self.blank = tk.PhotoImage(width=1, height=1)   # lets Labels be sized in pixels

        root.title(f"Museum Kiosk v{VERSION}")
        root.configure(bg=BG_COLOR)
        root.attributes("-fullscreen", True)
        root.protocol("WM_DELETE_WINDOW", lambda: None)   # blocks close button / Alt+F4

        self.build_ui()
        self.load_presentations()
        root.bind_all("<ButtonPress-1>", self.on_corner_tap, add="+")
        root.after(200, self.poll_hotkeys)
        if self.folder is None:
            root.after(300, self.choose_folder)           # first run

    # ----- UI -----
    def build_ui(self):
        tk.Label(self.root, text=TITLE_TEXT, font=("Segoe UI", 34, "bold"),
                 fg=TEXT_COLOR, bg=BG_COLOR).pack(pady=(40, 0))
        tk.Label(self.root, text=SUBTITLE_TEXT, font=("Segoe UI", 16),
                 fg=MUTED_COLOR, bg=BG_COLOR).pack(pady=(5, 20))

        container = tk.Frame(self.root, bg=BG_COLOR)
        container.pack(fill="both", expand=True, padx=SIDE_PAD, pady=(0, 30))
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        # yscrollincrement=1 -> one scroll unit is one pixel (natural drag-scroll)
        self.canvas = tk.Canvas(container, bg=BG_COLOR, highlightthickness=0,
                                yscrollincrement=1)
        self.scrollbar = tk.Scrollbar(container, orient="vertical",
                                      command=self.canvas.yview, width=SCROLLBAR_W)
        self.grid = tk.Frame(self.canvas, bg=BG_COLOR)
        self.grid_window = self.canvas.create_window((0, 0), window=self.grid, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.scrollbar.grid_remove()                      # shown only when needed

        self.grid.bind("<Configure>", lambda e: self.update_scroll())
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        self.canvas.bind_all("<MouseWheel>", lambda e: self.scroll(int(-e.delta / 120) * 60))

        # Version, bottom-right corner (created after the grid so it sits on top).
        tk.Label(self.root, text=f"v{VERSION}", font=("Segoe UI", 10),
                 fg=MUTED_COLOR, bg=BG_COLOR).place(relx=1.0, rely=1.0, anchor="se",
                                                    x=-12, y=-8)

    def on_canvas_resize(self, event):
        self.canvas.itemconfig(self.grid_window, width=event.width)
        self.update_scroll()

    def update_scroll(self):
        """Show the scrollbar (and allow scrolling) only if the cards don't fit."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.can_scroll = self.grid.winfo_reqheight() > self.canvas.winfo_height()
        if self.can_scroll:
            self.scrollbar.grid()
        else:
            self.scrollbar.grid_remove()
            self.canvas.yview_moveto(0)

    def scroll(self, pixels):
        if self.can_scroll:
            self.canvas.yview_scroll(pixels, "units")

    def list_presentations(self):
        if self.folder is None or not self.folder.is_dir():
            return None
        files = [f for f in self.folder.iterdir()
                 if f.is_file() and f.suffix.lower() in EXTENSIONS
                 and not f.name.startswith("~$")]          # skip Office lock files
        return sorted(files, key=lambda f: f.name.lower())

    def load_presentations(self):
        for widget in self.grid.winfo_children():
            widget.destroy()
        self.thumb_labels.clear()
        self.thumb_images.clear()
        self.pending_thumbs = []
        self.canvas.yview_moveto(0)

        if self.folder is None:
            self.message("No presentation folder has been set up.", ACCENT_COLOR)
            return
        try:
            files = self.list_presentations()
        except OSError:
            logging.exception("Could not read presentations folder")
            files = None

        if files is None:
            self.message(f"Presentations folder not found:\n{self.folder}", ACCENT_COLOR)
            return
        if not files:
            self.message("No presentations available.", MUTED_COLOR)
            return

        keep = set()
        for i, path in enumerate(files):
            unblock(path)
            self.make_card(path, i // COLUMNS, i % COLUMNS)
            png = self.thumb_path(path)
            if png is None:
                continue
            keep.add(png.name)
            if png.exists():
                self.set_thumb(path, png)
            else:
                self.pending_thumbs.append((path, png))
        for c in range(COLUMNS):
            self.grid.columnconfigure(c, weight=1, uniform="col")

        self.prune_thumbs(keep)
        if self.pending_thumbs and self.thumb_job is None:
            self.thumb_job = self.root.after(300, self.render_next_thumb)

    def message(self, text, color):
        tk.Label(self.grid, text=text, font=("Segoe UI", 18), fg=color,
                 bg=BG_COLOR, justify="center").pack(pady=40)

    def make_card(self, path, row, col):
        card = tk.Frame(self.grid, bg=CARD_COLOR, highlightbackground=ACCENT_COLOR,
                        highlightthickness=0, cursor="hand2")
        card.grid(row=row, column=col, padx=CARD_GAP, pady=CARD_GAP, sticky="nsew")

        thumb = tk.Label(card, image=self.blank, compound="center", text="Loading preview…",
                         width=self.thumb_w, height=self.thumb_h, font=("Segoe UI", 12),
                         fg=MUTED_COLOR, bg=BG_COLOR, bd=0)
        thumb.pack(padx=CARD_PAD, pady=(CARD_PAD, 8))
        title = tk.Label(card, text=display_name(path), font=("Segoe UI", 18, "bold"),
                         fg=TEXT_COLOR, bg=CARD_COLOR, wraplength=self.thumb_w,
                         justify="center")
        title.pack(padx=CARD_PAD)
        hint = tk.Label(card, text="TAP TO PLAY" if is_video(path) else "TAP TO START",
                        font=("Segoe UI", 11, "bold"), fg=ACCENT_COLOR, bg=CARD_COLOR)
        hint.pack(pady=(2, 12))
        self.thumb_labels[path] = thumb

        for widget in (card, thumb, title, hint):
            self.bind_tap(widget, card, lambda p=path: self.launch(p))

    def bind_tap(self, widget, card, on_tap):
        """Launch on release only; a drag scrolls the list instead."""
        state = {"y0": 0, "last": 0, "moved": False}

        def press(e):
            state.update(y0=e.y_root, last=e.y_root, moved=False)
            card.config(highlightthickness=3)

        def motion(e):
            dy, state["last"] = e.y_root - state["last"], e.y_root
            if abs(e.y_root - state["y0"]) > TAP_TOLERANCE_PX:
                state["moved"] = True
                card.config(highlightthickness=0)
            if state["moved"]:
                self.scroll(-dy)

        def release(e):
            card.config(highlightthickness=0)
            if not state["moved"]:
                on_tap()

        widget.bind("<ButtonPress-1>", press)
        widget.bind("<B1-Motion>", motion)
        widget.bind("<ButtonRelease-1>", release)

    # ----- previews -----
    def thumb_path(self, path):
        """Cache file name changes whenever the file (or the preview size) changes."""
        try:
            st = path.stat()
        except OSError:
            return None
        key = f"{path.resolve()}|{st.st_mtime_ns}|{st.st_size}|{self.thumb_w}"
        return THUMB_DIR / (hashlib.sha1(key.encode()).hexdigest() + ".png")

    def set_thumb(self, path, png):
        label = self.thumb_labels.get(path)
        if label is None:
            return
        try:
            image = tk.PhotoImage(file=str(png))
        except tk.TclError:
            logging.warning("Bad preview file %s - will rebuild", png.name)
            png.unlink(missing_ok=True)
            label.config(text="No preview")
            return
        self.thumb_images[path] = image
        label.config(image=image, text="", width=image.width(), height=image.height())

    def render_next_thumb(self):
        """Builds one missing preview per call, on the UI thread, so COM stays simple."""
        self.thumb_job = None
        if not self.pending_thumbs:
            quit_powerpoint_if_idle()
            return
        if self.show_active:
            self.thumb_job = self.root.after(2000, self.render_next_thumb)
            return
        path, png = self.pending_thumbs.pop(0)
        video = is_video(path)
        try:
            (export_video_thumbnail if video else export_deck_thumbnail)(path, png, self.thumb_w)
            self.set_thumb(path, png)
        except Exception:
            logging.exception("Preview failed: %s", path.name)
            label = self.thumb_labels.get(path)
            if label is not None:
                if video:
                    label.config(text="▶  Video", font=("Segoe UI", 20, "bold"))
                else:
                    label.config(text="No preview")
        self.thumb_job = self.root.after(50, self.render_next_thumb)

    def prune_thumbs(self, keep):
        try:
            for f in THUMB_DIR.glob("*.png"):
                if f.name not in keep:
                    f.unlink(missing_ok=True)
        except OSError:
            pass

    # ----- popup -----
    def show_popup(self, text, auto_close=False):
        self.hide_popup()
        box = tk.Frame(self.root, bg=CARD_COLOR, highlightbackground=ACCENT_COLOR,
                       highlightthickness=3, padx=50, pady=35)
        tk.Label(box, text=text, font=("Segoe UI", 20, "bold"), fg=TEXT_COLOR,
                 bg=CARD_COLOR, justify="center").pack()
        if auto_close:
            tk.Label(box, text="Tap to close", font=("Segoe UI", 11), fg=MUTED_COLOR,
                     bg=CARD_COLOR).pack(pady=(15, 0))
            for widget in (box, *box.winfo_children()):
                widget.bind("<ButtonRelease-1>", lambda e: self.hide_popup())
            self.popup_timer = self.root.after(POPUP_SECONDS * 1000, self.hide_popup)
        box.place(relx=0.5, rely=0.5, anchor="center")
        self.popup = box

    def hide_popup(self):
        if self.popup_timer is not None:
            self.root.after_cancel(self.popup_timer)
            self.popup_timer = None
        if self.popup is not None:
            self.popup.destroy()
            self.popup = None

    def show_failure(self, path):
        what = "video couldn't be played" if is_video(path) else "presentation couldn't be started"
        self.show_popup(f"Sorry, this {what}.\nPlease let a staff member know.", auto_close=True)

    # ----- running a show / video -----
    def launch(self, path):
        if self.show_active or self.settings_win is not None:
            return
        self.show_active = True
        unblock(path)
        self.show_popup(f"Starting {display_name(path)}…")
        self.root.update()        # paint the popup before PowerPoint blocks the UI

        loop = bool(self.settings.get("loop", True))
        try:
            self.root.attributes("-topmost", False)
            ctypes.windll.user32.AllowSetForegroundWindow(-1)   # ASFW_ANY
            self.show = VideoShow(path, loop) if is_video(path) else PowerPointShow(path, loop)
        except Exception:
            logging.exception("Could not start %s", path.name)
            self.root.update()    # swallow taps queued while we were blocked
            self.show_active = False
            self.show_failure(path)
            return

        logging.info("Started %s (loop=%s)", path.name, loop)
        self.show_shown = self.show_stopping = self.show_failed = False
        self.root.after(50, self.watch_show)

    def watch_show(self):
        show = self.show
        if not self.show_active or show is None:
            return
        if not show.is_open():                    # Esc, Back to menu, or it ended itself
            if self.show_shown or self.show_stopping:
                self.end_show("closed")
            else:
                self.end_show("window never appeared", failed=True)
            return
        if not self.show_shown and show.appeared():
            self.show_shown = True
            self.hide_popup()     # playing; the menu underneath should be clean on return
            self.show_back_button()
            if show.hwnd:
                try:
                    ctypes.windll.user32.SetForegroundWindow(ctypes.c_void_p(show.hwnd))
                except Exception:
                    pass
        self.keep_back_button_on_top()
        if self.show_shown:
            state = show.poll()
            if state == "error":
                self.show_failed = True
                self.stop_show("playback error")
            elif state == "ended":
                self.stop_show("reached the end")   # loop off: back to the menu
        self.root.after(SHOW_POLL_MS, self.watch_show)

    def stop_show(self, reason):
        """Ask the show/video to end; watch_show sees its window close and cleans up."""
        show = self.show
        if not self.show_active or show is None or not show.is_open():
            return
        logging.info("Stopping %s: %s", show.path.name, reason)
        self.show_stopping = True
        show.stop()

    def close_show(self):
        self.hide_back_button()
        show, self.show = self.show, None
        self.show_active = self.show_shown = self.show_stopping = self.show_failed = False
        if show is not None:
            try:
                show.close()
            except Exception:
                logging.exception("Could not close %s", show.path.name)

    def end_show(self, reason, failed=False):
        show = self.show
        failed = failed or self.show_failed
        logging.log(logging.WARNING if failed else logging.INFO, "Show ended: %s", reason)
        self.close_show()
        self.hide_popup()
        self.load_presentations()        # pick up any files added while away
        self.bring_to_front()
        if failed and show is not None:
            self.show_failure(show.path)

    def show_back_button(self):
        """Small always-on-top button over the show so touch users can leave it.
        It fades out after a few idle seconds; any touch, click or key brings it back."""
        self.hide_back_button()
        win = tk.Toplevel(self.root, bg="#000000")
        win.overrideredirect(True)                 # no title bar, no taskbar entry
        win.attributes("-topmost", True)
        win.attributes("-alpha", BACK_BUTTON_ALPHA)
        label = tk.Label(win, text=BACK_BUTTON_TEXT, font=("Segoe UI", 16, "bold"),
                         fg=TEXT_COLOR, bg="#000000", padx=22, pady=12, cursor="hand2")
        label.pack()
        for widget in (win, label):
            widget.bind("<ButtonRelease-1>", lambda e: self.stop_show("back button"))
        win.update_idletasks()
        x = self.root.winfo_screenwidth() - win.winfo_reqwidth() - 24
        win.geometry(f"+{x}+24")
        self.back_win = win
        self.back_alpha = BACK_BUTTON_ALPHA
        self.back_visible = True
        self.back_last_input = last_input_tick()
        self.back_active_at = time.monotonic()
        self.back_job = self.root.after(FADE_TICK_MS, self.update_back_button)

    def update_back_button(self):
        """Runs every FADE_TICK_MS while a show is up: fade in on activity, out when idle."""
        self.back_job = None
        win = self.back_win
        if win is None:
            return
        now = time.monotonic()
        tick = last_input_tick()
        if tick != self.back_last_input:           # touch / click / key anywhere
            self.back_last_input = tick
            self.back_active_at = now

        idle = now - self.back_active_at >= BACK_BUTTON_IDLE_SECONDS
        target = 0.0 if idle else BACK_BUTTON_ALPHA
        step = BACK_BUTTON_ALPHA / FADE_STEPS
        if self.back_alpha < target:
            self.back_alpha = min(target, self.back_alpha + step)
        elif self.back_alpha > target:
            self.back_alpha = max(target, self.back_alpha - step)

        visible = self.back_alpha > 0
        try:
            win.attributes("-alpha", self.back_alpha)
            if visible != self.back_visible:       # fully faded out: hide so it can't be tapped
                self.back_visible = visible
                hwnd = int(win.wm_frame(), 16)
                ctypes.windll.user32.ShowWindow(ctypes.c_void_p(hwnd),
                                                SW_SHOWNOACTIVATE if visible else SW_HIDE)
                if visible:
                    self.keep_back_button_on_top()
        except Exception:
            logging.exception("Back button fade failed")
        self.back_job = self.root.after(FADE_TICK_MS, self.update_back_button)

    def keep_back_button_on_top(self):
        """Re-assert topmost without taking focus from the show (PowerPoint/Edge can cover
        it, e.g. when a video starts)."""
        win = self.back_win
        if win is None or not self.back_visible:
            return
        try:
            hwnd = int(win.wm_frame(), 16)
            ctypes.windll.user32.SetWindowPos(ctypes.c_void_p(hwnd), ctypes.c_void_p(-1),
                                              0, 0, 0, 0, SWP_NOMOVE_NOSIZE_NOACTIVATE)
        except Exception:
            pass

    def hide_back_button(self):
        if self.back_job is not None:
            self.root.after_cancel(self.back_job)
            self.back_job = None
        if self.back_win is not None:
            self.back_win.destroy()
            self.back_win = None

    def bring_to_front(self):
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-topmost", True)
        self.root.lift()
        self.root.focus_force()
        self.root.after(500, lambda: self.root.attributes("-topmost", False))

    # ----- staff: hotkeys, corner taps, settings, folder -----
    def poll_hotkeys(self):
        try:
            ctrl_alt = key_down(VK_CONTROL) and key_down(VK_MENU)
            quit_down = ctrl_alt and key_down(VK_Q)
            settings_down = ctrl_alt and key_down(VK_S)
            esc_down = key_down(VK_ESCAPE)
        except Exception:
            quit_down = settings_down = esc_down = False
        if quit_down:
            self.quit()
            return
        if settings_down and not self.settings_key_latched:
            self.open_settings()
        if esc_down and not self.esc_latched:
            self.stop_show("Esc")
        self.settings_key_latched = settings_down
        self.esc_latched = esc_down
        self.root.after(150, self.poll_hotkeys)

    def on_corner_tap(self, event):
        """Settings open after tapping top-left, top-right, bottom-right, bottom-left in order."""
        if self.show_active or self.settings_win is not None:
            return
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x, y = event.x_root - self.root.winfo_rootx(), event.y_root - self.root.winfo_rooty()
        left, right = x < CORNER_PX, x >= w - CORNER_PX
        top, bottom = y < CORNER_PX, y >= h - CORNER_PX
        corner = (0 if top and left else 1 if top and right else
                  2 if bottom and right else 3 if bottom and left else None)
        now = time.monotonic()

        if corner == 0:                                   # top-left always (re)starts
            self.corner_step, self.corner_started = 1, now
        elif (corner is not None and corner == self.corner_step
              and now - self.corner_started <= CORNER_SECONDS):
            self.corner_step += 1
            if self.corner_step == 4:
                self.corner_step = 0
                logging.info("Settings opened by corner taps")
                self.root.after_idle(self.open_settings)
        else:
            self.corner_step = 0

    def quit(self):
        logging.info("Kiosk quit")
        self.close_show()
        self.root.destroy()

    def choose_folder(self, parent=None):
        options = {"parent": parent or self.root, "mustexist": True,
                   "title": "Choose the folder that holds the presentations"}
        if self.folder is not None and self.folder.is_dir():
            options["initialdir"] = str(self.folder)
        chosen = filedialog.askdirectory(**options)
        if not chosen:
            return False
        self.folder = Path(chosen)
        self.settings["presentations_folder"] = str(self.folder)
        if not save_settings(self.settings):
            messagebox.showwarning("Settings", f"Could not save settings to\n{SETTINGS_FILE}\n\n"
                                   "The folder will be used until the kiosk restarts.",
                                   parent=parent or self.root)
        logging.info("Presentations folder set to %s", self.folder)
        self.load_presentations()
        return True

    def open_settings(self):
        if self.settings_win is not None or self.show_active:
            return
        self.hide_popup()
        self.bring_to_front()

        win = tk.Toplevel(self.root, bg=CARD_COLOR, padx=30, pady=25)
        win.title("Kiosk Settings")
        win.resizable(False, False)
        win.transient(self.root)
        self.settings_win = win

        def close():
            self.settings_win = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", close)

        def text(content, color=TEXT_COLOR, size=12, bold=False, **pack):
            label = tk.Label(win, text=content, fg=color, bg=CARD_COLOR, justify="left",
                             font=("Segoe UI", size, "bold" if bold else "normal"),
                             wraplength=560)
            label.pack(anchor="w", **pack)
            return label

        text("Kiosk Settings", size=18, bold=True)
        text(f"Version {VERSION}", color=MUTED_COLOR, size=10, pady=(0, 18))

        text("Presentations folder", bold=True)
        folder_label = text(str(self.folder) if self.folder else "(not set)",
                            color=MUTED_COLOR, pady=(0, 8))

        def change_folder():
            if self.choose_folder(parent=win):
                folder_label.config(text=str(self.folder))

        button = {"font": ("Segoe UI", 13), "padx": 16, "pady": 6}   # touch-sized
        tk.Button(win, text="Change folder…", command=change_folder, **button).pack(anchor="w")

        loop_var = tk.BooleanVar(value=bool(self.settings.get("loop", True)))

        def toggle_loop():
            self.settings["loop"] = loop_var.get()
            save_settings(self.settings)
            logging.info("Loop set to %s", self.settings["loop"])

        tk.Checkbutton(win, text="Loop presentations and videos until Esc / Back to menu\n"
                                 "(unchecked: play once, then return to the menu)",
                       variable=loop_var, command=toggle_loop, font=("Segoe UI", 12),
                       justify="left", anchor="w", fg=TEXT_COLOR, bg=CARD_COLOR,
                       selectcolor=BG_COLOR, activebackground=CARD_COLOR,
                       activeforeground=TEXT_COLOR).pack(anchor="w", pady=(22, 0))

        text("Open this panel: tap the corners top-left, top-right, bottom-right, "
             "bottom-left, or press Ctrl+Alt+S.  Ctrl+Alt+Q quits.",
             color=MUTED_COLOR, size=10, pady=(22, 0))

        def quit_kiosk():
            if messagebox.askyesno("Quit Kiosk", "Close the kiosk and return to Windows?",
                                   parent=win):
                self.quit()

        row = tk.Frame(win, bg=CARD_COLOR)
        row.pack(fill="x", pady=(14, 0))
        tk.Button(row, text="Quit Kiosk", command=quit_kiosk, fg=TEXT_COLOR, bg=ACCENT_COLOR,
                  activeforeground=TEXT_COLOR, activebackground=ACCENT_COLOR,
                  **button).pack(side="left")
        tk.Button(row, text="Close", command=close, width=8, **button).pack(side="right")

        win.update_idletasks()
        x = (self.root.winfo_screenwidth() - win.winfo_reqwidth()) // 2
        y = (self.root.winfo_screenheight() - win.winfo_reqheight()) // 3
        win.geometry(f"+{x}+{y}")
        win.grab_set()
        win.focus_force()


def main():
    # Single instance: a second launch explains itself instead of silently exiting.
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    mutex = kernel32.CreateMutexW(None, False, "Local\\MuseumKioskLauncher")
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        ctypes.windll.user32.MessageBoxW(
            None, "The kiosk is already running.\n\nPress Ctrl+Alt+Q to quit it.",
            "Museum Kiosk", 0x40 | 0x40000)          # MB_ICONINFORMATION | MB_TOPMOST
        return

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # crisp text on scaled displays
    except Exception:
        pass

    ensure_data_dir()
    setup_logging()
    logging.info("Kiosk v%s starting (app: %s, data: %s)", VERSION, APP_DIR, DATA_DIR)
    root = tk.Tk()
    KioskLauncher(root)
    root.mainloop()
    kernel32.CloseHandle(mutex)


if __name__ == "__main__":
    main()