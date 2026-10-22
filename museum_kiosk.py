r"""
Museum Kiosk Launcher (MVP)
---------------------------
Fullscreen menu of the PowerPoint files in the presentations folder, each shown as a
card with a preview of its first slide. Tap one to start it as a slideshow.
Esc ends the show and the menu returns.

Presentations folder (first match wins):
    1. The folder saved in settings (chosen on first run, or in the settings panel)
    2. A "Presentations" folder next to this script / MuseumKiosk.exe
    3. Neither -> the app asks for a folder on startup

Settings, log and preview cache:  C:\ProgramData\MuseumKiosk\
    settings.json   presentations_folder, loop
    kiosk.log
    thumbs\         first-slide previews (rebuilt automatically when a deck changes)

Staff hotkeys (work during a show too):
    Ctrl+Alt+S   settings panel: presentations folder, loop on/off   (not during a show)
    Ctrl+Alt+Q   quit the kiosk (any running show is closed first)

Loop setting (applies to every deck, overrides the deck's own setting):
    on   loop until Esc
    off  play once, then return to the menu
Unattended advancing still needs slide timings (Transitions > Advance Slide > After).

Run from source:  py -m pip install -r requirements.txt   then   py museum_kiosk.py
Build exe:        build.bat   (puts MuseumKiosk.exe in the project folder)

Requires PowerPoint installed and licensed. Provided as-is.
"""

# VERSION - bump on every change:
#   0.0.x  minor changes (fixes, tweaks, text/config/color changes)
#   0.x.0  major revisions (features added/removed, behavior changes) - reset the last digit to 0
# This project stays in alpha: always 0.x.x, never bump to 1.0.0.
VERSION = "0.3.0"

import ctypes
import hashlib
import json
import logging
import os
import subprocess
import sys
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
DEFAULT_SETTINGS = {"presentations_folder": None, "loop": True}

TITLE_TEXT = "Choose a Presentation"
SUBTITLE_TEXT = "Tap a presentation to begin"
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
EXTENSIONS = {".pptx", ".ppt", ".ppsx", ".pps"}
CREATE_NO_WINDOW = 0x08000000
ERROR_ALREADY_EXISTS = 183
VK_CONTROL, VK_MENU, VK_Q, VK_S = 0x11, 0x12, 0x51, 0x53   # VK_MENU = Alt

# PowerPoint COM constants
MSO_TRUE, MSO_FALSE = -1, 0
PP_SHOW_TYPE_SPEAKER = 1     # normal full-screen show (even if the deck is set to "window")
PP_SLIDESHOW_DONE = 5        # View.State once a non-looping show reaches its end


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


def display_name(path):
    return path.stem.replace("_", " ").replace("-", " ")


def key_down(vk):
    return bool(ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000)


# ---------- PowerPoint (COM) ----------
def powerpoint():
    """PowerPoint is single-instance: this starts it hidden, or attaches to the running one."""
    return win32com.client.Dispatch("PowerPoint.Application")


def open_readonly(app, path):
    # Open(FileName, ReadOnly, Untitled, WithWindow): no editor window is ever shown.
    return app.Presentations.Open(str(path), MSO_TRUE, MSO_FALSE, MSO_FALSE)


def quit_powerpoint_if_idle():
    """Quit PowerPoint only if nothing else is open in it."""
    try:
        app = powerpoint()
        if app.Presentations.Count == 0:
            app.Quit()
    except Exception:
        logging.exception("Could not quit PowerPoint")


def export_thumbnail(path, png, width):
    pres = open_readonly(powerpoint(), path)
    try:
        ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
        pres.Slides.Item(1).Export(str(png), "PNG", width, round(width * ratio))
    finally:
        pres.Close()


# ---------- app ----------
class KioskLauncher:
    def __init__(self, root):
        self.root = root
        self.settings = load_settings()
        self.folder = resolve_folder(self.settings)
        self.pres = None              # presentation currently running as a show
        self.show_active = False      # also True while a show is starting
        self.settings_win = None
        self.settings_key_latched = False
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
        hint = tk.Label(card, text="TAP TO START", font=("Segoe UI", 11, "bold"),
                        fg=ACCENT_COLOR, bg=CARD_COLOR)
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
        """Cache file name changes whenever the deck (or the preview size) changes."""
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
        try:
            export_thumbnail(path, png, self.thumb_w)
            self.set_thumb(path, png)
        except Exception:
            logging.exception("Preview failed: %s", path.name)
            label = self.thumb_labels.get(path)
            if label is not None:
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

    # ----- running a show -----
    def launch(self, path):
        if self.show_active or self.settings_win is not None:
            return
        self.show_active = True
        unblock(path)
        self.show_popup(f"Starting {display_name(path)}…")
        self.root.update()        # paint the popup before PowerPoint blocks the UI

        pres = None
        try:
            pres = open_readonly(powerpoint(), path)
            show = pres.SlideShowSettings
            show.ShowType = PP_SHOW_TYPE_SPEAKER
            show.LoopUntilStopped = MSO_TRUE if self.settings.get("loop", True) else MSO_FALSE
            try:
                show.ShowPresenterView = MSO_FALSE   # no presenter view on multi-monitor PCs
            except Exception:
                pass
            self.root.attributes("-topmost", False)
            ctypes.windll.user32.AllowSetForegroundWindow(-1)   # ASFW_ANY
            window = show.Run()
        except Exception:
            logging.exception("Could not start %s", path.name)
            if pres is not None:
                try:
                    pres.Close()
                except Exception:
                    pass
            self.root.update()    # swallow taps queued while we were blocked
            self.show_active = False
            self.show_popup("Sorry, this presentation couldn't be started.\n"
                            "Please let a staff member know.", auto_close=True)
            return

        try:
            ctypes.windll.user32.SetForegroundWindow(window.HWND)
        except Exception:
            pass
        logging.info("Started %s (loop=%s)", path.name, self.settings.get("loop", True))
        self.pres = pres
        self.root.after(500, self.watch_show)

    def watch_show(self):
        if not self.show_active or self.pres is None:
            return
        try:
            view = self.pres.SlideShowWindow.View
            if view.State == PP_SLIDESHOW_DONE:   # non-looping show reached the end
                view.Exit()
                self.end_show("finished")
                return
        except Exception:                         # show window gone: Esc, or it ended itself
            self.end_show("closed")
            return
        self.root.after(500, self.watch_show)

    def close_show(self):
        pres, self.pres = self.pres, None
        self.show_active = False
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                logging.exception("Could not close presentation")
            quit_powerpoint_if_idle()

    def end_show(self, reason):
        logging.info("Show ended: %s", reason)
        self.close_show()
        self.hide_popup()
        self.load_presentations()        # pick up any files added while away
        self.bring_to_front()

    def bring_to_front(self):
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-topmost", True)
        self.root.lift()
        self.root.focus_force()
        self.root.after(500, lambda: self.root.attributes("-topmost", False))

    # ----- staff: hotkeys, settings, folder -----
    def poll_hotkeys(self):
        try:
            ctrl_alt = key_down(VK_CONTROL) and key_down(VK_MENU)
            quit_down = ctrl_alt and key_down(VK_Q)
            settings_down = ctrl_alt and key_down(VK_S)
        except Exception:
            quit_down = settings_down = False
        if quit_down:
            self.quit()
            return
        if settings_down and not self.settings_key_latched:
            self.open_settings()
        self.settings_key_latched = settings_down
        self.root.after(150, self.poll_hotkeys)

    def quit(self):
        logging.info("Quit hotkey pressed")
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

        tk.Button(win, text="Change folder…", font=("Segoe UI", 12),
                  command=change_folder).pack(anchor="w")

        loop_var = tk.BooleanVar(value=bool(self.settings.get("loop", True)))

        def toggle_loop():
            self.settings["loop"] = loop_var.get()
            save_settings(self.settings)
            logging.info("Loop set to %s", self.settings["loop"])

        tk.Checkbutton(win, text="Loop presentations until Esc\n"
                                 "(unchecked: play once, then return to the menu)",
                       variable=loop_var, command=toggle_loop, font=("Segoe UI", 12),
                       justify="left", anchor="w", fg=TEXT_COLOR, bg=CARD_COLOR,
                       selectcolor=BG_COLOR, activebackground=CARD_COLOR,
                       activeforeground=TEXT_COLOR).pack(anchor="w", pady=(22, 0))

        text("Ctrl+Alt+Q quits the kiosk.", color=MUTED_COLOR, size=10, pady=(22, 0))
        tk.Button(win, text="Close", font=("Segoe UI", 12), width=10,
                  command=close).pack(anchor="e", pady=(10, 0))

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
