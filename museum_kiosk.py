r"""
Museum Kiosk Launcher (MVP)
---------------------------
Fullscreen menu of the PowerPoint files in the "Presentations" folder that sits
next to this script (or next to MuseumKiosk.exe when built).
Tap/click one to start it as a slideshow. Esc ends the show and the menu returns.

Looping is controlled by each presentation's own setting:
    Slide Show > Set Up Slide Show > "Loop continuously until 'Esc'"
For an unattended loop the slides also need timings (Transitions > Advance Slide > After).

Quit:       Ctrl+Alt+Q (works during a show too). Any running show is force-closed first.
Build exe:  py -m PyInstaller --onefile --noconsole --name MuseumKiosk museum_kiosk.py
            then put dist\MuseumKiosk.exe next to the Presentations folder.

Requires PowerPoint installed and licensed. Python standard library only.
Provided as-is.
"""

# VERSION - bump on every change:
#   0.0.x  minor changes (fixes, tweaks, text/config/color changes)
#   0.x.0  major revisions (features added/removed, behavior changes) - reset the last digit to 0
# This project stays in alpha: always 0.x.x, never bump to 1.0.0.
VERSION = "0.2.0"

import ctypes
import logging
import os
import subprocess
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox


def app_dir():
    """Folder containing the .exe (when built) or this .py file.
    Used instead of the working directory, which changes depending on how the app is launched."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# ---------- CONFIGURATION ----------
APP_DIR = app_dir()
PRESENTATIONS_FOLDER = APP_DIR / "Presentations"
LOG_FILE = APP_DIR / "kiosk.log"

# None = auto-detect. Or a full path, e.g. r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
POWERPOINT_PATH = None

# Close PowerPoint after each show so visitors never land in the editor.
# Leave True on the kiosk. Closing is graceful (unsaved work prompts first).
CLOSE_POWERPOINT_AFTER_SHOW = True

TITLE_TEXT = "Choose a Presentation"
SUBTITLE_TEXT = "Tap a presentation to begin"
COLUMNS = 3

BG_COLOR = "#1a1a2e"
CARD_COLOR = "#16213e"
ACCENT_COLOR = "#e94560"
TEXT_COLOR = "#ffffff"
MUTED_COLOR = "#aaaaaa"
# -----------------------------------

TAP_TOLERANCE_PX = 12        # finger drift allowed before a tap becomes a scroll
START_GRACE_SECONDS = 20     # how long to wait for the slideshow to appear after launch
SHOW_WINDOW_CLASS = "screenClass"   # PowerPoint's fullscreen slideshow window
EXTENSIONS = {".pptx", ".ppt", ".ppsx", ".pps"}
CREATE_NO_WINDOW = 0x08000000
ERROR_ALREADY_EXISTS = 183
VK_CONTROL, VK_MENU, VK_Q = 0x11, 0x12, 0x51   # VK_MENU = Alt


# ---------- helpers ----------
def setup_logging():
    try:
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                            format="%(asctime)s %(levelname)s %(message)s")
    except OSError:
        logging.basicConfig(level=logging.INFO)


def find_powerpoint():
    """Return the path to POWERPNT.EXE, or None."""
    if POWERPOINT_PATH:
        return POWERPOINT_PATH if os.path.exists(POWERPOINT_PATH) else None

    # Registry "App Paths" covers Click-to-Run, MSI, 32- and 64-bit installs.
    import winreg
    key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ | view) as key:
                    path = winreg.QueryValue(key, None).strip('"')
                    if path and os.path.exists(path):
                        return path
            except OSError:
                pass

    # Fallback: common install locations.
    for base in (r"C:\Program Files", r"C:\Program Files (x86)"):
        for sub in (r"Microsoft Office\root\Office16", r"Microsoft Office\Office16",
                    r"Microsoft Office\root\Office15", r"Microsoft Office\Office15"):
            path = os.path.join(base, sub, "POWERPNT.EXE")
            if os.path.exists(path):
                return path
    return None


def slideshow_open():
    """True while a PowerPoint slideshow window exists (window class 'screenClass').
    Watching the show window, not the process, means an already-open PowerPoint
    doesn't stop the menu from coming back."""
    return bool(ctypes.windll.user32.FindWindowW(SHOW_WINDOW_CLASS, None))


def close_powerpoint(force=False):
    """Graceful close by default; force=True kills it outright (unsaved work is lost)."""
    cmd = ["taskkill", "/IM", "POWERPNT.EXE"]
    if force:
        cmd.insert(1, "/F")
    try:
        subprocess.run(cmd, capture_output=True, creationflags=CREATE_NO_WINDOW, timeout=10)
        logging.info("PowerPoint closed (force=%s)", force)
    except Exception:
        logging.exception("taskkill failed")


def quit_hotkey_down():
    """Ctrl+Alt+Q, checked globally so it works while PowerPoint has focus."""
    get = ctypes.windll.user32.GetAsyncKeyState
    return all(get(vk) & 0x8000 for vk in (VK_CONTROL, VK_MENU, VK_Q))


# ---------- app ----------
class KioskLauncher:
    def __init__(self, root):
        self.root = root
        self.show_active = False
        self.show_started_at = 0.0
        self.seen_show = False

        root.title(f"Museum Kiosk v{VERSION}")
        root.configure(bg=BG_COLOR)
        root.attributes("-fullscreen", True)
        root.protocol("WM_DELETE_WINDOW", lambda: None)   # blocks close button / Alt+F4
        root.bind("<Escape>", lambda e: "break")

        self.build_ui()
        self.load_presentations()
        root.after(200, self.poll_quit_hotkey)

    # ----- UI -----
    def build_ui(self):
        tk.Label(self.root, text=TITLE_TEXT, font=("Segoe UI", 34, "bold"),
                 fg=TEXT_COLOR, bg=BG_COLOR).pack(pady=(40, 0))
        tk.Label(self.root, text=SUBTITLE_TEXT, font=("Segoe UI", 16),
                 fg=MUTED_COLOR, bg=BG_COLOR).pack(pady=(5, 20))

        container = tk.Frame(self.root, bg=BG_COLOR)
        container.pack(fill="both", expand=True, padx=60, pady=(0, 30))

        # yscrollincrement=1 -> one scroll unit is one pixel (natural drag-scroll)
        self.canvas = tk.Canvas(container, bg=BG_COLOR, highlightthickness=0,
                                yscrollincrement=1)
        scrollbar = tk.Scrollbar(container, orient="vertical",
                                 command=self.canvas.yview, width=28)
        self.grid = tk.Frame(self.canvas, bg=BG_COLOR)
        window = self.canvas.create_window((0, 0), window=self.grid, anchor="nw")

        self.grid.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            window, width=e.width))
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(
            int(-e.delta / 120) * 60, "units"))

        # Version, bottom-right corner (created last so it sits on top).
        tk.Label(self.root, text=f"v{VERSION}", font=("Segoe UI", 10),
                 fg=MUTED_COLOR, bg=BG_COLOR).place(relx=1.0, rely=1.0, anchor="se",
                                                    x=-12, y=-8)

    def list_presentations(self):
        folder = PRESENTATIONS_FOLDER
        if not folder.is_dir():
            return None
        files = [f for f in folder.iterdir()
                 if f.is_file() and f.suffix.lower() in EXTENSIONS
                 and not f.name.startswith("~$")]          # skip Office lock files
        return sorted(files, key=lambda f: f.name.lower())

    def load_presentations(self):
        for widget in self.grid.winfo_children():
            widget.destroy()
        self.canvas.yview_moveto(0)

        try:
            files = self.list_presentations()
        except OSError:
            logging.exception("Could not read presentations folder")
            files = None

        if files is None:
            self.message(f"Presentations folder not found:\n{PRESENTATIONS_FOLDER}", ACCENT_COLOR)
            return
        if not files:
            self.message("No presentations available.", MUTED_COLOR)
            return

        for i, path in enumerate(files):
            self.make_card(path, i // COLUMNS, i % COLUMNS)
        for c in range(COLUMNS):
            self.grid.columnconfigure(c, weight=1, uniform="col")

    def message(self, text, color):
        tk.Label(self.grid, text=text, font=("Segoe UI", 18), fg=color,
                 bg=BG_COLOR, justify="center").pack(pady=40)

    def make_card(self, path, row, col):
        card = tk.Frame(self.grid, bg=CARD_COLOR, highlightbackground=ACCENT_COLOR,
                        highlightthickness=0, cursor="hand2")
        card.grid(row=row, column=col, padx=15, pady=15, sticky="nsew", ipady=25)

        name = path.stem.replace("_", " ").replace("-", " ")
        title = tk.Label(card, text=name, font=("Segoe UI", 18, "bold"), fg=TEXT_COLOR,
                         bg=CARD_COLOR, wraplength=300, justify="center")
        title.pack(pady=(20, 10), padx=15)
        hint = tk.Label(card, text="TAP TO START", font=("Segoe UI", 11, "bold"),
                        fg=ACCENT_COLOR, bg=CARD_COLOR)
        hint.pack(pady=(0, 15))

        for widget in (card, title, hint):
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
                self.canvas.yview_scroll(-dy, "units")

        def release(e):
            card.config(highlightthickness=0)
            if not state["moved"]:
                on_tap()

        widget.bind("<ButtonPress-1>", press)
        widget.bind("<B1-Motion>", motion)
        widget.bind("<ButtonRelease-1>", release)

    # ----- running a show -----
    def launch(self, path):
        if self.show_active:
            return
        ppt = find_powerpoint()
        if not ppt:
            logging.error("PowerPoint not found")
            messagebox.showerror("PowerPoint Not Found",
                                 "PowerPoint could not be found on this computer.\n\n"
                                 "Please contact staff.", parent=self.root)
            return
        try:
            subprocess.Popen([ppt, "/S", str(path)])        # /S = start as slideshow
        except OSError as e:
            logging.exception("Launch failed: %s", path)
            messagebox.showerror("Error", f"Could not start the presentation.\n\n{e}",
                                 parent=self.root)
            return

        logging.info("Started %s", path.name)
        self.show_active = True
        self.seen_show = False
        self.show_started_at = time.time()
        self.root.withdraw()
        self.root.after(500, self.watch_show)

    def watch_show(self):
        if not self.show_active:
            return
        if slideshow_open():
            self.seen_show = True
        elif self.seen_show:
            self.end_show("slideshow closed")
            return
        elif time.time() - self.show_started_at > START_GRACE_SECONDS:
            # Usually a PowerPoint prompt (activation, Protected View, repair) is in the way.
            logging.warning("Slideshow never appeared within %ss", START_GRACE_SECONDS)
            self.end_show("slideshow never appeared", force_close=True)
            return
        self.root.after(500, self.watch_show)

    def end_show(self, reason, force_close=False):
        logging.info("Show ended: %s", reason)
        self.show_active = False
        if CLOSE_POWERPOINT_AFTER_SHOW:
            close_powerpoint(force=force_close)
        self.load_presentations()        # pick up any files added while away
        self.bring_to_front()

    def bring_to_front(self):
        self.root.deiconify()
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-topmost", True)
        self.root.lift()
        self.root.focus_force()
        self.root.after(500, lambda: self.root.attributes("-topmost", False))

    # ----- quit -----
    def poll_quit_hotkey(self):
        try:
            down = quit_hotkey_down()
        except Exception:
            down = False
        if down:
            self.quit()
            return
        self.root.after(150, self.poll_quit_hotkey)

    def quit(self):
        logging.info("Quit hotkey pressed")
        if self.show_active:
            close_powerpoint(force=True)
            self.show_active = False
        self.root.destroy()


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

    setup_logging()
    logging.info("Kiosk v%s starting (folder: %s)", VERSION, PRESENTATIONS_FOLDER)
    root = tk.Tk()
    KioskLauncher(root)
    root.mainloop()
    kernel32.CloseHandle(mutex)


if __name__ == "__main__":
    main()