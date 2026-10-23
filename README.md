# Museum Kiosk Launcher

A fullscreen, touch-friendly menu for picking and playing PowerPoint presentations on a
Windows kiosk. Visitors tap a card (first-slide preview + title) and the deck plays as a
fullscreen slideshow. Esc or the on-screen **Back to menu** button returns to the menu.
The Windows desktop is never shown.

Bare-bones by design: one Python file, one dependency, no installer, no background
services. Provided as-is.

---

## Requirements

- Windows 10 or 11 (any edition)
- Microsoft PowerPoint, **installed, licensed and activated** (it does the actual playback)
- To build or run from source: Python 3.9+ and the packages in `requirements.txt`
  (`pywin32` for PowerPoint automation, `pyinstaller` for building the exe)

## Project files

| File | Purpose |
|---|---|
| `museum_kiosk.py` | The whole app |
| `requirements.txt` | `pywin32` (runtime), `pyinstaller` (build only) |
| `build.bat` | Builds `MuseumKiosk.exe` into the project folder |
| `Presentations/` | Test decks; used automatically when running from the project folder |

## Quick start (from source)

```bat
py -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python museum_kiosk.py
```

Exit with **Ctrl+Alt+Q**.

## Build the exe

```bat
build.bat
```

This installs the requirements (into `.venv` if it exists) and writes a single-file
`MuseumKiosk.exe` to the project folder. Rebuild after every code change.

---

## Using the kiosk

### Visitors
- **Tap a card** to start that presentation fullscreen.
- **Drag** to scroll when there are more decks than fit on screen. The scrollbar only
  appears when it's needed.
- **Back to menu** (top-right during a show) ends the show. It fades out after 3 idle
  seconds; any touch brings it back. Note that the waking tap also reaches PowerPoint and
  may advance the slide.
- **Esc** also ends the show, whichever window has focus.
- Right-click / long-press during a show opens PowerPoint's normal menu.

### Staff
| Action | Touch | Keyboard |
|---|---|---|
| Open settings | Tap the corners in order: **top-left → top-right → bottom-right → bottom-left**, within 5 seconds, on the menu screen | **Ctrl+Alt+S** (not during a show) |
| Quit the kiosk | Settings → **Quit Kiosk** | **Ctrl+Alt+Q** (works during a show too) |

The settings panel lets you:
- **Change the presentations folder**
- **Loop presentations until Esc.** On: every deck loops until someone exits. Off: every
  deck plays once, then returns to the menu. This overrides each deck's own loop setting.

If an error popup appears ("this presentation couldn't be started"), it closes by itself
after 8 seconds or when tapped. The reason is in `kiosk.log`.

---

## Preparing presentations

- Supported: `.pptx`, `.ppt`, `.ppsx`, `.pps`. Office lock files (`~$...`) are ignored.
- **Card title** = file name, with `_` and `-` shown as spaces. Cards are sorted
  alphabetically by file name.
- **Preview** = slide 1. Previews are generated once and rebuilt automatically when a file
  changes.
- **For unattended playback, give every slide a timing:** Transitions > Advance Slide >
  After. Without timings, slides only advance on a tap.
- New or changed files are picked up whenever the menu reappears (after a show ends).
  No restart is needed.
- Files downloaded or emailed are automatically **unblocked** (the "from the internet"
  mark is removed) so PowerPoint doesn't hold them in Protected View. This needs write
  access to the folder; otherwise unblock them manually:

  ```powershell
  Get-ChildItem 'D:\Kiosk Presentations' | Unblock-File
  ```

- Avoid **password-protected** decks (see Known limitations).

## Where the presentations folder comes from

The first match wins:

1. The folder saved in settings (chosen on first run, or in the settings panel)
2. A `Presentations` folder next to `MuseumKiosk.exe` / `museum_kiosk.py`
3. Neither: the app asks for a folder on startup

If a saved folder goes missing (unplugged drive, network share not ready), the menu shows
"Presentations folder not found" instead of prompting. Fix the folder, or change it in
settings.

## Data the app writes

Everything is in `C:\ProgramData\MuseumKiosk\`:

| Path | Contents |
|---|---|
| `settings.json` | `presentations_folder`, `loop` |
| `kiosk.log` | Starts, stops, errors |
| `thumbs\` | Cached first-slide previews |

When the app creates this folder, it grants the local **Users** group *Modify* on it (on
that folder only), so settings work whichever account runs the app first. To undo:
`icacls "C:\ProgramData\MuseumKiosk" /remove:g *S-1-5-32-545`

To **reset** the app: delete `settings.json` (the folder prompt returns) and/or `thumbs\`
(previews are rebuilt).

---

## Deploying to a kiosk PC

1. Install PowerPoint, sign in or activate it, and open it once by hand to clear any
   first-run prompts.
2. Copy `MuseumKiosk.exe` somewhere permanent, e.g. `C:\MuseumKiosk\`.
3. Put the presentations in a folder of your choice, e.g. `D:\Kiosk Presentations\`.
4. Create a standard local account for the kiosk and set it to sign in automatically
   (e.g. Sysinternals Autologon).
5. Signed in as that account, run `shell:startup` and put a shortcut to
   `MuseumKiosk.exe` in that folder.
6. Start the app once and choose the presentations folder when asked.
7. Lock the PC down using Windows itself. The app doesn't block the Win key, Alt+Tab
   or edge swipes:
   - Disable touch edge swipes: Group Policy *Computer Configuration > Administrative
     Templates > Windows Components > Edge UI > Allow edge swipe* = **Disabled**
     (or `HKLM\SOFTWARE\Policies\Microsoft\Windows\EdgeUI`, DWORD `AllowEdgeSwipe` = `0`).
   - Turn off sleep and the screen saver for the kiosk account.
8. If SentinelOne (or other EDR) flags the unsigned exe, add an exclusion for it.

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| "Couldn't be started" popup | Check `kiosk.log`. Usually PowerPoint isn't activated, the file is damaged, or it's an unsupported format. |
| The app freezes when a deck is tapped | PowerPoint is showing a dialog (activation, repair, password) behind the show. Fix PowerPoint or the deck, then Ctrl+Alt+Q and restart. |
| "No preview" on a card | PowerPoint couldn't open the deck to render slide 1. Check `kiosk.log`. |
| A deck doesn't advance on its own | Its slides have no timings (see Preparing presentations). |
| A deck ends instead of looping | Loop is off in settings. |
| "Presentations folder not found" | The saved folder is missing. Reconnect it, or change it in settings. |
| "The kiosk is already running" | Another copy is open. Ctrl+Alt+Q quits it. |

## Known limitations

- A PowerPoint dialog that appears while a deck is opening (activation, password, file
  repair) stops the app until it's dealt with. Use activated PowerPoint and decks
  without passwords.
- The tap that brings the Back to menu button back is also delivered to PowerPoint.
- Single monitor assumed. Presenter View is turned off for every show.
- The app quits PowerPoint after each show if nothing else is open in it.

## Development notes

- **Versioning:** `VERSION` at the top of `museum_kiosk.py`. Bump `0.0.x` for minor
  changes, `0.x.0` for features and behaviour changes. The project stays at `0.x.x`.
- Appearance and timing settings (colours, columns, title text, button text, fade
  timing, corner size) are constants near the top of `museum_kiosk.py`.
- PowerPoint is driven over COM (`pywin32`). Decks are always opened read-only and
  windowless, and never saved.
