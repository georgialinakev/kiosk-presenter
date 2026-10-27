# Museum Kiosk Launcher

A fullscreen, touch-friendly menu for picking and playing **PowerPoint presentations and
videos** on a Windows kiosk. Visitors tap a card (preview + title) and it plays full screen:
decks as a PowerPoint slideshow, videos in a full-screen player with touch media controls.
Esc or the on-screen **Back to menu** button returns to the menu. The Windows desktop is
never shown.

Bare-bones by design: one Python file, one dependency, no installer, no background
services. Provided as-is.

---

## Requirements

- Windows 10 or 11 (any edition)
- Microsoft PowerPoint, **installed, licensed and activated**. It plays the decks and makes
  every preview, including the video previews.
- Microsoft Edge (included with Windows 10/11). It plays the videos.
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
- **Tap a card** to play it full screen. Presentation cards say *Tap to start*; video cards
  say *Tap to play*.
- **Drag** to scroll when there are more items than fit on screen. The scrollbar only
  appears when it's needed.
- **Videos:** the media controls (play/pause, seek bar, time, volume) show when the screen
  is touched or clicked, and hide by themselves after a few seconds of playback.
- **Back to menu** (top-right) ends the presentation or video. It fades out after 3 idle
  seconds; any touch brings it back. The waking tap also reaches the slideshow or video,
  so it may advance a slide or show the video controls.
- **Esc** also ends playback, whichever window has focus.
- Right-click / long-press during a slideshow opens PowerPoint's normal menu.

### Staff
| Action | Touch | Keyboard |
|---|---|---|
| Open settings | Tap the corners in order: **top-left → top-right → bottom-right → bottom-left**, within 5 seconds, on the menu screen | **Ctrl+Alt+S** (not during playback) |
| Quit the kiosk | Settings → **Quit Kiosk** | **Ctrl+Alt+Q** (works during playback too) |

The settings panel lets you:

- **Change the presentations folder**
- **Loop presentations and videos.** On: everything loops until someone exits. Off:
  everything plays once, then returns to the menu. This overrides each deck's own loop
  setting.

If an error popup appears ("couldn't be started" / "couldn't be played"), it closes by
itself after 8 seconds or when tapped. The reason is in `kiosk.log`.

---

## Preparing content

Presentations and videos go in the **same folder**.

| Type | Extensions |
|---|---|
| Presentations | `.pptx`, `.ppt`, `.ppsx`, `.pps` |
| Videos | `.mp4`, `.m4v`, `.mov`, `.webm` |

- **Videos:** MP4 with H.264 video and AAC audio is the safest choice. Formats Edge can't
  play (e.g. `.wmv`, `.avi`, `.mkv`) aren't listed. Convert them to MP4 first.
- **Card title** = file name, with `_` and `-` shown as spaces. Cards are sorted
  alphabetically by file name. Prefix names with numbers (`01 Intro.mp4`) to control
  the order.
- **Preview** = slide 1 for decks, a frame about 2 seconds in for videos. Previews are made
  once and rebuilt automatically when a file changes. If PowerPoint can't open a video,
  its card shows **▶ Video** instead of a frame; the video still plays.
- **For unattended slideshows, give every slide a timing:** Transitions > Advance Slide >
  After. Without timings, slides only advance on a tap.
- New or changed files are picked up whenever the menu reappears (after playback ends).
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
| `thumbs\` | Cached previews |
| `player.html` | The video player page (rewritten for each video) |
| `edge-profile\` | Edge's data for the video player, kept separate from any user's Edge |

When the app creates this folder, it grants the local **Users** group *Modify* on it (on
that folder only), so settings work whichever account runs the app first. To undo:
`icacls "C:\ProgramData\MuseumKiosk" /remove:g *S-1-5-32-545`

To **reset** the app: delete `settings.json` (the folder prompt returns), `thumbs\`
(previews are rebuilt) and/or `edge-profile\` (Edge starts fresh).

---

## Deploying to a kiosk PC

1. Install PowerPoint, sign in or activate it, and open it once by hand to clear any
   first-run prompts.
2. Copy `MuseumKiosk.exe` somewhere permanent, e.g. `C:\MuseumKiosk\`.
3. Put the presentations and videos in a folder of your choice, e.g. `D:\Kiosk Presentations\`.
4. Create a standard local account for the kiosk and set it to sign in automatically
   (e.g. Sysinternals Autologon).
5. Signed in as that account, run `shell:startup` and put a shortcut to
   `MuseumKiosk.exe` in that folder.
6. Start the app once and choose the presentations folder when asked. If there are videos,
   play one to check that Edge opens straight into the player with sound.
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
| "Presentation couldn't be started" | Check `kiosk.log`. Usually PowerPoint isn't activated, the file is damaged, or it's an unsupported format. |
| "Video couldn't be played" | The video format isn't supported by Edge (convert to MP4 H.264/AAC), the file is damaged, or Edge is missing or blocked by policy. Check `kiosk.log`. |
| Video starts muted | Edge blocked autoplay with sound and the player fell back to muted. Tap the volume icon. Check that no Edge policy overrides autoplay. |
| The app freezes when a deck is tapped | PowerPoint is showing a dialog (activation, repair, password) behind the show. Fix PowerPoint or the deck, then Ctrl+Alt+Q and restart. |
| "No preview" on a card | PowerPoint couldn't open the deck to render slide 1. Check `kiosk.log`. |
| "▶ Video" instead of a preview | PowerPoint couldn't open the video for a preview frame. Playback is unaffected. |
| A deck doesn't advance on its own | Its slides have no timings (see Preparing content). |
| A deck or video ends instead of looping | Loop is off in settings. |
| "Presentations folder not found" | The saved folder is missing. Reconnect it, or change it in settings. |
| "The kiosk is already running" | Another copy is open. Ctrl+Alt+Q quits it. |

## Known limitations

- A PowerPoint dialog that appears while a deck is opening (activation, password, file
  repair) stops the app until it's dealt with. Use activated PowerPoint and decks
  without passwords.
- The tap that brings the Back to menu button back is also delivered to the slideshow or
  video.
- How quickly the video controls hide is decided by Edge (a few seconds) and isn't
  configurable.
- Single monitor assumed. Presenter View is turned off for every show.
- The app quits PowerPoint after each show if nothing else is open in it.

## Development notes

- **Versioning:** `VERSION` at the top of `museum_kiosk.py`. Bump `0.0.x` for minor
  changes, `0.x.0` for features and behaviour changes. The project stays at `0.x.x`.
- Appearance and timing settings (colours, columns, title text, button text, fade
  timing, corner size, video formats) are constants near the top of `museum_kiosk.py`.
- PowerPoint is driven over COM (`pywin32`). Decks are always opened read-only and
  windowless, and never saved.
- Videos play in Edge kiosk mode (`--kiosk --edge-kiosk-type=fullscreen --inprivate`,
  separate `--user-data-dir`) on a generated HTML5 page (`PLAYER_HTML` in the script).
  The page reports "ended" or "error" by changing its window title, which the app watches.