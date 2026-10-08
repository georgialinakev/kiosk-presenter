# Kiosk Presenter

Touch kiosk that lets visitors pick a PowerPoint presentation from a library and play it, then
returns to the library. It takes over the PC while it runs and only an admin with the PIN can
exit. Windows 10 22H2 or 11, any edition; needs licensed desktop PowerPoint.

- **UI:** `ui/`, the Claude Design frontend, running in WebView2 from local files.
- **Host:** `src/KioskPresenter.Host`, a .NET 10 WPF app: fullscreen window, bridge to the UI,
  native control strip over the slideshow, keyboard hook and focus guard, watchdog (`--watchdog`).
- **Playback:** `src/KioskPresenter.PowerPoint` drives the installed PowerPoint over COM.
- **Core:** `src/KioskPresenter.Core`: config, library scan straight from the .pptx zip, PIN hashing, bridge messages. Unit-tested.
- **Setup:** `src/KioskPresenter.Setup` builds `KioskPresenter-Setup-X.Y.Z.exe` (install, upgrade, uninstall).

Design brief: [docs/design-brief.md](docs/design-brief.md). The untouched Claude Design export is in
`docs/design/original/`.

## Develop (Windows)

Needs the .NET 10 SDK, PowerPoint and the WebView2 runtime (already on Windows 11).

```powershell
dotnet test tests/KioskPresenter.Core.Tests

# Put the demo decks where the kiosk looks for them
New-Item -ItemType Directory C:\ProgramData\KioskPresenter\Content -Force | Out-Null
Copy-Item tests\fixtures\*.pptx C:\ProgramData\KioskPresenter\Content\

# Normal window, no hooks, no watchdog, dev tools on (F12)
dotnet run --project src/KioskPresenter.Host -- --windowed --devtools

# Real kiosk behaviour without the watchdog (use a VM). Leave with Ctrl+Shift+Q, PIN, Exit kiosk.
dotnet run --project src/KioskPresenter.Host -- --no-watchdog
```

`--data C:\some\folder` points everything (config, content, logs) somewhere other than
`C:\ProgramData\KioskPresenter`.

The UI also runs in any browser with mock data: `python -m http.server 8766 --directory ui`, then
open <http://localhost:8766/index.html>. Press the backtick key for the state switcher.

## Release

Push a tag: `git tag v0.1.0 && git push origin v0.1.0`. The Release workflow builds
`KioskPresenter-Setup-0.1.0.exe` with the app, the UI, the demo decks and Microsoft's WebView2
bootstrapper inside, and attaches it and its SHA-256 to a GitHub release. Run the workflow by hand
for a test build without a release.

## Install on a kiosk

See [docs/install-guide.md](docs/install-guide.md). Short version: run the Setup exe as an admin,
pick the user the kiosk starts for, reboot or sign in as that user, set the PIN on first run.

## Configuration

`C:\ProgramData\KioskPresenter\config.json` (written on first run; edit, then Reload library in the admin panel):

| Key | Default | Meaning |
| --- | --- | --- |
| `contentFolder` | `null` (`…\KioskPresenter\Content`) | Folder of `.pptx`, `.ppsx`, `.ppt` |
| `brand.title`, `brand.subtitle` | "Welcome", "Choose a presentation to begin" | Library header |
| `brand.logo` | `null` | File name inside `…\KioskPresenter\Brand` |
| `brand.accent` | `null` | Hex colour such as `#4fd1b5` |
| `timers.*` | see `KioskConfig.cs` | Seconds: idle, prompt, strip hide and so on |
| `motion` | `true` | UI fades and scales |
| `pinHash` | set on first run | Delete the line to force PIN setup again |

Titles come from file names: `03_Age_of_Steam.pptx` shows as "Age of Steam" and sorts third.
Password-protected decks are skipped and listed in the admin panel.

Logs: `C:\ProgramData\KioskPresenter\Logs\kiosk-YYYYMMDD.log`, 14 days kept.

## Updating the UI from Claude Design

Copy the new export over `ui/`, rename `Kiosk App.dc.html` to `index.html`, then run
`python tools/patch-ui.py ui`. It reapplies every host edit (marked `HOST:`) and fails loudly if
the export changed under one of them. Then run the UI tests:

```bash
python3 -m http.server 8766 --directory ui &
npm install --no-save playwright@1 && npx playwright install chromium
node tests/ui/bridge-test.js
```

## Bridge

| Direction | Messages |
| --- | --- |
| UI to host | `ready`, `open`, `replay`, `home`, `next`, `prev`, `prompt`, `pinSubmit`, `setPin`, `admin` (`reload`, `exit`) |
| Host to UI | `library`, `opening`, `showStarted`, `slide`, `showEnded`, `home`, `openFailed`, `activity`, `pin`, `pinResult`, `adminStatus`, `setupResult` |

Shapes are in `src/KioskPresenter.Core/Bridge.cs` and `ui/mock-data.js`.
