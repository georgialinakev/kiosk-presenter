# Museum Kiosk — Design Handoff Brief

Oct 8, 2026 · @Kevin Hatfield

## Goal

Design the visitor-facing UI for a museum kiosk app: a library of presentations a visitor picks from, a fullscreen player, and a return to the library when they finish or walk away. While the app runs it takes over the PC, so the UI is the only thing a visitor sees: no Windows desktop, no PowerPoint editor. It stays that way until a staff member exits with the admin PIN.

This replaces a lost prototype, a basic Python/Tk card grid. The new build hosts the library screens as HTML/CSS in a WebView2 inside a native Windows shell. Your output is the real frontend, not a reference picture.

Content management is out of scope for the UI. Staff add or remove decks by dropping files into a folder, and the library picks them up within 30 seconds.

## Users, hardware, input

Visitors are walk-up members of the public of any age, standing or seated, with no instructions and a few seconds of patience. Staff use the kiosk rarely, through a hidden admin entry.

- **Primary display:** 1920×1080 landscape touchscreen at 100% scaling. Layouts must also hold at 1366×768 and at 3840×2160 with 200% scaling.
- **Stretch variant:** 1080×1920 portrait. Show the library in portrait once, so we know the grid reflows.
- **Viewing distance:** about an arm's length. Body text is read at about 0.6 m, headings from about 2 m.

Every action must work with all three input methods. No hover-only affordances, no gesture-only actions.

| Input | Must support | Notes |
| --- | --- | --- |
| Touch | Tap, vertical drag-scroll with momentum, horizontal swipe in the player | Primary. Visible pressed state on every target, not just hover |
| Mouse | Click, wheel scroll, hover states | Cursor visible only while the mouse moves; hide after 3 s idle |
| Keyboard | Arrow keys move focus, Enter opens, Esc goes back | Visible focus ring on every focusable element, Tab order matches visual order |

## Screens and states

Nine surfaces, five of them visitor-facing. Every state listed must be drawn; the backend is built against these frames.

| # | Screen | Purpose | States to draw |
| --- | --- | --- | --- |
| 1 | Library (home) | Pick a presentation | 1 deck, 6 decks (fits one screen), 24 decks (scrolls), a 70-character title, missing thumbnail, "Video" badge, empty folder, content unavailable |
| 2 | Opening | Covers the 2 to 8 s PowerPoint takes to load | Loading with deck title and thumbnail, slow load after 10 s with a Back button |
| 3 | Player control strip | Home, Previous, Next and a slide counter over the running show | Visible, auto-hidden after 4 s, pressed, on a light slide, on a dark slide, on a video slide |
| 4 | End of presentation | Offer replay or return | Default, with a 15 s auto-return countdown |
| 5 | Still there? | Idle prompt before resetting | Default with a 15 s countdown, over the player and over the library |
| 6 | Attract | Idle screensaver that invites a touch | Default loop, the moment of first touch |
| 7 | Admin PIN pad | Hidden staff entry | Empty, digits entered (masked), wrong PIN, locked out for 60 s after 5 tries |
| 8 | Admin panel | Status and a few staff actions | Healthy, with warnings (PowerPoint not found, content folder missing, offline), importing from a USB drive (file list, progress) |
| 9 | First-run setup | Shown once after install: set the admin PIN, confirm the content folder | Enter PIN, confirm PIN, PINs don't match, PowerPoint not found (warning, can continue), done |

&#91;embedded content: screen flow · 8 screens, one hub\]

The Library is the hub: every visitor path, including errors and idle timeouts, lands back there. The staff path branches off it and is never visible.

**Screen 1, library card contents:** a 16:9 thumbnail of slide 1, the title, an optional one-line description, the slide count, and a "Video" badge when the deck contains video. Header carries a logo slot, a title and a subtitle, all from config.

**Screen 3 is not HTML.** It's a native WPF window floating above PowerPoint's own fullscreen show. Keep it to simple shapes, icons and one text line. It must stay readable over any slide colour, so a dark translucent pill with light icons is the expected answer. Put it bottom centre, within seated reach.

**Screens 7, 8 and 9 are staff-only.** Utilitarian is fine; they inherit the tokens and need no polish beyond that. The admin panel shows app version, content folder, decks found, PowerPoint version, network status and last refresh time. It also shows which account the kiosk runs as and whether it starts automatically. Its actions are Reload library, Open content folder, Import from USB drive, Exit kiosk (with a confirm step), Sign out, Restart PC and Close.

**Copy rules for visitors:** no technical words. "This presentation isn't available right now" rather than "COM error" or "folder not found". Error detail appears only in the admin panel.

## Interaction rules

The UI has to survive impatient, imprecise, repeated taps. These rules are requirements, not suggestions.

**Touch**

- Minimum target 64×64 px at 1080p, at least 16 px between targets.
- Pressed feedback within 100 ms. Activation fires on release, and a drag of more than 10 px cancels it, so scrolling never opens a deck.
- After a deck is tapped, the library ignores input for 1.5 s. One tap opens one show.
- The slide area belongs to PowerPoint: tap and swipe there advance slides natively. Don't design invisible tap zones over the slide; they would block video controls.

**Mouse**

- Hover and pressed states on every target. Cursor hides after 3 s without movement.
- Wheel scrolls the library. A thin scrollbar may appear on mouse use only.

**Keyboard**

| Where | Keys | Action |
| --- | --- | --- |
| Library | Arrow keys | Move focus between cards |
| Library | Enter or Space | Open the focused deck |
| Player | Right, Space, Page Down | Next slide |
| Player | Left, Page Up | Previous slide |
| Player | Esc | Back to the library (PowerPoint owns other keys during a show) |
| Any modal | Enter / Esc | Primary action / dismiss |
| Anywhere | Ctrl+Shift+Q | Admin PIN pad |

Focus ring: 3 px in the accent colour with a 2 px offset, shown on keyboard use only (focus-visible), never removed.

**Hidden admin entry for touch:** press and hold the header logo for 5 s. No visual hint before it triggers.

**Idle behaviour (frames only, timers live in config):** library goes to Attract after 2 min. The player shows Still there? after 3 min without input, unless a video is playing. End of presentation returns home after 15 s.

**Motion:** fades and scales only, 250 ms or less, with a config switch to turn motion off. The UI makes no sound.

## Visual direction and tokens

Calm and gallery-like: the presentations are the colour, the chrome stays neutral. Default to a dark theme, since galleries are dim and a bright white screen glares. The museum's branding isn't known yet, so everything brand-specific must come from tokens and config.

The prototype's palette is a starting point you may refine, as long as contrast holds:

| Token | Prototype value | Role |
| --- | --- | --- |
| `--bg` | #1a1a2e | Page background |
| `--surface` | #16213e | Cards, panels |
| `--surface-raised` | (new) | Modals, control strip |
| `--text` | #ffffff | Primary text |
| `--text-muted` | #aaaaaa | Subtitles, metadata |
| `--accent` | #e94560 | Primary actions, badges, focus ring |
| `--accent-contrast` | (new) | Text on accent |
| `--danger` | (new) | Wrong PIN, admin warnings |
| `--scrim` | (new) | Backdrop behind modals |

**Brand slots in config:** logo image, accent colour, header title and subtitle, and a light theme variant. Show one alternate accent so we can see the theming works.

**Type:** bundle one open-licensed (OFL) web font so Windows 10 and 11 render identically; your pick. Minimum sizes at 1080p: display 48 px, screen title 32 px, card title 24 px, body 20 px, metadata 16 px. Contrast meets WCAG AA: 4.5:1 for body text, 3:1 for large text and icons.

**Icons:** one consistent outline set, shipped as inline SVG. The kiosk may be offline, so no CDN fonts, icons or images.

**Avoid:** blur and glass effects (the native control strip can't match them), emoji as icons (the prototype used 🎬), heavy gradients and drop shadows.

## Constraints and out of scope

The HTML screens run inside WebView2 (Chromium) from local files with no network, and a native host drives them over a message bridge.

**Technical constraints**

- Static output only: HTML, CSS, JS, fonts, SVG. A build step is fine if it emits static files; no runtime fetches to the internet.
- Vanilla JS or a small framework that compiles to static files. Modern CSS (grid, container queries, `:focus-visible`) is safe.
- All copy strings live in one file, so they can be edited or translated later without touching layout.
- Lock down browser behaviour in CSS/JS: no text selection, no context menu, no image dragging, no pinch-zoom, no anchors that navigate.

**Data contract.** Build against a mock data module with this shape. The host replaces the mock with real messages.

```json
{
  "type": "library",
  "status": "ok",
  "decks": [
    { "id": "a1b2", "title": "Kiosk Demo Gallery Tour", "description": "Six-slide click-through", "slideCount": 6, "hasVideo": false, "thumbnail": "thumbs/a1b2.png" }
  ]
}
```

`status` is one of `ok`, `empty` or `unavailable`. The UI sends commands back as `{ "cmd": "open", "id": "a1b2" }`, plus `home`, `replay`, `pinSubmit` (digits only; the host checks the hash) and `admin` with an action name, plus setPin on first run.

**Include a dev-only state switcher** that flips the mock through every state in the Screens and states table. It's how we'll review each state and how the test plan exercises them.

**Out of scope:** content management screens, analytics, multiple languages (strings file only), audio, the decks themselves, and Windows configuration (Setup handles that).

## Deliverables

Design is done when the HTML screens run against the mock with every state reachable. Backend work starts after that.

- [ ] Artboards for every screen and state in the table, at 1920×1080.
- [ ] Library at 1366×768 and 1080×1920 portrait.
- [ ] Library with one alternate accent colour, to prove theming.
- [ ] Control strip spec for native rebuild: sizes, spacing, colours, icon set, and all its states.
- [ ] Tokens as CSS custom properties in one file.
- [ ] Working HTML/CSS/JS for screens 1, 2 and 4 to 9, wired to the mock data module, with the dev state switcher.
- [ ] Keyboard-only walkthrough passes: every action reachable without touch or mouse.
- [ ] Contrast checked against WCAG AA for body text, large text and icons.

**Mock content:** use the two demo decks as real entries, "Kiosk Demo Gallery Tour" (6 slides, no video) and "Kiosk Demo Video Spotlight" (4 slides, Video badge). Pad the 24-deck state with plausible museum exhibit titles.
