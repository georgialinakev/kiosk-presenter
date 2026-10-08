"""Apply the host-integration edits to a fresh Claude Design export.

Usage: python tools/patch-ui.py ui
  (run after copying a new export over ui/; the export's "Kiosk App.dc.html" must be renamed to index.html)

Every edit is a literal find/replace that must match exactly once, so a new export
either patches cleanly or fails loudly and shows which edit no longer fits.
Each edit is marked "HOST:" in the output. Already-patched files are skipped.
"""
import sys
from pathlib import Path

ui = Path(sys.argv[1] if len(sys.argv) > 1 else "ui")
path = ui / "index.html"
src = path.read_text(encoding="utf-8")
if "HOST:" in src:
    sys.exit(f"{path} is already patched; copy a fresh export over it first.")

EDITS = [
    # 1. Vendored React so nothing loads from unpkg.com.
    (
        '<script src="./support.js"></script>',
        '<!-- HOST: React is vendored; support.js skips its CDN load when window.React exists. -->\n'
        '<script src="vendor/react.production.min.js"></script>\n'
        '<script src="vendor/react-dom.production.min.js"></script>\n'
        '<script src="./support.js"></script>',
    ),
    # 2. The runtime calls componentDidUpdate(prevProps) only; track slow ourselves.
    (
        "    if (this.live && (top !== this.lastTop || (!prevState.slow && this.state.slow))) {\n"
        "      this.lastTop = top;",
        "    // HOST: the dc runtime passes no prevState, so remember the last slow flag here.\n"
        "    const slowNow = !!this.state.slow, slowEdge = slowNow && !this.lastSlow;\n"
        "    this.lastSlow = slowNow;\n"
        "    if (this.live && (top !== this.lastTop || slowEdge)) {\n"
        "      this.lastTop = top;",
    ),
    # 3. Native boot: no mock decks, wait for the host's library, then say ready.
    (
        "      ...this.base(), ready: true, lib: M.initial, mouse: false, devOpen: false, preset: null,",
        "      // HOST: in WebView2 start empty and wait for the host's first library message.\n"
        "      ...this.base(), ready: true, lib: this.isNative() ? { status: 'ok', decks: [] } : M.initial,\n"
        "      waiting: this.isNative(), video: false, mouse: false, devOpen: false, preset: null,",
    ),
    (
        "    this.offBridge = window.KioskBridge.on(m => this.onHost(m));",
        "    this.offBridge = window.KioskBridge.on(m => this.onHost(m));\n"
        "    // HOST: ask for the library, config-dependent state and anything pending.\n"
        "    if (this.isNative()) window.KioskBridge.send({ cmd: 'ready' });",
    ),
    # 4. New host messages.
    (
        "      case 'library': this.setState({ lib: { status: m.status, decks: m.decks || [] } }); break;",
        "      case 'library': this.setState({ lib: { status: m.status, decks: m.decks || [] }, waiting: false }); break;\n"
        "      // HOST: messages the real host adds on top of the design's mock protocol.\n"
        "      case 'slide': this.lastInput = now; this.setState({ slide: m.n || 1, total: m.total || this.state.total, video: !!m.video }); break;\n"
        "      case 'activity': if (this.state.idle == null) { this.lastInput = now; this.libReset = false; } break;\n"
        "      case 'home': this.goLibrary(); break;\n"
        "      case 'openFailed': this.goLibrary(); break;\n"
        "      case 'pin': this.openPin(); break;",
    ),
    (
        "        this.setState({ screen: 'player', slide: 1, total: m.slideCount || 1, stripVisible: true, slideKind: null, stripPressed: null });",
        "        this.setState({ screen: 'player', slide: 1, total: m.slideCount || 1, stripVisible: true, slideKind: null, stripPressed: null, video: !!m.video });",
    ),
    # 5. Going home because the host said so must not echo a home command back.
    (
        "  goHome = () => {\n"
        "    this.send({ cmd: 'home' });\n"
        "    this.lastInput = Date.now(); this.endAt = null; this.libReset = false;",
        "  goHome = () => {\n"
        "    this.send({ cmd: 'home' });\n"
        "    this.goLibrary();\n"
        "  };\n"
        "  // HOST: local half of goHome, used when the host already knows (Esc in PowerPoint, open failed).\n"
        "  goLibrary = () => {\n"
        "    if (this.state.idle != null) this.send({ cmd: 'prompt', visible: false });\n"
        "    this.lastInput = Date.now(); this.endAt = null; this.libReset = false;",
    ),
    # 6. Next/previous drive PowerPoint through the host; position comes back as 'slide'.
    (
        "  next() { const s = this.state; if (s.slide < s.total) this.setState({ slide: s.slide + 1, slideKind: null }); else this.goEnd(); }\n"
        "  prev() { const s = this.state; if (s.slide > 1) this.setState({ slide: s.slide - 1, slideKind: null }); }\n"
        "  stay() { this.lastInput = Date.now(); this.stripAt = Date.now(); this.setState({ idle: null, idleOver: null }); }",
        "  next() { if (this.isNative()) return this.send({ cmd: 'next' }); const s = this.state; if (s.slide < s.total) this.setState({ slide: s.slide + 1, slideKind: null }); else this.goEnd(); }\n"
        "  prev() { if (this.isNative()) return this.send({ cmd: 'prev' }); const s = this.state; if (s.slide > 1) this.setState({ slide: s.slide - 1, slideKind: null }); }\n"
        "  stay() {\n"
        "    // HOST: tell the host to drop its window back behind the slideshow.\n"
        "    if (this.state.idle != null) this.send({ cmd: 'prompt', visible: false });\n"
        "    this.lastInput = Date.now(); this.stripAt = Date.now(); this.setState({ idle: null, idleOver: null });\n"
        "  }\n"
        "  isNative() { return !!(window.KioskBridge && window.KioskBridge.isNative) && !this.props.embedded; }",
    ),
    # 7. Still there? appearing must lift the host window over the slideshow.
    (
        "      if (s.idle !== left) { patch.idle = left; patch.idleOver = 'player'; }",
        "      if (s.idle !== left) { if (s.idle == null) this.send({ cmd: 'prompt', visible: true }); patch.idle = left; patch.idleOver = 'player'; }",
    ),
    # 8. Idle exemption uses the host's per-slide video flag.
    (
        "  videoPlaying() { return this.slideKindNow() === 'video'; }",
        "  videoPlaying() { return this.isNative() ? !!this.state.video : this.slideKindNow() === 'video'; }",
    ),
    # 9. Library message hidden while waiting for the first host library.
    (
        "      libOk, libMessage: !libOk,",
        "      libOk, libMessage: !libOk && !s.waiting,",
    ),
    # 10. Native mode draws no fake slide or HTML strip: PowerPoint and the WPF strip do that.
    (
        "      isPlayer: s.screen === 'player', devTag: this.live && C.dev,",
        "      isPlayer: s.screen === 'player', devTag: this.live && C.dev, mockShow: !this.isNative(),",
    ),
    (
        '  <img src="{{ slideSrc }}" alt="" style="position: absolute; inset: 0; width: 100%; height: 100%; object-fit: contain;">',
        '  <!-- HOST: in WebView2 the real slideshow and the native WPF strip sit above this layer. -->\n'
        '  <sc-if value="{{ mockShow }}"><div style="position: absolute; inset: 0;">\n'
        '  <img src="{{ slideSrc }}" alt="" style="position: absolute; inset: 0; width: 100%; height: 100%; object-fit: contain;">',
    ),
]

# The strip's wrapper closes right before the player layer's own closing tag.
STRIP_CLOSE_OLD = (
    "      </button>\n"
    "    </div>\n"
    "  </div>\n"
    "</div>\n"
    "</sc-if>\n"
    "\n"
    '<sc-if value="{{ isEnd }}"'
)
STRIP_CLOSE_NEW = (
    "      </button>\n"
    "    </div>\n"
    "  </div>\n"
    "  </div></sc-if>\n"
    "</div>\n"
    "</sc-if>\n"
    "\n"
    '<sc-if value="{{ isEnd }}"'
)
EDITS.append((STRIP_CLOSE_OLD, STRIP_CLOSE_NEW))

for old, new in EDITS:
    n = src.count(old)
    if n != 1:
        sys.exit(f"edit expected 1 match, found {n}:\n{old[:120]}")
    src = src.replace(old, new)

path.write_text(src, encoding="utf-8")
print(f"applied {len(EDITS)} edits to {path}")

# mock-data.js: keep the first bridge if the runtime evaluates the script twice.
mock = ui / "mock-data.js"
text = mock.read_text(encoding="utf-8")
head = "(function () {\n"
if "HOST:" not in text:
    if text.count(head) < 1:
        sys.exit("mock-data.js: IIFE start not found")
    text = text.replace(head, head + "  // HOST: the dc runtime can run helmet scripts twice; a second copy would orphan the first bridge's listeners.\n  if (window.KioskBridge) return;\n", 1)
    mock.write_text(text, encoding="utf-8")
    print(f"patched {mock}")

# config.js: let the host's injected config win.
config = ui / "config.js"
text = config.read_text(encoding="utf-8")
if "KIOSK_HOST_CONFIG" not in text:
    text = text.rstrip("\n") + "\n\n// HOST: in WebView2 the native host injects window.KIOSK_HOST_CONFIG (built from\n// C:\\ProgramData\\KioskPresenter\\config.json) before any page script runs.\nif (window.KIOSK_HOST_CONFIG) window.KIOSK_CONFIG = window.KIOSK_HOST_CONFIG;\n"
    config.write_text(text, encoding="utf-8")
    print(f"patched {config}")

# React must be vendored next to index.html.
for f in ("vendor/react.production.min.js", "vendor/react-dom.production.min.js"):
    if not (ui / f).exists():
        print(f"WARNING: {ui / f} missing; copy React 18.3.1 UMD builds there (npm pack react@18.3.1 react-dom@18.3.1).")
