// Kiosk configuration. The native host may overwrite this file per install.
window.KIOSK_CONFIG = {
  brand: {
    logo: null,                     // e.g. "brand/logo.svg"; null shows the logo slot
    title: 'Hall of Discovery',
    subtitle: 'Choose a presentation to begin',
    accent: null                    // hex string overrides --accent, e.g. "#4fd1b5"
  },
  motion: true,
  firstRun: false,                  // host sets true until setup completes
  timers: {                         // seconds
    libraryIdle: 120,               // library silently scrolls back to the top
    playerIdle: 180,                // player → Still there? (skipped while video plays)
    prompt: 15,                     // Still there? countdown
    endAutoReturn: 15,              // end of presentation → library (0 = stay)
    openingSlow: 10,                // Back button appears on Opening
    stripHide: 4,                   // control strip auto-hide
    cursorHide: 3,
    inputLock: 1.5,                 // after opening a deck
    adminHold: 5                    // logo press-and-hold for the PIN pad
  },
  dev: true                         // shows the dev state switcher (press ` to toggle). Set false in production.
};

// HOST: in WebView2 the native host injects window.KIOSK_HOST_CONFIG (built from
// C:\ProgramData\KioskPresenter\config.json) before any page script runs.
if (window.KIOSK_HOST_CONFIG) window.KIOSK_CONFIG = window.KIOSK_HOST_CONFIG;
