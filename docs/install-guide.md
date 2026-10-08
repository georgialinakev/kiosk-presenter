# Install guide

For the tech installing Kiosk Presenter on a client PC.

## Before you go

1. Sign in as the account the kiosk will run as and open one of the client's decks in PowerPoint.
   It must open and play with no sign-in, activation or "product notice" prompts. If it doesn't,
   fix Office licensing for that account first; the kiosk can't hide those prompts.
2. Get the decks from the client (email is fine). Prefer local MP4 video over YouTube.
3. Download `KioskPresenter-Setup-X.Y.Z.exe` from the GitHub release and check its SHA-256.

## Install

1. Copy the Setup exe to the PC and run it. Windows may show "Windows protected your PC": choose
   More info, then Run anyway (the exe isn't code-signed yet). Approve the admin prompt.
2. Enter the title for the top of the screen, pick the user the kiosk starts for, and leave the
   demo presentations ticked for the walkthrough. Click Install.
3. Optional, for a PC that should boot straight into the kiosk: run Microsoft's Sysinternals
   Autologon and enable auto sign-in for that user.
4. Copy the client's decks into `C:\ProgramData\KioskPresenter\Content`.
5. Sign out and sign in as the kiosk user (or reboot). The kiosk starts by itself.
6. First run asks for a 4 to 8 digit admin PIN. Store it with the client's other credentials.

## Check it

- Open each deck, page through it, let one reach the end, press Home mid-show.
- Press the Windows key, Alt+Tab, Alt+F4: nothing should happen.
- Leave a deck idle for 3 minutes: "Are you still there?" appears, then it returns to the library.
- Admin: press and hold the logo for 5 seconds (or Ctrl+Shift+Q), enter the PIN.
  Exit kiosk gives the normal desktop back until the next sign-in.

## Day to day

- **New decks:** copy them into the content folder (remote session is fine). They appear within 30 seconds.
- **Change the title, logo or timers:** edit `C:\ProgramData\KioskPresenter\config.json`, then
  admin panel, Reload library.
- **Logs:** `C:\ProgramData\KioskPresenter\Logs`.
- **Maintenance sign-in:** hold Shift while Windows signs in to skip auto sign-in.

## Upgrade or remove

- **Upgrade:** run the newer Setup exe. PIN, settings and decks are kept.
- **Uninstall:** Settings, Apps, Kiosk Presenter, Uninstall. It removes the app, the auto-start
  task and the shortcut. Tick "Also delete presentations, settings and logs" to remove
  `C:\ProgramData\KioskPresenter` too. Turn off auto sign-in in Autologon if you set it.
- **Silent:** `KioskPresenter-Setup-X.Y.Z.exe /quiet /user:PC\kiosk /demo /title:"Hall of Discovery"`;
  remove with `/uninstall /quiet` (add `/purge` to delete data).
