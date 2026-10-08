// Every visible string. Edit or translate here; layout never hard-codes copy.
// {n}, {total}, {title}, {drive}, {p} are filled in at runtime.
window.KIOSK_STRINGS = {
  library: {
    slides: '{n} slides',
    slidesOne: '1 slide',
    video: 'Video',
    open: 'Open {title}',
    emptyTitle: 'No presentations yet',
    emptyBody: 'New presentations will appear here soon.',
    unavailableTitle: 'This isn\u2019t available right now',
    unavailableBody: 'Please ask a member of staff for help.'
  },
  opening: {
    opening: 'Opening',
    slowTitle: 'This is taking longer than usual',
    slowBody: 'You can keep waiting, or go back and choose another presentation.',
    back: 'Back to all presentations'
  },
  player: {
    home: 'All presentations',
    prev: 'Previous slide',
    next: 'Next slide',
    counter: '{n} / {total}'
  },
  end: {
    kicker: 'End of presentation',
    replay: 'Watch again',
    home: 'All presentations'
  },
  idle: {
    title: 'Are you still there?',
    body: 'This presentation will close in {n} seconds.',
    stay: 'I\u2019m still here',
    leave: 'Close presentation'
  },
  pin: {
    title: 'Staff access',
    prompt: 'Enter the admin PIN',
    wrong: 'Incorrect PIN. Please try again.',
    del: 'Delete',
    submit: 'Enter',
    cancel: 'Cancel'
  },
  admin: {
    title: 'Kiosk admin',
    close: 'Close',
    status: 'Status',
    actions: 'Actions',
    healthy: 'Everything is working',
    problems: '{n} problems need attention',
    problemsOne: '1 problem needs attention',
    fields: {
      version: 'App version',
      folder: 'Content folder',
      decks: 'Presentations found',
      ppt: 'PowerPoint',
      network: 'Network',
      refresh: 'Last refresh',
      account: 'Runs as',
      autostart: 'Starts automatically'
    },
    online: 'Online',
    offline: 'Offline',
    yes: 'Yes',
    no: 'No',
    notFound: 'Not found',
    missing: 'Missing',
    reload: 'Reload library',
    exit: 'Exit kiosk',
    exitTitle: 'Exit kiosk mode?',
    exitBody: 'Visitors will see the Windows desktop until the kiosk is started again.',
    exitConfirm: 'Exit kiosk',
    cancel: 'Cancel'
  },
  setup: {
    step: 'Step {n} of 2',
    enterTitle: 'Create an admin PIN',
    enterBody: 'Staff use this PIN to leave the kiosk or change settings. Use 4 to 8 digits.',
    confirmTitle: 'Enter the PIN again',
    confirmBody: 'Type the same digits to confirm.',
    mismatch: 'Those PINs didn\u2019t match. Please create it again.',
    doneTitle: 'Setup complete',
    doneBody: 'To return to the admin panel, press and hold the logo for 5 seconds, or press Ctrl+Shift+Q.',
    start: 'Start kiosk'
  }
};
