// Mock data + mock host. In WebView2 the real host replaces this over chrome.webview messages.
(function () {
  const d = (id, title, description, slideCount, hasVideo) =>
    ({ id, title, description, slideCount, hasVideo, thumbnail: 'thumbs/' + id + '.png' });

  const DECKS = [
    d('a1b2', 'Kiosk Demo Gallery Tour', 'Six-slide click-through', 6, false),
    d('c3d4', 'Kiosk Demo Video Spotlight', 'Four slides with a short film', 4, true),
    d('n1le', 'Ancient Rivers: Life Along the Nile', 'Farming, trade and belief in the valley', 14, false),
    d('st3m', 'The Age of Steam', 'Engines that changed how we work', 9, false),
    d('mask', 'Masks of West Africa', 'Carving, ceremony and performance', 12, true),
    d('sky1', 'Starlight: How We Map the Sky', null, 10, false),
    d('tide', 'Tides, Trade and Tall Ships: Four Centuries of the City\u2019s Harbour Quay', 'Merchants, sailors and the river', 18, false),
    d('dino', 'Dinosaurs of the Badlands', 'Fossils from the summer dig', 11, true),
    d('silk', 'Textiles of the Silk Road', 'Weaving traditions from east to west', 8, false),
    d('clok', 'Clockwork and Automata', 'Mechanical marvels from the collection', 7, false),
    d('prnt', 'The Printing Revolution', null, 9, false),
    d('glas', 'Glass: From Sand to Window', 'How a furnace works', 6, false),
    d('coal', 'Voices of the Coalfield', 'Miners and families in their own words', 5, true),
    d('bird', 'Birds of the Wetland Reserve', 'A year of migration', 12, false),
    d('brnz', 'The Bronze Age Hoard', 'Found by a farmer in 1987', 8, false),
    d('mini', 'Portraits in Miniature', null, 10, false),
    d('maps', 'Maps That Changed the World', 'From clay tablets to satellites', 13, false),
    d('lite', 'The Lighthouse Keepers', 'Life on the rock', 6, true),
    d('foss', 'Shells, Stones and Fossils', 'Start your own collection', 9, false),
    d('obsv', 'Women of the Observatory', 'The computers who catalogued the stars', 11, false),
    d('rail', 'The Railway Comes to Town', null, 7, false),
    d('ming', 'Ceramics of the Ming Dynasty', 'Blue, white and the potter\u2019s wheel', 10, false),
    d('cabn', 'The Natural History Cabinet', 'A Victorian collector\u2019s room', 8, false),
    d('manu', 'Medieval Manuscripts Up Close', 'Ink, gold and parchment', 12, false)
  ];
  const byId = Object.fromEntries(DECKS.map(x => [x.id, x]));
  const pick = ids => ids.map(id => ({ ...byId[id] }));

  const SETS = {
    one: pick(['a1b2']),
    six: pick(['a1b2', 'c3d4', 'n1le', 'st3m', 'mask', 'sky1']),
    all: pick(DECKS.map(x => x.id)),
    long: pick(['a1b2', 'c3d4', 'tide', 'n1le', 'st3m', 'mask']),
    nothumb: pick(['a1b2', 'c3d4', 'n1le', 'st3m', 'mask', 'sky1']).map(x => x.id === 'st3m' ? { ...x, thumbnail: null } : x),
    video: pick(['c3d4', 'mask', 'dino', 'coal', 'lite', 'a1b2'])
  };
  const lib = (set, status = 'ok') => ({ type: 'library', status, decks: status === 'ok' ? SETS[set] : [] });

  const STATUS_OK = {
    version: '1.0.0 (build 214)', folder: 'C:\\KioskContent', folderOk: true, decks: 24,
    ppt: 'Microsoft PowerPoint 16.0 (Microsoft 365)', network: true,
    lastRefresh: '10:42:18', account: 'MUSEUM\\kiosk-lobby', autostart: true, warnings: []
  };
  const STATUS_WARN = {
    ...STATUS_OK, ppt: null, folderOk: false, network: false, decks: 0, lastRefresh: '09:15:02',
    warnings: [
      { id: 'ppt', title: 'PowerPoint not found', detail: 'COM class PowerPoint.Application is not registered. Install Microsoft PowerPoint 2016 or later, then reload.' },
      { id: 'folder', title: 'Content folder missing', detail: 'C:\\KioskContent does not exist or is not readable by MUSEUM\\kiosk-lobby (ERROR_PATH_NOT_FOUND).' },
      { id: 'net', title: 'Offline', detail: 'No network connection since 09:15. Local presentations still play; remote status reporting is paused.' }
    ]
  };

  // Every reviewable state. `s` is merged into the app's base state.
  const P = (id, group, label, s) => ({ id, group, label, s });
  const PRESETS = [
    P('lib-1', 'Library', '1 deck', { screen: 'library', lib: lib('one') }),
    P('lib-6', 'Library', '6 decks (one screen)', { screen: 'library', lib: lib('six') }),
    P('lib-24', 'Library', '24 decks (scrolls)', { screen: 'library', lib: lib('all') }),
    P('lib-long', 'Library', '70-character title', { screen: 'library', lib: lib('long') }),
    P('lib-nothumb', 'Library', 'Missing thumbnail', { screen: 'library', lib: lib('nothumb') }),
    P('lib-video', 'Library', 'Video badge', { screen: 'library', lib: lib('video') }),
    P('lib-empty', 'Library', 'Empty folder', { screen: 'library', lib: lib(null, 'empty') }),
    P('lib-unavailable', 'Library', 'Content unavailable', { screen: 'library', lib: lib(null, 'unavailable') }),
    P('open', 'Opening', 'Loading', { screen: 'opening', lib: lib('six'), deckId: 'c3d4', slow: false }),
    P('open-slow', 'Opening', 'Slow load (10 s)', { screen: 'opening', lib: lib('six'), deckId: 'c3d4', slow: true }),
    P('play-visible', 'Player strip', 'Visible', { screen: 'player', lib: lib('six'), deckId: 'a1b2', slide: 1, stripVisible: true }),
    P('play-hidden', 'Player strip', 'Auto-hidden (4 s)', { screen: 'player', lib: lib('six'), deckId: 'a1b2', slide: 1, stripVisible: false }),
    P('play-pressed', 'Player strip', 'Pressed (Next)', { screen: 'player', lib: lib('six'), deckId: 'a1b2', slide: 2, stripVisible: true, stripPressed: 'next', slideKind: 'dark' }),
    P('play-light', 'Player strip', 'On a light slide', { screen: 'player', lib: lib('six'), deckId: 'a1b2', slide: 2, stripVisible: true, slideKind: 'light' }),
    P('play-dark', 'Player strip', 'On a dark slide', { screen: 'player', lib: lib('six'), deckId: 'a1b2', slide: 3, stripVisible: true, slideKind: 'dark' }),
    P('play-video', 'Player strip', 'On a video slide', { screen: 'player', lib: lib('six'), deckId: 'c3d4', slide: 2, stripVisible: true, slideKind: 'video' }),
    P('end', 'End', 'Default', { screen: 'end', lib: lib('six'), deckId: 'a1b2' }),
    P('idle-player', 'Still there?', 'Over the player', { screen: 'player', lib: lib('six'), deckId: 'a1b2', slide: 3, slideKind: 'dark', stripVisible: false, idle: 15, idleOver: 'player' }),
    P('pin-empty', 'Admin PIN', 'Empty', { screen: 'library', lib: lib('six'), pin: { digits: '', error: null } }),
    P('pin-digits', 'Admin PIN', 'Digits entered', { screen: 'library', lib: lib('six'), pin: { digits: '246', error: null } }),
    P('pin-wrong', 'Admin PIN', 'Wrong PIN', { screen: 'library', lib: lib('six'), pin: { digits: '', error: 'wrong' } }),
    P('admin-ok', 'Admin panel', 'Healthy', { screen: 'admin', lib: lib('all'), admin: { status: STATUS_OK, confirmExit: false } }),
    P('admin-warn', 'Admin panel', 'Warnings', { screen: 'admin', lib: lib('all'), admin: { status: STATUS_WARN, confirmExit: false } }),
    P('admin-exit', 'Admin panel', 'Exit kiosk (confirm)', { screen: 'admin', lib: lib('all'), admin: { status: STATUS_OK, confirmExit: true } }),
    P('setup-enter', 'First-run setup', 'Enter PIN', { screen: 'setup', setup: { step: 'enter', digits: '', first: '', error: null } }),
    P('setup-confirm', 'First-run setup', 'Confirm PIN', { screen: 'setup', setup: { step: 'confirm', digits: '24', first: '2468', error: null } }),
    P('setup-mismatch', 'First-run setup', 'PINs don\u2019t match', { screen: 'setup', setup: { step: 'enter', digits: '', first: '', error: 'mismatch' } }),
    P('setup-done', 'First-run setup', 'Done', { screen: 'setup', setup: { step: 'done', digits: '', first: '', error: null } })
  ];

  // ---- bridge ----
  const listeners = [];
  const emit = (msg, delay = 0) => setTimeout(() => listeners.forEach(fn => fn(msg)), delay);
  const native = window.chrome && window.chrome.webview;
  if (native) native.addEventListener('message', e => listeners.forEach(fn => fn(e.data)));

  const mock = { pin: '2468', current: lib('all'), status: STATUS_OK };
  const clock = () => new Date().toTimeString().slice(0, 8);

  function mockHost(m) {
    switch (m.cmd) {
      case 'open':
      case 'replay': {
        const deck = mock.current.decks.find(x => x.id === m.id);
        emit({ type: 'opening', id: m.id });
        emit({ type: 'showStarted', id: m.id, slideCount: deck ? deck.slideCount : 1 }, 2800);
        break;
      }
      case 'home': emit(mock.current); break;
      case 'pinSubmit':
        if (m.digits === mock.pin) { emit({ type: 'pinResult', ok: true }, 150); emit({ type: 'adminStatus', status: mock.status }, 160); }
        else emit({ type: 'pinResult', ok: false }, 150);
        break;
      case 'setPin': mock.pin = m.digits; emit({ type: 'setupResult', ok: true }, 100); break;
      case 'admin':
        if (m.action === 'reload') { mock.status = { ...mock.status, lastRefresh: clock() }; emit(mock.current, 300); emit({ type: 'adminStatus', status: mock.status }, 300); }
        break;
    }
  }

  window.KioskBridge = {
    send(msg) { console.log('[bridge →]', JSON.stringify(msg)); native ? native.postMessage(msg) : mockHost(msg); },
    on(fn) { listeners.push(fn); return () => listeners.splice(listeners.indexOf(fn), 1); },
    isNative: !!native
  };
  window.KIOSK_MOCK = { DECKS, SETS, PRESETS, STATUS_OK, STATUS_WARN, lib, initial: lib('all'), mockPin: mock.pin };
})();
