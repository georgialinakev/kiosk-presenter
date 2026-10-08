// Drives ui/index.html in headless Chromium: mock mode, then a simulated WebView2 host.
// Run: serve ui/ on port 8766 (python3 -m http.server 8766 --directory ui), then node tests/ui/bridge-test.js
// CHROME_PATH overrides the browser; UI_URL overrides the page URL.
const { chromium } = require('playwright');
const BASE = process.env.UI_URL || 'http://localhost:8766/index.html';
const assert = (c, m) => { if (!c) { console.log('FAIL:', m); process.exitCode = 1; } else console.log('ok  -', m); };

(async () => {
  const b = await chromium.launch(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {});

  // ---- mock mode (plain browser) ----
  let p = await b.newPage({ viewport: { width: 1920, height: 1080 } });
  const ext = [];
  await p.route(u => !u.href.startsWith('http://localhost'), r => { ext.push(r.request().url()); r.abort(); });
  const errs = [];
  p.on('pageerror', e => errs.push(e.message));
  p.on('console', m => { if (m.type() === 'error' && !/404/.test(m.text())) errs.push(m.text()); });
  await p.goto(BASE); await p.waitForTimeout(1200);
  assert(await p.locator('[data-deck]').count() === 24, 'mock mode shows 24 mock decks');
  await p.locator('[data-deck="a1b2"]').click(); await p.waitForTimeout(3500);
  assert(await p.locator('[data-strip]').count() === 1, 'mock mode: opening then mock player with HTML strip');
  assert(errs.length === 0, 'mock mode: no page errors ' + JSON.stringify(errs));
  assert(ext.length === 0, 'mock mode: no external requests ' + JSON.stringify(ext));
  await p.close();

  // ---- native mode (stub chrome.webview) ----
  p = await b.newPage({ viewport: { width: 1920, height: 1080 } });
  const errs2 = [];
  p.on('pageerror', e => errs2.push(e.message));
  p.on('console', m => { if (m.type() === 'error' && !/404/.test(m.text())) errs2.push(m.text()); });
  await p.addInitScript(() => {
    const ls = []; window.__sent = [];
    window.chrome = window.chrome || {};
    window.chrome.webview = { postMessage: m => window.__sent.push(m), addEventListener: (t, f) => ls.push(f) };
    window.__host = m => ls.forEach(f => f({ data: m }));
  });
  await p.route('**/config.js', async r => {
    const res = await r.fetch(); const body = (await res.text()).replace('playerIdle: 180', 'playerIdle: 2').replace('dev: true', 'dev: false');
    r.fulfill({ response: res, body });
  });
  await p.goto(BASE); await p.waitForTimeout(1200);
  const sent = () => p.evaluate(() => window.__sent.map(m => m.cmd + (m.id ? ':' + m.id : '') + (m.visible != null ? ':' + m.visible : '')));
  const host = m => p.evaluate(x => window.__host(x), m);

  assert((await sent()).includes('ready'), 'native: UI sends ready');
  assert(await p.locator('[data-deck]').count() === 0, 'native: no mock decks before library');
  assert(!(await p.getByText('No presentations yet').isVisible().catch(() => false)), 'native: no empty message while waiting');

  await host({ type: 'library', status: 'ok', decks: [
    { id: 'd1', title: 'Kiosk Demo Gallery Tour', description: null, slideCount: 6, hasVideo: false, thumbnail: null },
    { id: 'd2', title: 'Kiosk Demo Video Spotlight', description: null, slideCount: 4, hasVideo: true, thumbnail: null }] });
  await p.waitForTimeout(300);
  assert(await p.locator('[data-deck]').count() === 2, 'native: host library renders 2 decks');

  await p.locator('[data-deck="d1"]').click(); await p.waitForTimeout(200);
  assert((await sent()).includes('open:d1'), 'native: tap sends open');
  await host({ type: 'opening', id: 'd1' });
  await host({ type: 'showStarted', id: 'd1', slideCount: 6 }); await p.waitForTimeout(300);
  assert(await p.locator('[data-layer="player"]').count() === 1, 'native: showStarted enters player');
  assert(await p.locator('[data-strip]').count() === 0, 'native: no HTML strip or fake slide in player');
  await host({ type: 'slide', n: 3, total: 6, video: false }); await p.waitForTimeout(100);
  await p.keyboard.press('ArrowRight'); await p.waitForTimeout(100);
  assert((await sent()).includes('next'), 'native: arrow key in player sends next');
  await host({ type: 'showEnded' }); await p.waitForTimeout(300);
  assert(await p.locator('[data-layer="end"]').count() === 1, 'native: showEnded shows End screen');
  await host({ type: 'home' }); await p.waitForTimeout(300);
  assert(await p.locator('[data-layer="library"]').count() === 1, 'native: host home returns to library');
  assert((await sent()).filter(x => x === 'home').length === 0, 'native: host-initiated home does not echo a home command');

  // Still there? with short timers.
  await host({ type: 'opening', id: 'd2' });
  await host({ type: 'showStarted', id: 'd2', slideCount: 4, video: false });
  await p.waitForTimeout(2800);
  assert(await p.locator('[data-layer="idle"]').count() === 1, 'native: Still there? appears after idle');
  assert((await sent()).includes('prompt:true'), 'native: prompt visible sent');
  await host({ type: 'activity' }); await p.waitForTimeout(200);
  assert(await p.locator('[data-layer="idle"]').count() === 1, 'native: host activity does not dismiss an open prompt');
  await p.locator('[data-layer="idle"] [data-autofocus]').first().click().catch(async () => p.locator('[data-layer="idle"] button').first().click());
  await p.waitForTimeout(200);
  assert((await sent()).includes('prompt:false'), 'native: prompt cleared sent');
  await host({ type: 'slide', n: 2, total: 4, video: true }); await p.waitForTimeout(2800);
  assert(await p.locator('[data-layer="idle"]').count() === 0, 'native: no idle prompt on a video slide');
  await host({ type: 'home' }); await p.waitForTimeout(200);

  // PIN + admin.
  await host({ type: 'pin' }); await p.waitForTimeout(200);
  assert(await p.locator('[data-layer="pin"]').count() === 1, 'native: host pin opens PIN pad');
  for (const k of '2468') await p.keyboard.press(k);
  await p.keyboard.press('Enter'); await p.waitForTimeout(100);
  assert((await p.evaluate(() => window.__sent.filter(m => m.cmd === 'pinSubmit').map(m => m.digits))).includes('2468'), 'native: pinSubmit carries digits');
  await host({ type: 'adminStatus', status: { version: '0.1.0', folder: 'C:\\x', folderOk: true, decks: 2, ppt: 'PowerPoint 16', network: true, lastRefresh: '10:00:00', account: 'PC\\kiosk', autostart: true, warnings: [] } });
  await host({ type: 'pinResult', ok: true }); await p.waitForTimeout(300);
  assert(await p.locator('[data-layer="admin"]').count() === 1, 'native: correct PIN shows admin panel');
  assert(errs2.length === 0, 'native: no page errors ' + JSON.stringify(errs2));
  await p.close();

  // ---- host-injected config (what UiConfigScript.Build emits): first-run setup ----
  p = await b.newPage({ viewport: { width: 1920, height: 1080 } });
  const errs3 = [];
  p.on('pageerror', e => errs3.push(e.message));
  await p.addInitScript(() => {
    const ls = []; window.__sent = [];
    window.chrome = window.chrome || {};
    window.chrome.webview = { postMessage: m => window.__sent.push(m), addEventListener: (t, f) => ls.push(f) };
    window.__host = m => ls.forEach(f => f({ data: m }));
    window.KIOSK_HOST_CONFIG = {
      brand: { logo: null, title: 'Hall of Discovery', subtitle: 'Choose a presentation to begin', accent: '#4fd1b5' },
      motion: true, firstRun: true,
      timers: { libraryIdle: 120, playerIdle: 180, prompt: 15, endAutoReturn: 15, openingSlow: 10, stripHide: 4, cursorHide: 3, inputLock: 1.5, adminHold: 5 },
      dev: false,
    };
  });
  await p.goto(BASE); await p.waitForTimeout(1200);
  assert(await p.locator('[data-layer="setup"]').count() === 1, 'host config: firstRun shows setup');
  assert(await p.locator('[data-dev]').count() === 0, 'host config: dev switcher hidden');
  for (const k of '13579') await p.keyboard.press(k);
  await p.keyboard.press('Enter'); await p.waitForTimeout(100);
  for (const k of '13579') await p.keyboard.press(k);
  await p.keyboard.press('Enter'); await p.waitForTimeout(100);
  const setPin = await p.evaluate(() => window.__sent.filter(m => m.cmd === 'setPin').map(m => m.digits));
  assert(setPin.includes('13579'), 'host config: confirmed PIN sends setPin');
  await p.evaluate(() => window.__host({ type: 'setupResult', ok: true })); await p.waitForTimeout(200);
  assert(await p.getByText('Setup complete').count() === 1, 'host config: setupResult shows done');
  const accent = await p.evaluate(() => getComputedStyle(document.querySelector('[data-kiosk]')).getPropertyValue('--accent').trim());
  assert(accent.toLowerCase() === '#4fd1b5', 'host config: brand accent applied (' + accent + ')');
  assert(errs3.length === 0, 'host config: no page errors ' + JSON.stringify(errs3));
  await b.close();
})();
