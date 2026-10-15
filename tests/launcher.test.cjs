const { test } = require('node:test');
const assert = require('node:assert/strict');
const { JSDOM, VirtualConsole } = require('jsdom');
const fs = require('node:fs');
const path = require('node:path');
const { context } = require('./helpers.cjs');

const html = fs.readFileSync(path.join(__dirname, '..', 'site', 'index.html'), 'utf8');
const PAGE_URL = 'https://dangeoghegan.github.io/lesson-grader-gas/';
const EXEC_URL = 'https://script.google.com/macros/s/AKfycbTESTDEPLOYMENTID0000000000000000000/exec';
const ALT_URL = 'https://script.google.com/macros/s/AKfycbSECONDDEPLOYMENTID111111111111111111/exec';
const DEV_URL = 'https://script.google.com/macros/s/AKfycbTESTDEPLOYMENTID0000000000000000000/dev';

/* ------------------------------------------------------------------ *
 * Apps Script side: the read-only launcher probe in doGet
 * ------------------------------------------------------------------ */
function gas({ email = 'dan@school.edu.au', properties = {} } = {}) {
  const outputs = [];
  const g = context({
    Session: { getActiveUser: () => ({ getEmail: () => email }) },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => (Object.prototype.hasOwnProperty.call(properties, key) ? properties[key] : null),
        setProperty: (key, value) => { properties[key] = value; }
      })
    },
    HtmlService: {
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createTemplateFromFile: name => ({
        evaluate: () => {
          const out = { file: name, title: '', meta: {}, frame: 'DEFAULT' };
          out.setTitle = value => { out.title = value; return out; };
          out.addMetaTag = (key, value) => { out.meta[key] = value; return out; };
          out.setXFrameOptionsMode = value => { out.frame = value; return out; };
          return out;
        }
      })
    },
    ContentService: {
      MimeType: { JSON: 'application/json', JAVASCRIPT: 'application/javascript' },
      createTextOutput: text => {
        const out = { text, mime: '' };
        out.setMimeType = value => { out.mime = value; return out; };
        outputs.push(out);
        return out;
      }
    }
  }).load('AppConfig.gs', 'WebAccess.gs', 'WebAppIntegration.gs');
  return { g, outputs };
}

test('doGet still serves the Apps Script web app when no launcher probe is requested', () => {
  const { g } = gas();
  const output = g.doGet({ parameter: {} });
  assert.equal(output.file, 'WebApp');
  assert.equal(output.title, 'Lesson Grader');
  assert.equal(g.doGet().file, 'WebApp');
  assert.equal(g.doGet({ parameter: { usp: 'sharing' } }).file, 'WebApp');
  // HtmlService ignores an in-file viewport tag; doGet must add it (REVIEW.md F05).
  assert.equal(output.meta.viewport, 'width=device-width, initial-scale=1');
  // Framing by other origins stays blocked unless the owner opts in (F34).
  assert.equal(output.frame, 'DEFAULT');
});

test('owner can opt in to launcher embedding; default stays non-embeddable', () => {
  const { g } = gas({ properties: { LESSON_GRADER_ALLOW_EMBED: 'true' } });
  assert.equal(g.doGet({ parameter: {} }).frame, 'ALLOWALL');
});

test('launcher probe answers JSON and JSONP without leaking the key, allowlist or full email', () => {
  const properties = {
    LESSON_GRADER_ALLOWED_EMAILS: 'dan@school.edu.au, other@school.edu.au',
    GEMINI_API_KEY: 'THIS_IS_A_TEST_KEY_NOT_REAL_123456'
  };
  const { g, outputs } = gas({ properties });

  const json = g.doGet({ parameter: { lgapi: 'status' } });
  assert.equal(json.mime, 'application/json');
  const body = JSON.parse(json.text);
  assert.equal(body.success, true);
  assert.equal(body.probe, 'status');
  assert.equal(body.schema, '1');
  assert.equal(body.data.authorised, true);
  assert.equal(body.data.accountDetected, true);
  assert.equal(body.data.account, 'da***@school.edu.au');
  assert.equal(body.data.keyConfigured, true);
  assert.equal(body.data.allowlistConfigured, true);
  assert.ok(!json.text.includes('THIS_IS_A_TEST_KEY_NOT_REAL_123456'), 'key must never leave the server');
  assert.ok(!json.text.includes('dan@school.edu.au'), 'full email must never leave the server');
  assert.ok(!json.text.includes('other@school.edu.au'), 'allowlist must never leave the server');

  const jsonp = g.doGet({ parameter: { lgapi: 'status', callback: 'lgProbe_abc123' } });
  assert.equal(jsonp.mime, 'application/javascript');
  assert.ok(jsonp.text.startsWith('lgProbe_abc123({'));
  assert.ok(jsonp.text.endsWith('})'));
  assert.deepEqual(JSON.parse(jsonp.text.slice('lgProbe_abc123('.length, -1)), body);
  assert.equal(outputs.length, 2);
});

test('launcher probe is a closed allowlist: no caller-chosen function, no callback injection', () => {
  const { g } = gas({ properties: { LESSON_GRADER_ALLOWED_EMAILS: 'dan@school.edu.au' } });
  let gradingRan = false;
  g.LessonGraderWeb = { open: () => { gradingRan = true; } };

  for (const call of ['dropGrades', 'apiWebApprove', 'constructor', 'status ', 'STATUS']) {
    const output = g.doGet({ parameter: { lgapi: call } });
    const body = JSON.parse(output.text);
    if (call.trim().toLowerCase() === 'status') {
      assert.equal(body.success, true, 'the one supported probe still answers');
    } else {
      assert.equal(body.success, false, `probe "${call}" must be refused`);
      assert.match(body.message, /read-only/);
    }
    assert.equal(output.mime, 'application/json');
  }
  assert.equal(gradingRan, false);

  const injected = g.doGet({ parameter: { lgapi: 'status', callback: 'x);alert(1);//' } });
  assert.equal(injected.mime, 'application/json', 'an unsafe callback name must fall back to plain JSON');
  assert.ok(!injected.text.includes('alert(1)'));
});

test('launcher probe reports an unallowlisted or signed-out teacher without throwing', () => {
  const denied = gas({ email: 'outsider@elsewhere.edu.au', properties: { LESSON_GRADER_ALLOWED_EMAILS: 'dan@school.edu.au' } });
  const deniedBody = JSON.parse(denied.g.doGet({ parameter: { lgapi: 'status' } }).text);
  assert.equal(deniedBody.success, true);
  assert.equal(deniedBody.data.authorised, false);
  assert.equal(deniedBody.data.accountDetected, true);
  assert.equal(deniedBody.data.account, 'ou***@elsewhere.edu.au');
  assert.equal(deniedBody.data.keyConfigured, false);

  const signedOut = gas({ email: '', properties: { LESSON_GRADER_ALLOWED_EMAILS: 'dan@school.edu.au' } });
  const signedOutBody = JSON.parse(signedOut.g.doGet({ parameter: { lgapi: 'status' } }).text);
  assert.equal(signedOutBody.data.authorised, false);
  assert.equal(signedOutBody.data.accountDetected, false);
  assert.equal(signedOutBody.data.account, '');

  const unconfigured = gas({ email: 'dan@school.edu.au', properties: {} });
  const unconfiguredBody = JSON.parse(unconfigured.g.doGet({ parameter: { lgapi: 'status' } }).text);
  assert.equal(unconfiguredBody.data.authorised, false, 'an empty allowlist denies everyone');
  assert.equal(unconfiguredBody.data.allowlistConfigured, false);
});

/* ------------------------------------------------------------------ *
 * Launcher page
 * ------------------------------------------------------------------ */
function makeStorage(initial = {}) {
  const map = new Map(Object.entries(initial));
  return {
    map,
    getItem: key => (map.has(key) ? map.get(key) : null),
    setItem: (key, value) => map.set(key, String(value)),
    removeItem: key => map.delete(key),
    clear: () => map.clear(),
    key: index => [...map.keys()][index] ?? null,
    get length() { return map.size; }
  };
}

function loadPage({ storage = makeStorage(), search = '', fetchImpl = null, prompts = [] } = {}) {
  const appended = [];
  const errors = [];
  const virtualConsole = new VirtualConsole();
  virtualConsole.on('jsdomError', error => errors.push(error.message));
  virtualConsole.on('error', (...args) => errors.push(args.join(' ')));
  const dom = new JSDOM(html, {
    runScripts: 'dangerously',
    url: PAGE_URL + search,
    virtualConsole,
    beforeParse(window) {
      Object.defineProperty(window, 'localStorage', { value: storage, configurable: true });
      window.confirm = () => true;
      window.prompt = (message, value) => { prompts.push(value === undefined ? message : value); return null; };
      window.scrollTo = () => {};
      window.HTMLElement.prototype.scrollIntoView = () => {};
      if (fetchImpl) window.fetch = fetchImpl;
      else delete window.fetch;
      /* Watch for the JSONP probe <script> before the document is parsed. */
      const headAppend = window.HTMLHeadElement.prototype.appendChild;
      window.HTMLHeadElement.prototype.appendChild = function (node) {
        if (node.tagName === 'SCRIPT') appended.push(node);
        return headAppend.call(this, node);
      };
    }
  });
  return {
    dom,
    window: dom.window,
    storage,
    appended,
    prompts,
    errors,
    /* Closing the window cancels the probe timeout so the suite stays fast. */
    close: () => dom.window.close()
  };
}

async function waitFor(check, what) {
  for (let i = 0; i < 120; i++) {
    if (check()) return;
    await new Promise(resolve => setTimeout(resolve, 5));
  }
  throw new Error('Timed out: ' + what);
}

const ALLOWED_PAYLOAD = {
  success: true, probe: 'status', schema: '1',
  data: { app: 'Lesson Grader', authorised: true, accountDetected: true, account: 'da***@school.edu.au', keyConfigured: true, allowlistConfigured: true }
};

function statusOf(window) {
  return {
    state: window.document.getElementById('status').getAttribute('data-state'),
    label: window.document.getElementById('status-label').textContent,
    detail: window.document.getElementById('status').title
  };
}

function storedSettings(storage) {
  const raw = storage.getItem('lessonGrader.launcher.v1');
  return raw ? JSON.parse(raw) : null;
}

function submitSetup(window, url, label = '') {
  window.document.getElementById('setup-url').value = url;
  window.document.getElementById('setup-label').value = label;
  window.document.getElementById('setup-form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true }));
}

function openSettings(window) {
  window.document.getElementById('btn-settings').dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
}

function click(window, id) {
  window.document.getElementById(id).dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
}

/* Answer the next JSONP probe exactly as the Apps Script deployment would. */
async function answerProbe(page, payload) {
  await waitFor(() => page.appended.length, 'a probe request to be sent');
  const script = page.appended[page.appended.length - 1];
  const callback = new URL(script.src).searchParams.get('callback');
  page.window[callback](payload);
  await waitFor(() => statusOf(page.window).state !== 'checking', 'the probe answer to reach the status pill');
  return new URL(script.src);
}

test('the shipped page contains no deployment address, key or teacher list', () => {
  assert.ok(!/macros\/s\/[A-Za-z0-9_-]{10,}/.test(html), 'no hardcoded deployment id');
  assert.ok(!/AIza[0-9A-Za-z_-]{10,}/.test(html), 'no API key shaped string');
  assert.ok(!/@school\.edu\.au/.test(html), 'no teacher address baked into the page');
  assert.ok(!html.includes('google.script.run'), 'the launcher must not pretend to run inside Apps Script');
  assert.match(html, /localStorage/);
});

test('the page boots cleanly and answers with the setup view', t => {
  const page = loadPage();
  t.after(page.close);
  assert.equal(page.errors.length, 0, 'no script errors on load: ' + page.errors.join(' | '));
  assert.equal(typeof page.window.LessonGraderLauncher.start, 'function');
  assert.equal(page.window.document.getElementById('view-setup').hidden, false);
  assert.equal(page.window.document.getElementById('view-help').hidden, false, 'setup guidance is visible before anything is saved');
  assert.equal(page.window.document.querySelectorAll('#help-blocks details').length, 5);
});

test('a fresh visit shows setup and stores nothing until an address is saved', t => {
  const page = loadPage();
  t.after(page.close);
  const { window, storage } = page;
  const doc = window.document;
  assert.equal(doc.getElementById('view-setup').hidden, false);
  assert.equal(doc.getElementById('view-embed').hidden, true);
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), null);
  assert.equal(statusOf(window).label, 'Not connected yet');
  assert.equal(storage.map.size, 0);

  submitSetup(window, '::::');
  assert.equal(doc.getElementById('setup-flash').hidden, false);
  assert.match(doc.getElementById('setup-flash').textContent, /not a web address/i);
  assert.equal(storage.map.size, 0, 'an invalid address must not be saved');

  submitSetup(window, 'http://script.google.com/macros/s/AKfycbTESTDEPLOYMENTID0000000000000000000/exec');
  assert.match(doc.getElementById('setup-flash').textContent, /https/i);
  assert.equal(storage.map.size, 0, 'a non-https address must not be saved');

  submitSetup(window, 'https://script.google.com/macros/s/shortid/edit');
  assert.match(doc.getElementById('setup-flash').textContent, /deployment looks like/i);
  assert.equal(storage.map.size, 0);
});

test('saving an /exec address persists it and mounts the Apps Script app in the frame', async t => {
  const page = loadPage();
  t.after(page.close);
  const { window, storage } = page;
  submitSetup(window, EXEC_URL + '?usp=sharing', 'Live deployment');
  await waitFor(() => window.LessonGraderLauncher.mountedUrl(), 'iframe mounted');

  const saved = storedSettings(storage);
  assert.equal(saved.schema, 1);
  assert.equal(saved.connections.length, 1);
  assert.equal(saved.connections[0].label, 'Live deployment');
  assert.equal(saved.connections[0].url, EXEC_URL + '?usp=sharing', 'the teacher’s own query string is preserved');
  assert.equal(saved.activeId, saved.connections[0].id);
  assert.deepEqual(Object.keys(saved).sort(), ['activeId', 'connections', 'prefs', 'schema'], 'only addresses and preferences are stored');

  const doc = window.document;
  assert.equal(doc.getElementById('view-embed').hidden, false);
  assert.equal(doc.getElementById('view-setup').hidden, true);
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), EXEC_URL + '?usp=sharing');
  assert.equal(doc.getElementById('frame-url').textContent, EXEC_URL + '?usp=sharing');
  assert.equal(doc.getElementById('btn-open').getAttribute('href'), EXEC_URL + '?usp=sharing');
  assert.equal(doc.getElementById('btn-reload').hidden, false);
  assert.equal(window.LessonGraderLauncher.state.settings.prefs.openMode, 'embed');
});

test('the saved address survives a reload without touching the page code', async t => {
  const first = loadPage();
  t.after(first.close);
  submitSetup(first.window, EXEC_URL, 'Live deployment');
  const persisted = first.storage.map.get('lessonGrader.launcher.v1');
  assert.ok(persisted);
  assert.ok(!html.includes('AKfycbTESTDEPLOYMENTID'), 'the address is not in the shipped markup');

  const second = loadPage({ storage: makeStorage({ 'lessonGrader.launcher.v1': persisted }) });
  t.after(second.close);
  const doc = second.window.document;
  assert.equal(doc.getElementById('view-setup').hidden, true);
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), EXEC_URL, 'the frame address comes from storage');
  assert.equal(doc.getElementById('frame-url').textContent, EXEC_URL);
  assert.equal(second.window.LessonGraderLauncher.activeConnection().label, 'Live deployment');
  await answerProbe(second, ALLOWED_PAYLOAD);
  assert.equal(statusOf(second.window).state, 'connected');
});

test('corrupt or foreign storage never breaks the page', t => {
  for (const raw of ['not json', '{"connections": 5}', '{"connections":[{"url":"javascript:alert(1)"}]}']) {
    const page = loadPage({ storage: makeStorage({ 'lessonGrader.launcher.v1': raw }) });
    t.after(page.close);
    assert.equal(page.window.document.getElementById('view-setup').hidden, false);
    assert.equal(page.window.LessonGraderLauncher.state.settings.connections.length, 0);
  }
});

test('connection check asks only the read-only probe, over JSONP when fetch is unavailable', async t => {
  const page = loadPage();
  t.after(page.close);
  submitSetup(page.window, EXEC_URL, 'Live deployment');
  const src = await answerProbe(page, ALLOWED_PAYLOAD);

  assert.equal(src.origin, 'https://script.google.com');
  assert.equal(src.pathname, '/macros/s/AKfycbTESTDEPLOYMENTID0000000000000000000/exec');
  assert.equal(src.searchParams.get('lgapi'), 'status');
  assert.match(src.searchParams.get('callback'), /^lgProbe_[a-z0-9]+$/);
  assert.ok(src.searchParams.get('_').length > 0, 'cache buster present');

  assert.equal(statusOf(page.window).state, 'connected');
  assert.equal(statusOf(page.window).label, 'Connected');
  assert.match(statusOf(page.window).detail, /da\*\*\*@school\.edu\.au/);
  assert.equal(page.appended.length, 1, 'one request per check');
  assert.equal(page.window.document.head.querySelectorAll('script[src*="lgapi"]').length, 0, 'the probe script is removed afterwards');
});

test('connection check explains an allowlist denial, a signed-out browser and an unreachable deployment', async t => {
  const denied = loadPage();
  t.after(denied.close);
  submitSetup(denied.window, EXEC_URL);
  await answerProbe(denied, { success: true, data: { authorised: false, accountDetected: true, account: 'ou***@elsewhere.edu.au', keyConfigured: false, allowlistConfigured: true } });
  assert.equal(statusOf(denied.window).state, 'denied');
  assert.match(statusOf(denied.window).detail, /LESSON_GRADER_ALLOWED_EMAILS/);

  const signedOut = loadPage();
  t.after(signedOut.close);
  submitSetup(signedOut.window, EXEC_URL);
  await answerProbe(signedOut, { success: true, data: { authorised: false, accountDetected: false, account: '', keyConfigured: false, allowlistConfigured: false } });
  assert.equal(statusOf(signedOut.window).state, 'anonymous');
  assert.match(statusOf(signedOut.window).detail, /own tab/);

  const offline = loadPage();
  t.after(offline.close);
  submitSetup(offline.window, EXEC_URL);
  await waitFor(() => offline.appended.length, 'a probe request to be sent');
  offline.appended[0].dispatchEvent(new offline.window.Event('error'));
  await waitFor(() => statusOf(offline.window).state === 'unreachable', 'unreachable status');
  assert.match(statusOf(offline.window).detail, /own tab/);
});

test('a refused or unreadable probe answer is reported, not treated as success', t => {
  const page = loadPage();
  t.after(page.close);
  const describe = page.window.LessonGraderLauncher.describeProbe;
  assert.equal(describe({ success: false, message: 'Unknown launcher probe "dropGrades".' }).state, 'rejected');
  assert.equal(describe(null).state, 'unreachable');
  assert.equal(describe('<html>sign in</html>').state, 'unreachable');
  assert.equal(describe({ success: true, data: {} }).state, 'anonymous');
});

test('a cross-origin GET is preferred when the deployment answers with CORS headers', async t => {
  const calls = [];
  const fetchImpl = url => {
    calls.push(url);
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(ALLOWED_PAYLOAD) });
  };
  const page = loadPage({ fetchImpl });
  t.after(page.close);
  submitSetup(page.window, EXEC_URL);
  await waitFor(() => statusOf(page.window).state === 'connected', 'status from fetch');
  assert.equal(calls.length, 1);
  const requested = new URL(calls[0]);
  assert.equal(requested.searchParams.get('lgapi'), 'status');
  assert.equal(requested.searchParams.get('callback'), null, 'no JSONP callback on the fetch attempt');
  assert.equal(page.appended.length, 0, 'no JSONP fallback needed');
});

test('the probe is never sent to a host that is not script.google.com', async t => {
  const page = loadPage();
  t.after(page.close);
  submitSetup(page.window, 'https://example.com/macros/s/AKfycbTESTDEPLOYMENTID0000000000000000000/exec');
  await waitFor(() => statusOf(page.window).state === 'unchecked', 'probe skipped');
  assert.equal(page.appended.length, 0);
  assert.equal(page.window.document.getElementById('view-embed').hidden, true, 'untrusted hosts never run inside the frame');
  assert.equal(page.window.document.getElementById('view-open').hidden, false);
  assert.match(statusOf(page.window).detail, /only sent to script\.google\.com/);
});

test('settings can update, switch between and delete saved deployments', async t => {
  const page = loadPage();
  t.after(page.close);
  const { window, storage } = page;
  submitSetup(window, EXEC_URL, 'Live deployment');
  await waitFor(() => window.LessonGraderLauncher.mountedUrl(), 'first mount');

  const launcher = window.LessonGraderLauncher;
  const doc = window.document;
  openSettings(window);
  assert.equal(doc.getElementById('settings-backdrop').hidden, false);
  assert.equal(doc.querySelectorAll('#conn-list .conn').length, 1);

  const second = launcher.upsertConnection(ALT_URL, 'Testing deployment', '');
  assert.equal(second.ok, true);
  assert.equal(doc.querySelectorAll('#conn-list .conn').length, 2);
  assert.equal(storedSettings(storage).connections.length, 2);
  assert.equal(launcher.activeConnection().label, 'Testing deployment');

  const duplicate = launcher.upsertConnection(ALT_URL, 'Again', '');
  assert.equal(duplicate.ok, false);
  assert.match(duplicate.error, /already saved/);

  const liveId = storedSettings(storage).connections.filter(item => item.url === EXEC_URL)[0].id;
  launcher.activate(liveId);
  assert.equal(storedSettings(storage).activeId, liveId);
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), EXEC_URL);

  const updated = launcher.upsertConnection(DEV_URL, '', liveId);
  assert.equal(updated.ok, true);
  assert.match(updated.warning, /development address/i);
  assert.equal(storedSettings(storage).connections.length, 2, 'updating must not add a row');
  assert.equal(launcher.activeConnection().label, 'Live deployment', 'an existing label survives an address change');
  await waitFor(() => doc.getElementById('app-frame').getAttribute('src') === DEV_URL, 'the frame follows the new address');

  launcher.removeConnection(launcher.activeConnection().id);
  assert.equal(storedSettings(storage).connections.length, 1);
  assert.equal(launcher.activeConnection().url, ALT_URL, 'another saved deployment takes over');

  launcher.removeConnection(launcher.activeConnection().id);
  assert.equal(doc.getElementById('view-setup').hidden, false);
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), null);
  assert.equal(storedSettings(storage).connections.length, 0);
});

test('open-in-tab preference is remembered and reload re-requests the app', async t => {
  const page = loadPage();
  t.after(page.close);
  const { window, storage } = page;
  submitSetup(window, EXEC_URL);
  await waitFor(() => window.LessonGraderLauncher.mountedUrl(), 'mounted');

  const doc = window.document;
  openSettings(window);
  doc.getElementById('mode-tab').checked = true;
  doc.getElementById('mode-tab').dispatchEvent(new window.Event('change', { bubbles: true }));
  assert.equal(doc.getElementById('view-open').hidden, false);
  assert.equal(doc.getElementById('view-embed').hidden, true);
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), null);
  assert.equal(doc.getElementById('open-big').getAttribute('href'), EXEC_URL);
  assert.match(doc.getElementById('open-details').textContent, /New browser tab/);
  assert.equal(storedSettings(storage).prefs.openMode, 'tab');

  click(window, 'open-embed');
  await waitFor(() => doc.getElementById('app-frame').getAttribute('src'), 're-embedded');
  assert.equal(storedSettings(storage).prefs.openMode, 'embed');

  await waitFor(() => doc.getElementById('app-frame').getAttribute('src'), 'mounted again');
  click(window, 'frame-tab');
  assert.equal(doc.getElementById('view-open').hidden, false, 'the frame bar can escape to a tab');
  assert.equal(storedSettings(storage).prefs.openMode, 'tab');
  assert.match(doc.getElementById('toast').textContent, /own tab/);
  click(window, 'open-embed');
  await waitFor(() => doc.getElementById('app-frame').getAttribute('src'), 'back in the frame');

  const before = doc.getElementById('app-frame').getAttribute('src');
  click(window, 'btn-reload');
  const after = doc.getElementById('app-frame').getAttribute('src');
  assert.notEqual(before, after);
  assert.match(after, /lgr=\d+$/);
});

test('settings export, import and clear round-trip without secrets', t => {
  const page = loadPage();
  t.after(page.close);
  const { window, storage, prompts } = page;
  submitSetup(window, EXEC_URL, 'Live deployment');
  const launcher = window.LessonGraderLauncher;
  const doc = window.document;

  openSettings(window);
  click(window, 'data-export');
  assert.equal(prompts.length, 1);
  const exported = JSON.parse(prompts[0]);
  assert.equal(exported.connections[0].url, EXEC_URL);
  assert.deepEqual(Object.keys(exported).sort(), ['activeId', 'connections', 'prefs', 'schema']);
  assert.ok(!JSON.stringify(exported).includes('key'), 'no secret-shaped fields');

  const cleared = loadPage();
  t.after(cleared.close);
  const imported = cleared.window.LessonGraderLauncher.importSettings(JSON.stringify(exported));
  assert.equal(imported.ok, true);
  assert.equal(imported.count, 1);
  assert.equal(cleared.window.document.getElementById('app-frame').getAttribute('src'), EXEC_URL);
  assert.equal(cleared.window.LessonGraderLauncher.importSettings('nope').ok, false);
  assert.equal(cleared.window.LessonGraderLauncher.importSettings('{"connections":[{"url":"https://example.org/"}]}').ok, false);
  assert.match(cleared.window.LessonGraderLauncher.shareLink(), /gas=https%3A%2F%2Fscript\.google\.com/);

  launcher.removeConnection(launcher.activeConnection().id);
  assert.equal(storedSettings(storage).connections.length, 0);
  assert.equal(doc.getElementById('view-setup').hidden, false);
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), null);
});

test('a share link prefills the address but never saves it silently', t => {
  const linked = 'gas=' + encodeURIComponent(EXEC_URL) + '&label=' + encodeURIComponent('Shared by a colleague');
  const page = loadPage({ search: '?' + linked });
  t.after(page.close);
  const doc = page.window.document;
  assert.equal(doc.getElementById('setup-url').value, EXEC_URL);
  assert.equal(doc.getElementById('setup-label').value, 'Shared by a colleague');
  assert.match(doc.getElementById('setup-flash').textContent, /Check it matches your own Apps Script project/);
  assert.equal(page.storage.map.size, 0, 'nothing is stored until the teacher saves it');
  assert.equal(doc.getElementById('app-frame').getAttribute('src'), null);
  assert.ok(!page.window.location.search.includes('gas='), 'the address is scrubbed from the visible URL');

  const existing = loadPage({
    storage: makeStorage({
      'lessonGrader.launcher.v1': JSON.stringify({
        schema: 1,
        connections: [{ id: 'c1', label: 'Live', url: ALT_URL }],
        activeId: 'c1',
        prefs: { openMode: 'embed', autoProbe: false }
      })
    }),
    search: '?' + linked
  });
  t.after(existing.close);
  assert.equal(existing.window.document.getElementById('settings-backdrop').hidden, false, 'existing users get the settings dialog to review it');
  assert.equal(existing.window.document.getElementById('conn-url').value, EXEC_URL);
  assert.equal(storedSettings(existing.storage).connections.length, 1, 'still nothing added');
});

test('the settings dialog is keyboard accessible and closing it returns focus', t => {
  const page = loadPage();
  t.after(page.close);
  const doc = page.window.document;
  const button = doc.getElementById('btn-settings');
  button.focus();
  click(page.window, 'btn-settings');
  assert.equal(doc.getElementById('settings-backdrop').hidden, false);
  assert.equal(doc.getElementById('settings-dialog').getAttribute('aria-modal'), 'true');
  assert.equal(doc.body.style.overflow, 'hidden');

  doc.dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  assert.equal(doc.getElementById('settings-backdrop').hidden, true);
  assert.equal(doc.body.style.overflow, '');
  assert.equal(doc.activeElement, button);
});
