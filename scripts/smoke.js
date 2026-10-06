import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { CDP, Page, findBrowser } from '../lib/cdp.js';
import { teamsDOM, DEFAULT_SELECTORS } from '../lib/dom.js';
import { fixtureHTML, runDOMSuite } from '../test/dom-suite.js';
import { createAdapter } from '../index.js';

// Runs outside the Codex sandbox on the user's Mac. It sends only fixture
// messages to a blank local page, never to Teams or any external recipient.
const directory = await fs.mkdtemp('/tmp/teams-smoke-');
let cdp, adapter;
try {
  cdp = new CDP(await findBrowser(process.env.TEAMS_BROWSER_PATH), directory + '/browser', { headless: true });
  const page = await cdp.page('about:blank');
  assert.ok(page instanceof Page);
  console.log(await runDOMSuite(page, teamsDOM, DEFAULT_SELECTORS, fixtureHTML));
  adapter = await createAdapter({ dataDir: directory + '/adapter', settings: {} }, {
    createBrowser: () => ({ serial: fn => fn(), login: async () => ({}), dom: async () => ({ ready: true }), close: async () => {} }),
  });
  const service = adapter.services[0], auth = service.authMethods[0];
  const flow = await auth.connect({ config: { label: 'Smoke fixture' } });
  const conn = await auth.poll({ pending: flow.pending });
  const req = { method: 'GET', url: new URL('/status', service.baseUrl), headers: new Headers() };
  auth.authorize(req, conn);
  const response = await fetch(req.url, { headers: req.headers });
  assert.equal(response.status, 200);
  assert.equal((await response.json()).ready, true);
  assert.equal((await fetch(req.url, { headers: req.headers })).status, 401);
  console.log('Private Chrome pipe and real loopback HTTP smoke tests passed.');
} finally {
  await adapter?.dispose();
  await cdp?.close();
  await fs.rm(directory, { recursive: true, force: true });
}
