import http from 'node:http';
import { randomBytes, randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import { TeamsBrowser, profilePath } from './lib/browser.js';
import { SendLedger } from './lib/ledger.js';
import { AdapterError } from './lib/errors.js';
import { openapi } from './lib/openapi.js';

function integer(url, name, fallback, min, max) {
  const raw = url.searchParams.get(name);
  if (raw === null) return fallback;
  if (!/^\d+$/.test(raw)) throw new AdapterError('invalid_query', `${name} must be an integer between ${min} and ${max}.`, 400);
  const value = Number(raw);
  if (!Number.isInteger(value) || value < min || value > max) throw new AdapterError('invalid_query', `${name} must be between ${min} and ${max}.`, 400);
  return value;
}
async function body(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > 128 * 1024) throw new AdapterError('body_too_large', 'Maximum JSON body is 128 KiB.', 413);
    chunks.push(chunk);
  }
  let value;
  try { value = JSON.parse(Buffer.concat(chunks).toString()); }
  catch { throw new AdapterError('invalid_json', 'Body must be a JSON object.', 400); }
  if (!value || Array.isArray(value) || typeof value !== 'object') throw new AdapterError('invalid_json', 'Body must be a JSON object.', 400);
  if (Object.keys(value).some(k => !['text', 'idempotencyKey'].includes(k))) throw new AdapterError('invalid_body', 'Only text and idempotencyKey are accepted.', 400);
  return value;
}

export async function createAdapter(ctx, {
  createBrowser = (dir, settings) => new TeamsBrowser(dir, settings),
  createServer = handler => http.createServer(handler),
} = {}) {
  const browsers = new Map();
  const pending = new Map();
  const invocations = new Map();
  let disposed = false;
  const getBrowser = profile => {
    const dir = profilePath(ctx.dataDir, profile);
    if (!browsers.has(profile)) browsers.set(profile, createBrowser(dir, ctx.settings));
    return browsers.get(profile);
  };
  const prune = setInterval(() => {
    const now = Date.now();
    for (const [token, value] of invocations) if (value.expires < now) invocations.delete(token);
    for (const [profile, value] of pending) if (value.expires < now) {
      pending.delete(profile);
      const browser = browsers.get(profile);
      browsers.delete(profile);
      browser?.close().catch(() => {});
    }
  }, 30_000);
  prune.unref();
  const respond = (response, status, value) => {
    response.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' });
    response.end(JSON.stringify(value));
  };
  const server = createServer(async (request, response) => {
    const token = String(request.headers.authorization ?? '').replace(/^Bearer\s+/i, '');
    const invocation = invocations.get(token);
    invocations.delete(token);
    if (!invocation || invocation.expires < Date.now()) return respond(response, 401, { error: { code: 'invalid_invocation', message: 'Invalid or expired invocation.' } });
    try {
      const url = new URL(request.url ?? '/', 'http://localhost');
      if (request.method !== invocation.method || url.pathname !== invocation.pathname || url.search !== invocation.search) throw new AdapterError('invocation_mismatch', 'Invocation is bound to a different request.', 403);
      const browser = getBrowser(invocation.profile);
      const match = url.pathname.match(/^\/chats\/([^/]+)\/messages$/);
      let operation;
      if (request.method === 'GET' && url.pathname === '/status') operation = () => browser.dom('status');
      else if (request.method === 'POST' && url.pathname === '/login') operation = () => browser.login();
      else if (request.method === 'GET' && url.pathname === '/diagnostics') operation = () => browser.dom('diagnostics');
      else if (request.method === 'GET' && url.pathname === '/chats') operation = () => browser.listChats();
      else if (request.method === 'GET' && match) {
        const id = decodeURIComponent(match[1]);
        const limit = integer(url, 'limit', 50, 1, 200), olderPages = integer(url, 'olderPages', 0, 0, 5);
        operation = () => browser.messages(id, limit, olderPages);
      } else if (request.method === 'POST' && match) {
        const id = decodeURIComponent(match[1]);
        const input = await body(request);
        operation = () => new SendLedger(profilePath(ctx.dataDir, invocation.profile)).send(
          input.idempotencyKey, id, input.text,
          async () => {
            const prepared = await browser.prepare(id);
            if (response.destroyed) throw new AdapterError('client_disconnected', 'Caller disconnected before sending.', 499);
            return prepared;
          },
          prepared => browser.send(prepared, input.text),
        );
      } else throw new AdapterError('not_found', 'Unknown operation.', 404);
      respond(response, 200, await browser.serial(() => {
        if (response.destroyed) throw new AdapterError('client_disconnected', 'Caller disconnected before execution.', 499);
        return operation();
      }));
    } catch (error) {
      respond(response, error.status ?? 500, { error: { code: error.code ?? 'adapter_error', message: error instanceof AdapterError ? error.message : 'Local Teams adapter failed. Check browser and plugin configuration.' } });
    }
  });
  server.requestTimeout = 15_000;
  try { await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); }); }
  catch (error) { clearInterval(prune); throw error; }
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  return {
    services: [{
      id: 'teams-web', name: 'Teams Web', icon: 'icon.svg',
      description: 'Read and reply to chats in a dedicated local browser. No Graph API.',
      // Stable logical URL: saved calls survive reloads. authorize rewrites it
      // to the current loopback adapter before any network request is made.
      baseUrl: 'http://teams.localhost', allowedHosts: ['teams.localhost', new URL(baseUrl).host], openapi,
      authMethods: [{
        id: 'browser', name: 'Sign in in a local browser',
        description: 'Opens Chrome/Edge on the machine running this Switchboard instance. Sign in manually; the browser profile retains your session.',
        fields: [{ key: 'label', label: 'Account label', required: true, placeholder: 'Work Teams', description: 'A name for this dedicated browser profile.' }],
        async connect({ config, connection }) {
          if (disposed) throw new AdapterError('disposed', 'Plugin has been unloaded.', 503);
          const profile = connection?.credentials?.profile ?? randomUUID();
          const label = String(config.label ?? connection?.account?.label ?? 'Teams Web').trim().slice(0, 100);
          const browser = getBrowser(profile);
          await browser.serial(() => browser.login());
          const expires = Date.now() + 15 * 60_000;
          pending.set(profile, { label, expires });
          return {
            device: { userCode: 'LOCAL-BROWSER', verificationUri: TEAMS_URL, expiresIn: 900, interval: 3 },
            pending: { profile, label, expires },
          };
        },
        async poll({ pending: flow }) {
          const active = pending.get(flow.profile);
          if (!active || active.expires < Date.now()) throw new AdapterError('login_expired', 'Sign-in expired. Start connecting again.');
          const browser = getBrowser(flow.profile);
          const status = await browser.serial(() => browser.dom('status'));
          if (!status.ready) return { wait: true };
          pending.delete(flow.profile);
          return { credentials: { profile: flow.profile }, config: {}, account: { id: flow.profile, label: active.label } };
        },
        authorize(request, connection) {
          if (disposed) throw new AdapterError('disposed', 'Plugin has been unloaded.', 503);
          profilePath(ctx.dataDir, connection.credentials.profile);
          // Resolve the logical service URL without a DNS/network request.
          request.url.host = new URL(baseUrl).host;
          request.url.protocol = 'http:';
          const token = randomBytes(32).toString('base64url');
          invocations.set(token, { profile: connection.credentials.profile, method: request.method, pathname: request.url.pathname, search: request.url.search, expires: Date.now() + 10_000 });
          request.headers.set('authorization', `Bearer ${token}`);
        },
        async revoke(connection) {
          const profile = connection.credentials.profile;
          const dir = profilePath(ctx.dataDir, profile);
          pending.delete(profile);
          const browser = browsers.get(profile);
          browsers.delete(profile);
          for (const [token, value] of invocations) if (value.profile === profile) invocations.delete(token);
          await browser?.close();
          await fs.rm(dir, { recursive: true, force: true });
        },
      }],
    }],
    async dispose() {
      disposed = true; clearInterval(prune); invocations.clear(); pending.clear();
      await Promise.all([...browsers.values()].map(browser => browser.close()));
      await new Promise(resolve => { server.close(resolve); server.closeIdleConnections(); });
    },
  };
}

const TEAMS_URL = 'https://teams.microsoft.com/v2/';
export default function setup(ctx) { return createAdapter(ctx); }
