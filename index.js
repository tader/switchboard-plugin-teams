import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomBytes, randomUUID } from 'node:crypto';
import { TeamsSession, importRequired, readyCredentials } from './lib/api-auth.js';
import { TeamsAPI, capabilities, unsupported } from './lib/teams-api.js';
import { SendLedger } from './lib/ledger.js';
import { AdapterError } from './lib/errors.js';
import { messageText, messageId, recipientEmail, reactionInput } from './lib/input.js';
import { richContent } from './lib/content.js';
import { openapi } from './lib/api-openapi.js';
import { validateTokenBundle } from './lib/token-bundle.js';

const profilePath = (dataDir, profile) => {
  if (typeof profile !== 'string' || !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(profile)) throw new AdapterError('invalid_profile', 'Invalid connection profile.', 400);
  return path.join(dataDir, 'profiles', profile);
};
function integer(url, name, fallback, min, max) {
  const raw = url.searchParams.get(name);
  if (raw === null) return fallback;
  if (!/^\d+$/.test(raw) || Number(raw) < min || Number(raw) > max) throw new AdapterError('invalid_query', `${name} must be between ${min} and ${max}.`, 400);
  return Number(raw);
}
function boolean(url, name) {
  const raw = url.searchParams.get(name);
  if (raw === null) return false;
  if (!['true', 'false'].includes(raw)) throw new AdapterError('invalid_query', `${name} must be true or false.`, 400);
  return raw === 'true';
}
async function body(request, allowed) {
  let size = 0; const chunks = [];
  for await (const chunk of request) {
    size += Buffer.byteLength(chunk);
    if (size > 128 * 1024) throw new AdapterError('body_too_large', 'Maximum body size is 128 KiB.', 413);
    chunks.push(Buffer.from(chunk));
  }
  let value;
  try { value = JSON.parse(Buffer.concat(chunks).toString()); }
  catch { throw new AdapterError('invalid_json', 'Body must be JSON.', 400); }
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).some(key => !allowed.includes(key))) throw new AdapterError('invalid_body', `Only ${allowed.join(', ')} are accepted.`, 400);
  return value;
}
function contentInput(input) {
  if ((input.text === undefined) === (input.content === undefined)) throw new AdapterError('invalid_body', 'Provide exactly one of text or content.', 400);
  if (input.replyToMessageId !== undefined) unsupported();
  return { ...input, ...(input.content === undefined ? { text: messageText(input.text) } : { content: richContent(input.content) }) };
}

export async function createAdapter(ctx, {
  createSession = (profile, dir, settings) => new TeamsSession(profile, dir, settings),
  createAPI = session => new TeamsAPI(session),
  createServer = handler => http.createServer(handler),
} = {}) {
  const sessions = new Map(), pending = new Map(), invocations = new Map(); let disposed = false;
  const get = profile => {
    if (disposed) throw new AdapterError('disposed', 'Plugin has been unloaded.', 503);
    const directory = profilePath(ctx.dataDir, profile);
    if (!sessions.has(profile)) {
      const session = createSession(profile, directory, ctx.settings);
      sessions.set(profile, { session, api: createAPI(session), ledger: new SendLedger(directory) });
    }
    return sessions.get(profile);
  };
  const prune = setInterval(() => {
    const now = Date.now();
    for (const [token, value] of invocations) if (value.expires < now) invocations.delete(token);
    for (const [profile, flow] of pending) if (flow.expires < now) {
      pending.delete(profile); flow.entry.session.close().catch(() => {});
    }
  }, 30_000); prune.unref();
  const respond = (response, status, value) => {
    if (response.destroyed) return;
    response.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' });
    response.end(JSON.stringify(value));
  };
  const server = createServer(async (request, response) => {
    const token = String(request.headers.authorization ?? '').replace(/^Bearer\s+/i, '');
    const invocation = invocations.get(token); invocations.delete(token);
    if (!invocation || invocation.expires < Date.now()) return respond(response, 401, { error: { code: 'invalid_invocation', message: 'Invalid or expired invocation.' } });
    let activeSession;
    try {
      const url = new URL(request.url, 'http://localhost');
      if (invocation.method !== request.method || invocation.pathname !== url.pathname || invocation.search !== url.search) throw new AdapterError('invocation_mismatch', 'Invocation does not match this request.', 403);
      const { session, api, ledger } = get(invocation.profile); activeSession = session;
      const checkCancelled = () => { if (response.destroyed) throw new AdapterError('client_disconnected', 'Caller disconnected before the operation.', 499); };
      checkCancelled();
      if (request.method === 'GET' && ['/status', '/diagnostics'].includes(url.pathname)) return respond(response, 200, { ...session.status(), capabilities });
      if (request.method === 'GET' && url.pathname === '/capabilities') return respond(response, 200, capabilities);
      if (request.method === 'POST' && url.pathname === '/login') {
        if (pending.has(invocation.profile)) throw new AdapterError('signin_pending', 'Complete the pending connection sign-in.', 503);
        return respond(response, 200, await session.login());
      }
      const messages = url.pathname.match(/^\/chats\/([^/]+)\/messages$/);
      const mutation = url.pathname.match(/^\/chats\/([^/]+)\/messages\/([^/]+)(\/reactions)?$/);
      const history = url.pathname.match(/^\/(chats|channels)\/([^/]+)\/history$/);
      let operation;
      if (request.method === 'GET' && ['/chats', '/unread/chats'].includes(url.pathname)) {
        const unread = url.pathname === '/unread/chats' || boolean(url, 'unreadOnly'); operation = () => api.chats(unread);
      } else if (request.method === 'GET' && url.pathname === '/channels') operation = () => api.channels();
      else if (request.method === 'GET' && url.pathname === '/unread/messages') {
        const max = integer(url, 'maxChats', 3, 1, 5), limit = integer(url, 'limitPerChat', 50, 1, 100); operation = () => api.unread(max, limit);
      } else if (request.method === 'GET' && messages) {
        const id = decodeURIComponent(messages[1]), limit = integer(url, 'limit', 50, 1, 200), older = integer(url, 'olderPages', 0, 0, 5); operation = () => api.messages(id, limit, older);
      } else if (request.method === 'GET' && history) {
        const id = decodeURIComponent(history[2]), limit = integer(url, 'limit', 100, 1, 200), format = url.searchParams.get('format') ?? 'json', cursor = url.searchParams.get('cursor') ?? undefined;
        if (!['json', 'ndjson'].includes(format) || cursor && !/^[a-f0-9-]{36}$/.test(cursor)) throw new AdapterError('invalid_query', 'Use a returned cursor and json or ndjson format.', 400);
        operation = () => api.history(id, { limit, format, cursor, kind: history[1] === 'chats' ? 'chat' : 'channel' });
      } else if (request.method === 'GET' && url.pathname === '/people') {
        const email = recipientEmail(url.searchParams.get('query')); operation = () => api.people(email);
      } else if (request.method === 'POST' && messages) {
        const id = decodeURIComponent(messages[1]), input = contentInput(await body(request, ['text', 'content', 'replyToMessageId', 'idempotencyKey']));
        operation = () => ledger.run(input.idempotencyKey, ['api-send', id, input], () => api.prepareSend(id, input), prepared => { checkCancelled(); return api.send(prepared); }, 'send_uncertain');
      } else if (request.method === 'POST' && url.pathname === '/conversations') {
        const input = await body(request, ['email', 'text', 'idempotencyKey']); input.email = recipientEmail(input.email);
        if (input.text === undefined) {
          if (input.idempotencyKey !== undefined) throw new AdapterError('invalid_body', 'idempotencyKey requires text.', 400);
          operation = () => api.existingConversation(input.email);
        } else {
          input.text = messageText(input.text);
          operation = () => ledger.run(input.idempotencyKey, ['api-person-send', input], async () => {
            const conversation = await api.existingConversation(input.email);
            return { ...await api.prepareSend(conversation.chatId, input), conversation };
          }, async prepared => { checkCancelled(); return { ...prepared.conversation, ...await api.send(prepared), messageSent: true }; }, 'send_uncertain');
        }
      } else if (mutation) {
        const id = decodeURIComponent(mutation[1]), mid = messageId(decodeURIComponent(mutation[2]));
        if (request.method === 'GET' && mutation[3]) operation = () => api.reactions(id, mid);
        else {
          const kind = request.method === 'POST' && mutation[3] ? 'reaction' : request.method === 'PATCH' && !mutation[3] ? 'edit' : request.method === 'DELETE' && !mutation[3] ? 'delete' : null;
          if (!kind) throw new AdapterError('not_found', 'Unknown Teams operation.', 404);
          let input = await body(request, kind === 'reaction' ? ['reaction', 'selected', 'idempotencyKey'] : kind === 'edit' ? ['text', 'content', 'expectedText', 'idempotencyKey'] : ['expectedText', 'idempotencyKey']);
          if (kind === 'reaction') input = { ...input, ...reactionInput(input.reaction, input.selected) };
          else { input.expectedText = messageText(input.expectedText, 'expectedText', true); if (kind === 'edit') input = contentInput(input); }
          operation = () => ledger.mutate(input.idempotencyKey, ['api', kind, id, mid], input, () => api.prepareMutation(id, mid, kind, input), prepared => { checkCancelled(); return api.mutate(prepared); });
        }
      } else if (url.pathname.startsWith('/search/') || /\/(threads|unread)(\/|$)/.test(url.pathname)) unsupported();
      else throw new AdapterError('not_found', 'Unknown Teams operation.', 404);
      const result = await session.serial(() => { checkCancelled(); return operation(); });
      respond(response, 200, result);
    } catch (error) {
      if (activeSession?.credentials?.authMode === 'tokens' && error.code === 'api_unauthorized') { error = importRequired(); activeSession.authError = error; }
      respond(response, error.status ?? 500, { error: { code: error instanceof AdapterError ? error.code : 'adapter_error', message: error instanceof AdapterError ? error.message : 'Local Teams API adapter failed.' } });
    }
  });
  server.requestTimeout = 15_000;
  try { await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); }); }
  catch (error) { clearInterval(prune); throw error; }
  const origin = `http://127.0.0.1:${server.address().port}`;
  const entryFor = profile => {
    const directory = profilePath(ctx.dataDir, profile);
    const session = createSession(profile, directory, ctx.settings);
    return { session, api: createAPI(session), ledger: new SendLedger(directory) };
  };
  const replaceEntry = async (profile, entry) => {
    if (disposed) throw new AdapterError('disposed', 'Plugin has been unloaded.', 503);
    const flow = pending.get(profile);
    if (flow) { pending.delete(profile); await flow.entry.session.close(); }
    await sessions.get(profile)?.session.close();
    if (disposed) throw new AdapterError('disposed', 'Plugin has been unloaded.', 503);
    for (const [token, value] of invocations) if (value.profile === profile) invocations.delete(token);
    sessions.set(profile, entry);
  };
  const sharedAuth = {
    async authorize(request, connection, { force = false } = {}) {
      const { session } = get(connection.credentials.profile); session.adopt(connection.credentials);
      const control = ['/status', '/diagnostics', '/capabilities', '/login'].includes(request.url.pathname);
      if (!control && pending.has(session.profile)) throw new AdapterError('signin_pending', 'Complete the pending connection sign-in.', 503);
      if (!control) await session.ensure(force);
      request.url.host = new URL(origin).host; request.url.protocol = 'http:';
      const token = randomBytes(32).toString('base64url');
      invocations.set(token, { profile: session.profile, method: request.method, pathname: request.url.pathname, search: request.url.search, expires: Date.now() + 10_000 });
      request.headers.set('authorization', `Bearer ${token}`);
      // Switchboard encrypts these credentials; API tokens never leave the local adapter.
      if (session.credentials && session.credentials !== connection.credentials) return { credentials: session.credentials };
    },
    async revoke(connection) {
      const profile = connection.credentials.profile, directory = profilePath(ctx.dataDir, profile);
      const flow = pending.get(profile); pending.delete(profile); await flow?.entry.session.close(); const entry = sessions.get(profile); sessions.delete(profile);
      for (const [token, value] of invocations) if (value.profile === profile) invocations.delete(token);
      await entry?.session.close(); await fs.rm(directory, { recursive: true, force: true, maxRetries: 3 });
    },
  };
  return {
    services: [{ id: 'teams-web', name: 'Teams', icon: 'icon.svg', description: 'Teams APIs with local browser sign-in or imported tokens.',
      baseUrl: 'http://teams.localhost', allowedHosts: ['teams.localhost', new URL(origin).host], openapi,
      authMethods: [{ id: 'browser', name: 'Microsoft sign-in through local Chrome',
        description: 'Sign in once in a dedicated local Chrome profile. Tokens are stored encrypted by Switchboard. Chrome closes after authentication; API calls renew tokens automatically when possible.',
        fields: [{ key: 'label', label: 'Account label', required: true }, { key: 'loginHint', label: 'Microsoft work email', description: 'Optional account hint for sign-in.' }],
        async connect({ config, connection }) {
          const profile = connection?.credentials?.profile ?? randomUUID();
          const previousFlow = pending.get(profile);
          if (previousFlow) { await previousFlow.entry.session.close(); pending.delete(profile); }
          const entry = entryFor(profile), { session } = entry;
          if (connection?.credentials) session.adopt({ ...connection.credentials, authMode: 'browser' });
          const label = String(config.label || connection?.account?.label || 'Teams').trim().slice(0, 100);
          const loginHint = config.loginHint ? recipientEmail(config.loginHint) : undefined;
          const expires = Date.now() + 15 * 60_000, attempt = randomUUID(); pending.set(profile, { label, expires, entry, attempt });
          try {
            // Finish any renewal on the same browser profile before launching sign-in.
            await sessions.get(profile)?.session.queue;
            await session.login(loginHint);
          } catch (error) { pending.delete(profile); await session.close(); throw error; }
          return { device: { userCode: 'LOCAL-BROWSER', verificationUri: 'https://login.microsoftonline.com/', expiresIn: 900, interval: 3 }, pending: { profile, expires, attempt } };
        },
        async poll({ pending: flow }) {
          const active = pending.get(flow.profile);
          if (!active || active.attempt !== flow.attempt || active.expires < Date.now()) throw new AdapterError('login_expired', 'Sign-in expired. Reconnect to try again.', 503);
          const { session } = active.entry;
          if (session.authError) { pending.delete(flow.profile); await session.close(); throw session.authError; }
          if (!session.status().ready) return { wait: true };
          pending.delete(flow.profile);
          await replaceEntry(flow.profile, active.entry);
          return { credentials: session.credentials, config: {}, account: { id: flow.profile, label: active.label } };
        },
        ...sharedAuth,
      }, {
        id: 'tokens', name: 'Import Teams tokens',
        description: 'Paste a bundle from auth:export on a machine with Chrome or Edge. No browser is used here. Re-import when Microsoft access tokens expire.',
        fields: [{ key: 'label', label: 'Account label', required: true }, { key: 'tokenBundle', label: 'Teams token bundle (JSON)', type: 'secret', required: true }],
        async connect({ config, connection }) {
          const profile = connection?.credentials?.profile ?? randomUUID();
          const entry = entryFor(profile);
          try {
            const bundle = await validateTokenBundle(config.tokenBundle, connection?.credentials);
            const credentials = await entry.session.exchange({ ...bundle, profile }, entry.session.fetcher);
            entry.session.adopt(credentials);
            await entry.api.directory(true);
            if (!readyCredentials(credentials)) throw importRequired();
            if (disposed) throw new AdapterError('disposed', 'Plugin has been unloaded.', 503);
            await replaceEntry(profile, entry);
            return { credentials, config: {}, account: { id: profile, label: String(config.label || connection?.account?.label || 'Teams').trim().slice(0, 100) } };
          } catch (error) { await entry.session.close(); throw error; }
        },
        ...sharedAuth,
      }],
    }],
    async dispose() {
      disposed = true; clearInterval(prune);
      await Promise.all([...pending.values()].map(flow => flow.entry.session.close()));
      pending.clear(); invocations.clear();
      await Promise.all([...sessions.values()].map(entry => entry.session.close()));
      await new Promise(resolve => { server.close(resolve); server.closeIdleConnections(); });
    },
  };
}
export default function setup(ctx) { return createAdapter(ctx); }
