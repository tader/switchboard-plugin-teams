import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { EventEmitter } from 'node:events';
import { Readable } from 'node:stream';
import { createAdapter } from '../index.js';
import { TeamsSession, exchangeChatToken } from '../lib/api-auth.js';
import { TeamsAPI, messagePayload } from '../lib/teams-api.js';
import { requestJSON, chatOrigin, tokenClaims } from '../lib/api-http.js';
import { AdapterError } from '../lib/errors.js';
import { encodeTokenBundle } from '../lib/token-bundle.js';
import { signedBundle, signedToken, proprietaryToken, serverAccepts } from './token-fixture.js';

const thread = '19:test@thread.v2';
const expiry = () => Date.now() + 3600_000;
const bundle = profile => ({ profile, authVersion: 1, identity: { tenant: '11111111-1111-1111-1111-111111111111', oid: profile, name: 'Test user', email: 'test@example.com' },
  tokens: { skype: { value: 'skype-' + profile, expiresAt: expiry() }, chatsvcagg: { value: 'csa-' + profile, expiresAt: expiry() } } });
async function fixture(t, { forbidBrowser = false, dataDir } = {}) {
  const directory = dataDir ?? await fs.mkdtemp(path.join(os.tmpdir(), 'teams-api-test-'));
  const requests = [], sessions = new Map(), authCalls = []; let handler, failSend = false, pageLink = null, authFailure = false, rejectDirectory = false, rejectAuthz = false;
  const response = value => new Response(JSON.stringify(value), { status: 200 });
  const fetcher = async (raw, options) => {
    const url = new URL(raw); requests.push({ url, ...options });
    if (url.pathname.endsWith('/authz')) {
      if (rejectAuthz) return new Response(null, { status: 401 });
      const token = options.headers.Authorization.slice(7);
      if (!token.startsWith('skype-') && !serverAccepts(token)) return new Response(null, { status: 401 });
      const claims = tokenClaims(token);
      const chat = claims.oid ? signedToken('skype', { skypeid: `orgid:${claims.oid}`, tid: claims.tid }) : 'chat-' + token.split('skype-')[1];
      return response({ tokens: { skypeToken: chat, expiresIn: 3600 }, region: 'emea', regionGtms: {} });
    }
    if (url.pathname.endsWith('/teams/users/me') && rejectDirectory) return new Response(null, { status: 401 });
    if (url.pathname.endsWith('/teams/users/me')) {
      const token = options.headers.Authorization.slice(7);
      if (!token.startsWith('csa-') && !serverAccepts(token)) return new Response(null, { status: 401 });
    }
    if (url.pathname.endsWith('/teams/users/me')) return response({ chats: [{ id: thread, title: 'Test chat', isRead: false, isOneOnOne: true, members: [], consumptionHorizon: { originalArrivalTime: Date.parse('2026-10-07T12:00:00Z') } }], teams: [], metadata: { isPartialData: false } });
    if (url.pathname.endsWith('/messages') && (options.method ?? 'GET') === 'GET') {
      const chat = options.headers.Authentication.slice('skypetoken='.length);
      const self = tokenClaims(chat).skypeid?.replace(/^orgid:/, '') ?? chat.split('chat-')[1];
      return response({ messages: [{ id: '1', from: '8:orgid:' + self, content: 'Original', messagetype: 'Text', originalarrivaltime: '2026-10-07T12:01:00Z', version: 'v1', properties: {} }], _metadata: { backwardLink: pageLink } });
    }
    if (url.pathname.endsWith('/messages') && options.method === 'POST') {
      if (failSend) throw new Error('secret server body must not escape');
      return response({ OriginalArrivalTime: 1234 });
    }
    if (/\/messages\/1(?:\/properties)?$/.test(url.pathname)) return new Response(null, { status: 204 });
    throw new Error('Unexpected test request: ' + url.pathname);
  };
  const instance = await createAdapter({ dataDir: directory, settings: {} }, {
    createSession(profile, dir, settings) {
      const session = new TeamsSession(profile, dir, settings, { fetcher, createAuthenticator: () => {
        if (forbidBrowser) throw new Error('A browser must never be created');
        return { async acquire(options) { authCalls.push(options); if (authFailure) throw new AdapterError('signin_required', 'Sign in again.', 503); return bundle(profile); }, async close() {},
      }; } }); sessions.set(profile, session); return session;
    }, createAPI: session => new TeamsAPI(session, { fetcher }),
    createServer(callback) {
      handler = callback; const server = new EventEmitter();
      server.listen = (_, host, done) => { assert.equal(host, '127.0.0.1'); done(); };
      server.address = () => ({ port: 1234 }); server.close = done => done(); server.closeIdleConnections = () => {}; return server;
    },
  });
  t.after(async () => { await instance.dispose(); await fs.rm(directory, { recursive: true, force: true }); });
  const auth = instance.services[0].authMethods[0];
  async function connect() {
    const flow = await auth.connect({ config: { label: 'Test' } });
    await sessions.get(flow.pending.profile).loginTask;
    return auth.poll({ pending: flow.pending });
  }
  async function authorize(connection, pathname, method = 'GET', force = false) {
    const outgoing = { method, url: new URL(pathname, 'http://teams.localhost'), headers: new Headers() };
    const result = await auth.authorize(outgoing, connection, { force });
    if (result?.credentials) connection.credentials = result.credentials;
    return outgoing;
  }
  async function invoke(outgoing, input, override = {}) {
    const request = Readable.from(input === undefined ? [] : [Buffer.from(JSON.stringify(input))]);
    request.headers = Object.fromEntries(outgoing.headers); request.method = override.method ?? outgoing.method; request.url = override.pathname ?? outgoing.url.pathname + outgoing.url.search;
    let result;
    await handler(request, { destroyed: false, writeHead(status) { result = { status }; }, end(body) { result.body = JSON.parse(body); } });
    return result;
  }
  const call = async (connection, pathname, method = 'GET', input) => invoke(await authorize(connection, pathname, method), input);
  return { instance, directory, sessions, authCalls, requests, connect, call, invoke, authorize, auth, setFailSend: value => failSend = value, setLink: value => pageLink = value, setAuthFailure: value => authFailure = value, setRejectDirectory: value => rejectDirectory = value, setRejectAuthz: value => rejectAuthz = value, tokenAuth: instance.services[0].authMethods.find(method => method.id === 'tokens') };
}

test('API plugin reads with Chrome closed, persists renewed credentials and scopes invocation capabilities', async t => {
  const f = await fixture(t), a = await f.connect(), b = await f.connect();
  assert.notEqual(a.credentials.profile, b.credentials.profile);
  const listed = await f.call(a, '/chats'); assert.equal(listed.status, 200);
  const id = listed.body.items[0].id;
  const read = await f.call(a, `/chats/${id}/messages`);
  assert.equal(read.body.items[0].text, 'Original'); assert.equal(read.body.mayMarkRead, false);
  assert.equal(f.authCalls.length, 2, 'data reads never launch a browser with fresh credentials');
  const messageRequest = f.requests.find(request => request.url.pathname.endsWith('/messages'));
  assert.equal(messageRequest.headers.Authentication, 'skypetoken=chat-' + a.credentials.profile);
  assert.equal(messageRequest.redirect, 'error');
  const status = await f.call(a, '/status');
  assert.equal(status.body.transport, 'api'); assert.equal(status.body.browserOpen, false);
  assert.ok(!JSON.stringify(status.body).includes(a.credentials.tokens.skype.value));
  const out = await f.authorize(a, '/status'); assert.equal((await f.invoke(out)).status, 200); assert.equal((await f.invoke(out)).status, 401);
  assert.equal((await f.invoke(await f.authorize(a, '/status'), undefined, { pathname: '/chats' })).status, 403);
  a.credentials.tokens.skype.expiresAt = 1;
  await f.call(a, '/chats'); assert.equal(f.authCalls.length, 3); assert.equal(f.authCalls.at(-1).interactive, false);
  assert.ok(a.credentials.tokens.skype.expiresAt > Date.now());
  const stale = await f.authorize(a, '/status'); await f.auth.revoke(a); assert.equal((await f.invoke(stale)).status, 401);
});

test('API plugin renews only the chat token when possible and refuses silent MFA without hanging', async t => {
  const f = await fixture(t), connection = await f.connect();
  connection.credentials.chatToken.expiresAt = 1;
  await f.call(connection, '/chats'); assert.equal(f.authCalls.length, 1);
  f.setAuthFailure(true); connection.credentials.tokens.skype.expiresAt = 1;
  await assert.rejects(f.call(connection, '/chats'), { code: 'signin_required' });
  const status = await f.call(connection, '/status'); assert.equal(status.body.signInRequired, true);
  f.setAuthFailure(false); await f.call(connection, '/login', 'POST');
  await f.sessions.get(connection.credentials.profile).loginTask;
  assert.equal((await f.call(connection, '/status')).body.ready, true);
});

test('legacy connections expose reconnect guidance and migration gaps without opening a browser', async t => {
  const f = await fixture(t), connection = { credentials: { profile: randomUUID() } };
  assert.equal((await f.call(connection, '/status')).body.signInRequired, true);
  await assert.rejects(f.call(connection, '/chats'), { code: 'signin_required' });
  assert.equal(f.authCalls.length, 0);
  const fresh = await f.connect();
  assert.equal((await f.call(fresh, '/search/messages?query=test')).status, 501);
  assert.ok(!f.instance.services[0].openapi.paths['/search/messages']);
});

test('API writes preserve duplicate protection, quarantine uncertain sends and validate before dispatch', async t => {
  const f = await fixture(t), connection = await f.connect();
  const id = (await f.call(connection, '/chats')).body.items[0].id, endpoint = `/chats/${id}/messages`;
  const input = { text: 'Hello', idempotencyKey: 'send-test-0001' };
  assert.equal((await f.call(connection, endpoint, 'POST', input)).body.status, 'accepted_by_api');
  assert.equal((await f.call(connection, endpoint, 'POST', input)).body.replayed, true);
  assert.equal(f.requests.filter(request => request.method === 'POST' && request.url.pathname.endsWith('/messages')).length, 1);
  assert.equal((await f.call(connection, endpoint, 'POST', { ...input, text: 'Changed' })).body.error.code, 'idempotency_conflict');
  f.setFailSend(true);
  const uncertain = { ...input, idempotencyKey: 'send-test-0002' };
  assert.equal((await f.call(connection, endpoint, 'POST', uncertain)).body.error.code, 'send_uncertain');
  const count = f.requests.length;
  assert.equal((await f.call(connection, endpoint, 'POST', uncertain)).body.error.code, 'send_uncertain'); assert.equal(f.requests.length, count);
  assert.equal((await f.call(connection, endpoint, 'POST', { ...input, replyToMessageId: '1' })).status, 501);
  assert.equal((await f.call(connection, endpoint, 'POST', { ...input, unexpected: true })).status, 400);
  const edit = await f.call(connection, endpoint + '/1', 'PATCH', { text: 'Changed', expectedText: 'Original', idempotencyKey: 'edit-test-0001' });
  assert.equal(edit.body.status, 'accepted_by_api');
  const refused = await f.call(connection, endpoint + '/1', 'DELETE', { expectedText: 'wrong', idempotencyKey: 'delete-test-0001' });
  assert.equal(refused.body.error.code, 'message_changed');
});

test('history cursors replay, remain account-bound and reject credential-exfiltrating backward links', async t => {
  const f = await fixture(t), connection = await f.connect();
  const id = (await f.call(connection, '/chats')).body.items[0].id;
  f.setLink(`https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/${encodeURIComponent(thread)}/messages?cursor=2`);
  const first = await f.call(connection, `/chats/${id}/history?limit=1`), cursor = first.body.nextCursor;
  assert.ok(cursor);
  const second = await f.call(connection, `/chats/${id}/history?limit=1&cursor=${cursor}`);
  assert.equal(second.body.items.length, 0, 'deduplicate overlapping server pages');
  const count = f.requests.length;
  assert.deepEqual(await f.call(connection, `/chats/${id}/history?limit=1&cursor=${cursor}`), second); assert.equal(f.requests.length, count);
  const b = await f.connect(); assert.equal((await f.call(b, `/chats/${id}/history?limit=1&cursor=${cursor}`)).status, 400);
  f.setLink('https://evil.example/steal');
  const malicious = await f.call(connection, `/chats/${id}/history?limit=1`);
  const follow = await f.call(connection, `/chats/${id}/history?limit=1&cursor=${malicious.body.nextCursor}`);
  assert.equal(follow.body.error.code, 'api_cursor_rejected');
  assert.ok(f.requests.every(request => request.url.hostname !== 'evil.example'));
});

test('transport redacts remote errors, rejects untrusted regional hosts, and formats content safely', async () => {
  await assert.rejects(requestJSON('https://teams.microsoft.com/test', {}, async () => new Response('token=secret', { status: 403 })), error => error.code === 'api_forbidden' && !error.message.includes('secret'));
  assert.throws(() => chatOrigin({ region: 'emea', regionGtms: { chatService: 'https://evil.example' } }), { code: 'api_region_invalid' });
  const payload = messagePayload({ content: [{ type: 'paragraph', runs: [{ text: '<script>', marks: ['bold'] }] }] }, 'Test');
  assert.equal(payload.content, '<p><strong>&lt;script&gt;</strong></p>');
  assert.throws(() => messagePayload({ content: [{ type: 'paragraph', runs: [{ mention: { email: 'person@example.com', name: 'Person' } }] }] }, 'Test'), { code: 'api_operation_unsupported' });
  let tokenHeader;
  await exchangeChatToken(bundle('profile'), async (_url, options) => { tokenHeader = options.headers; return new Response(JSON.stringify({ tokens: { skypeToken: 'chat', expiresIn: 3600 }, region: 'emea' })); });
  assert.equal(tokenHeader.Authorization, 'Bearer skype-profile');
});

test('imported connections read and renew chat tokens without ever creating a browser', async t => {
  const f = await fixture(t, { forbidBrowser: true });
  const connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle()) } });
  assert.deepEqual(connection.config, {});
  assert.equal(f.tokenAuth.fields.find(field => field.key === 'tokenBundle').type, 'secret');
  const listed = await f.call(connection, '/chats'); assert.equal(listed.status, 200);
  const read = await f.call(connection, `/chats/${listed.body.items[0].id}/messages`); assert.equal(read.body.items[0].text, 'Original');
  let status = (await f.call(connection, '/status')).body;
  assert.equal(status.authMode, 'tokens'); assert.equal(status.tokenImportRequired, false); assert.equal(status.browserOpen, false);
  assert.ok(!JSON.stringify(status).includes(connection.credentials.tokens.skype.value));
  const exchanges = f.requests.filter(request => request.url.pathname.endsWith('/authz')).length;
  connection.credentials.chatToken.expiresAt = 1;
  await f.call(connection, '/chats');
  assert.equal(f.requests.filter(request => request.url.pathname.endsWith('/authz')).length, exchanges + 1);
  assert.ok(connection.credentials.chatToken.expiresAt > Date.now());
  const login = await f.call(connection, '/login', 'POST'); assert.equal(login.body.error.code, 'token_import_required');
  assert.equal((await f.call(connection, '/status')).body.ready, true, 'requesting login does not invalidate fresh imported tokens');
  connection.credentials.tokens.skype.expiresAt = 1;
  status = (await f.call(connection, '/status')).body;
  assert.equal(status.tokenImportRequired, true); assert.equal(status.signInRequired, true); assert.equal(status.ready, false);
  await assert.rejects(f.call(connection, '/chats'), { code: 'token_import_required' });
  assert.equal(f.authCalls.length, 0);
  assert.deepEqual(await fs.readdir(f.directory), [], 'no browser profile is created');
});

test('re-import replaces cached credentials, preserves the ledger and invalidates old invocations', async t => {
  const f = await fixture(t, { forbidBrowser: true });
  let connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle()) } });
  const id = (await f.call(connection, '/chats')).body.items[0].id, input = { text: 'Fixture only', idempotencyKey: 'portable-send' };
  const sent = await f.call(connection, `/chats/${id}/messages`, 'POST', input); assert.equal(sent.status, 200);
  const old = f.sessions.get(connection.credentials.profile), outstanding = await f.authorize(connection, '/status');
  await assert.rejects(f.tokenAuth.connect({ config: { label: 'Bad', tokenBundle: 'not json' }, connection }), { code: 'invalid_token_bundle' });
  assert.equal(old.disposed, false); assert.equal((await f.call(connection, '/status')).body.ready, true);
  const requestsBefore = f.requests.length;
  await assert.rejects(f.tokenAuth.connect({ config: { label: 'Other', tokenBundle: encodeTokenBundle(signedBundle({ oid: '33333333-3333-3333-3333-333333333333' })) }, connection }), { code: 'account_mismatch' });
  assert.equal(f.requests.length, requestsBefore, 'account mismatches are rejected before any API calls');
  f.setRejectDirectory(true);
  await assert.rejects(f.tokenAuth.connect({ config: { label: 'Failed', tokenBundle: encodeTokenBundle(signedBundle()) }, connection }), { code: 'api_unauthorized' });
  f.setRejectDirectory(false); assert.equal((await f.call(connection, '/status')).body.ready, true);
  const profile = connection.credentials.profile; connection.credentials.tokens.skype.expiresAt = 1;
  connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle({ exp: Math.floor(Date.now() / 1000) + 7200 })) }, connection });
  assert.equal(connection.credentials.profile, profile); assert.equal(old.disposed, true);
  assert.equal((await f.invoke(outstanding)).status, 401);
  assert.equal((await f.call(connection, '/status')).body.ready, true);
  assert.deepEqual((await f.call(connection, `/chats/${id}/messages`, 'POST', input)).body, { ...sent.body, replayed: true });
  assert.equal(f.requests.filter(request => request.method === 'POST' && request.url.pathname.endsWith('/messages')).length, 1);
  assert.equal(f.authCalls.length, 0);
});

test('persisted imported credentials recover after reload and retain duplicate protection', async t => {
  const first = await fixture(t, { forbidBrowser: true });
  const connection = await first.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle()) } });
  const id = (await first.call(connection, '/chats')).body.items[0].id, input = { text: 'Fixture only', idempotencyKey: 'reload-send' };
  const sent = await first.call(connection, `/chats/${id}/messages`, 'POST', input);
  const saved = JSON.parse(JSON.stringify(connection)); await first.instance.dispose();
  const restarted = await fixture(t, { forbidBrowser: true, dataDir: first.directory });
  assert.equal((await restarted.call(saved, '/status')).body.authMode, 'tokens');
  assert.equal((await restarted.call(saved, '/chats')).status, 200);
  assert.deepEqual((await restarted.call(saved, `/chats/${id}/messages`, 'POST', input)).body, { ...sent.body, replayed: true });
  assert.equal(restarted.requests.filter(request => request.method === 'POST' && request.url.pathname.endsWith('/messages')).length, 0);
  assert.equal(restarted.authCalls.length, 0);
});

test('rejected access tokens, forced renewal and rejected exchanges require import without browser fallback', async t => {
  const f = await fixture(t, { forbidBrowser: true });
  let connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle()) } });
  f.setRejectDirectory(true);
  const rejected = await f.call(connection, '/chats'); assert.equal(rejected.status, 503); assert.equal(rejected.body.error.code, 'token_import_required');
  assert.equal((await f.call(connection, '/status')).body.tokenImportRequired, true);
  f.setRejectDirectory(false); await assert.rejects(f.call(connection, '/chats'), { code: 'token_import_required' });
  connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle()) }, connection });
  await assert.rejects(f.authorize(connection, '/chats', 'GET', true), { code: 'token_import_required' });
  connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle()) }, connection });
  connection.credentials.chatToken.expiresAt = 1; f.setRejectAuthz(true);
  await assert.rejects(f.call(connection, '/chats'), { code: 'token_import_required' });
  assert.equal((await f.call(connection, '/status')).body.tokenImportRequired, true);
  assert.equal(f.authCalls.length, 0);
  await f.tokenAuth.revoke(connection);
  assert.equal(await fs.stat(path.join(f.directory, 'profiles', connection.credentials.profile)).then(() => true, () => false), false);
});

test('browser connections can reconnect and switch authentication methods without losing their profile', async t => {
  const f = await fixture(t);
  let connection = await f.connect();
  const profile = connection.credentials.profile;
  const original = f.sessions.get(profile);
  connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(signedBundle({ oid: profile })) }, connection });
  assert.equal(connection.credentials.profile, profile); assert.equal(original.disposed, true);
  assert.equal((await f.call(connection, '/status')).body.authMode, 'tokens');
  const imported = f.sessions.get(profile);
  const flow = await f.auth.connect({ config: { label: 'Local' }, connection });
  await f.sessions.get(profile).loginTask;
  connection = await f.auth.poll({ pending: flow.pending });
  assert.equal(connection.credentials.profile, profile); assert.equal(imported.disposed, true);
  assert.equal((await f.call(connection, '/status')).body.authMode, 'browser');
  f.setAuthFailure(true);
  const failed = await f.auth.connect({ config: { label: 'Local' }, connection });
  await f.sessions.get(profile).loginTask;
  await assert.rejects(f.auth.poll({ pending: failed.pending }), { code: 'signin_required' });
  assert.equal((await f.call(connection, '/status')).body.ready, true);
  f.setAuthFailure(false);
  const retried = await f.auth.connect({ config: { label: 'Local' }, connection });
  await f.sessions.get(profile).loginTask;
  connection = await f.auth.poll({ pending: retried.pending });
  assert.equal((await f.call(connection, '/chats')).status, 200);
  const cancelled = await f.auth.connect({ config: { label: 'Local' }, connection });
  await f.sessions.get(profile).loginTask;
  await assert.rejects(f.call(connection, '/chats'), { code: 'signin_pending' });
  const replacement = await f.auth.connect({ config: { label: 'Local' }, connection });
  await f.sessions.get(profile).loginTask;
  await assert.rejects(f.auth.poll({ pending: cancelled.pending }), { code: 'login_expired' });
  connection = await f.auth.poll({ pending: replacement.pending });
  assert.equal((await f.call(connection, '/chats')).status, 200);
});

test('Microsoft proprietary access-token signatures are validated by Teams before import succeeds', async t => {
  const f = await fixture(t, { forbidBrowser: true });
  const bundle = signedBundle();
  for (const service of ['skype', 'chatsvcagg']) bundle.tokens[service].value = proprietaryToken(service);
  const connection = await f.tokenAuth.connect({ config: { label: 'Central', tokenBundle: encodeTokenBundle(bundle) } });
  assert.equal((await f.call(connection, '/chats')).status, 200);
  assert.ok(!f.requests.some(request => request.url.pathname.endsWith('/discovery/keys')), 'clients do not verify Microsoft access-token signatures');
  for (const service of ['skype', 'chatsvcagg']) {
    const forged = structuredClone(bundle); forged.tokens[service].value += '-tampered';
    await assert.rejects(f.tokenAuth.connect({ config: { label: 'Bad', tokenBundle: encodeTokenBundle(forged) }, connection }), { code: 'api_unauthorized' });
    assert.equal((await f.call(connection, '/status')).body.ready, true);
  }
  assert.equal(f.authCalls.length, 0);
});

test('the identity returned by the trusted Teams auth service must match the imported account', async () => {
  const credentials = { ...signedBundle(), authMode: 'tokens' };
  const fetcher = async () => new Response(JSON.stringify({ tokens: { skypeToken: signedToken('skype', { skypeid: 'orgid:33333333-3333-3333-3333-333333333333' }), expiresIn: 3600 }, region: 'emea' }));
  await assert.rejects(exchangeChatToken(credentials, fetcher), { code: 'account_mismatch' });
});
