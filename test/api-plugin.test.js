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
import { requestJSON, chatOrigin } from '../lib/api-http.js';
import { AdapterError } from '../lib/errors.js';

const thread = '19:test@thread.v2';
const expiry = () => Date.now() + 3600_000;
const bundle = profile => ({ profile, authVersion: 1, identity: { tenant: '11111111-1111-1111-1111-111111111111', oid: profile, name: 'Test user', email: 'test@example.com' },
  tokens: { skype: { value: 'skype-' + profile, expiresAt: expiry() }, chatsvcagg: { value: 'csa-' + profile, expiresAt: expiry() } } });
async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-api-test-'));
  const requests = [], sessions = new Map(), authCalls = []; let handler, failSend = false, pageLink = null, authFailure = false;
  const response = value => new Response(JSON.stringify(value), { status: 200 });
  const fetcher = async (raw, options) => {
    const url = new URL(raw); requests.push({ url, ...options });
    if (url.pathname.endsWith('/authz')) return response({ tokens: { skypeToken: 'chat-' + options.headers.Authorization.split('skype-')[1], expiresIn: 3600 }, region: 'emea', regionGtms: {} });
    if (url.pathname.endsWith('/teams/users/me')) return response({ chats: [{ id: thread, title: 'Test chat', isRead: false, isOneOnOne: true, members: [], consumptionHorizon: { originalArrivalTime: Date.parse('2026-10-07T12:00:00Z') } }], teams: [], metadata: { isPartialData: false } });
    if (url.pathname.endsWith('/messages') && (options.method ?? 'GET') === 'GET') {
      const self = options.headers.Authentication.split('chat-')[1];
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
      const session = new TeamsSession(profile, dir, settings, { fetcher, createAuthenticator: () => ({
        async acquire(options) { authCalls.push(options); if (authFailure) throw new AdapterError('signin_required', 'Sign in again.', 503); return bundle(profile); }, async close() {},
      }) }); sessions.set(profile, session); return session;
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
  return { instance, directory, sessions, authCalls, requests, connect, call, invoke, authorize, auth, setFailSend: value => failSend = value, setLink: value => pageLink = value, setAuthFailure: value => authFailure = value };
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
