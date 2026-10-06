import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { EventEmitter } from 'node:events';
import { Readable } from 'node:stream';
import { createAdapter } from '../index.js';

async function fixture(t) {
  const dataDir = await fs.mkdtemp('/tmp/teams-adapter-');
  const profiles = new Map();
  let handler;
  const instance = await createAdapter({ dataDir, settings: {} }, {
    createBrowser: dir => {
      const browser = {
        serial: fn => fn(), login: async () => ({ opened: true }), dom: async action => ({ ready: true, action }),
        listChats: async unreadOnly => ({ profile: dir, unreadOnly }),
        unreadMessages: async (maxChats, limitPerChat) => ({ maxChats, limitPerChat, mayMarkRead: true }),
        searchPeople: async query => ({ query, items: [{ email: 'alice@example.com' }] }),
        conversation: async email => ({ email, messageSent: false }),
        prepareRecipient: async email => { browser.preparedCount = (browser.preparedCount ?? 0) + 1; return { chat: { title: 'Alice', key: { attribute: 'recipient', value: email } } }; },
        send: async () => { browser.sentCount = (browser.sentCount ?? 0) + 1; return { status: 'observed_in_chat', message: { id: 'sent1' } }; },
        chatId: () => 'person-chat-id', close: async () => {},
      };
      profiles.set(dir, browser); return browser;
    },
    createServer: callback => {
      handler = callback;
      const server = new EventEmitter();
      server.listen = (_port, host, ready) => { assert.equal(host, '127.0.0.1'); ready(); };
      server.address = () => ({ port: 12345 });
      server.close = done => done();
      server.closeIdleConnections = () => {};
      return server;
    },
  });
  t.after(async () => { await instance.dispose(); await fs.rm(dataDir, { recursive: true, force: true }); });
  const service = instance.services[0], auth = service.authMethods[0];
  async function connection() {
    const flow = await auth.connect({ config: { label: 'test' } });
    return auth.poll({ pending: flow.pending });
  }
  function request(conn, pathname, method = 'GET') {
    const outgoing = { url: new URL(pathname, service.baseUrl), method, headers: new Headers() };
    auth.authorize(outgoing, conn); return outgoing;
  }
  async function invoke(out, { pathname, method, body = '', destroyed = false } = {}) {
    const request = Readable.from(body ? [Buffer.from(body)] : []);
    request.headers = Object.fromEntries(out.headers);
    request.url = pathname ?? out.url.pathname + out.url.search;
    request.method = method ?? out.method;
    const result = {};
    const response = { destroyed, writeHead: status => result.status = status, end: text => result.body = JSON.parse(text) };
    await handler(request, response); return result;
  }
  return { service, auth, connection, request, invoke, profiles };
}

test('single-use invocation tokens bind account, method, path and query', async t => {
  const f = await fixture(t), conn = await f.connection();
  const out = f.request(conn, '/status');
  assert.equal((await f.invoke(out)).status, 200);
  assert.equal((await f.invoke(out)).status, 401);
  assert.equal((await f.invoke({ ...out, headers: new Headers() })).status, 401);
  assert.equal((await f.invoke(f.request(conn, '/status'), { pathname: '/chats' })).status, 403);
  assert.equal((await f.invoke(f.request(conn, '/status?a=1'), { pathname: '/status?a=2' })).status, 403);
  assert.equal((await f.invoke(f.request(conn, '/status'), { method: 'POST' })).status, 403);
});

test('connections have separate profiles and malformed requests fail before browser', async t => {
  const f = await fixture(t), a = await f.connection(), b = await f.connection();
  assert.notEqual(a.credentials.profile, b.credentials.profile);
  assert.equal(f.profiles.size, 2);
  assert.equal(f.service.baseUrl, 'http://teams.localhost');
  const result = await f.invoke(f.request(a, '/chats/invalid/messages?limit=-1'));
  assert.equal(result.status, 400);
  assert.equal(result.body.error.code, 'invalid_query');
  const wrongBody = await f.invoke(f.request(a, '/chats/invalid/messages', 'POST'), { body: '{"text":"hello","extra":true}' });
  assert.equal(wrongBody.status, 400);
  assert.equal(wrongBody.body.error.code, 'invalid_body');
  assert.equal((await f.invoke(f.request(a, '/status'), { destroyed: true })).status, 499);
});

test('revocation invalidates pending invocations', async t => {
  const f = await fixture(t), conn = await f.connection();
  const out = f.request(conn, '/status');
  await f.auth.revoke(conn);
  assert.equal((await f.invoke(out)).status, 401);
});

test('unread and people operations validate bounds and route the requested filters', async t => {
  const f = await fixture(t), conn = await f.connection();
  assert.equal((await f.invoke(f.request(conn, '/unread/chats'))).body.unreadOnly, true);
  assert.equal((await f.invoke(f.request(conn, '/chats?unreadOnly=false'))).body.unreadOnly, false);
  assert.equal((await f.invoke(f.request(conn, '/chats?unreadOnly=maybe'))).status, 400);
  assert.deepEqual((await f.invoke(f.request(conn, '/unread/messages?maxChats=2&limitPerChat=10'))).body, { maxChats: 2, limitPerChat: 10, mayMarkRead: true });
  assert.equal((await f.invoke(f.request(conn, '/unread/messages?maxChats=6'))).status, 400);
  assert.equal((await f.invoke(f.request(conn, '/people'))).status, 400);
  assert.equal((await f.invoke(f.request(conn, '/people?query=Alice'))).body.query, 'Alice');
});

test('opening a conversation does not send; initial messages have durable recipient-bound retries', async t => {
  const f = await fixture(t), conn = await f.connection();
  const browser = [...f.profiles.values()][0];
  const post = input => f.invoke(f.request(conn, '/conversations', 'POST'), { body: JSON.stringify(input) });
  assert.deepEqual((await post({ email: 'Alice@Example.COM' })).body, { email: 'alice@example.com', messageSent: false });
  assert.equal(browser.sentCount, undefined);
  assert.equal((await post({ email: 'Alice Adams', text: 'Hi', idempotencyKey: 'person-send-1' })).status, 400);
  assert.equal((await post({ email: 'alice@example.com', text: 'Hi' })).status, 400);
  const first = await post({ email: 'alice@example.com', text: 'Hi', idempotencyKey: 'person-send-1' });
  const retry = await post({ email: 'ALICE@EXAMPLE.COM', text: 'Hi', idempotencyKey: 'person-send-1' });
  assert.equal(first.body.chatId, 'person-chat-id');
  assert.equal(retry.body.replayed, true);
  assert.equal(browser.sentCount, 1);
  assert.equal(browser.preparedCount, 1);
  assert.equal((await post({ email: 'bob@example.com', text: 'Hi', idempotencyKey: 'person-send-1' })).body.error.code, 'idempotency_conflict');
});
