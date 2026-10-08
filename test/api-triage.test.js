import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { TeamsAPI } from '../lib/teams-api.js';
import { TeamsTriage, triageProtocols } from '../lib/triage.js';
import { SendLedger } from '../lib/ledger.js';
import { requestJSON } from '../lib/api-http.js';
import { openapi } from '../lib/api-openapi.js';

const fixtureData = JSON.parse(await fs.readFile(new URL('./fixtures/triage.json', import.meta.url), 'utf8'));
const enabled = { ...triageProtocols, channelReplies: true, readMarkers: true };
const channel = '19:channel@thread.tacv2', root = '1791374403000';
const rawMessage = (time, content = 'New arrival', from = '8:orgid:peer', parentMessageId) => ({ id: String(time), originalarrivaltime: time,
  from, imdisplayname: 'Peer', content, messagetype: 'Text', version: '1', clientmessageid: String(time), ...(parentMessageId ? { parentMessageId } : {}) });
async function fixture(t, { protocols = triageProtocols } = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-triage-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const data = structuredClone(fixtureData), requests = [], links = new Map();
  const state = { sends: 0, marks: 0, failSend: false, failRead: false, afterSend: null, rateLimit: null };
  const credentials = { identity: { tenant: 'tenant', oid: 'self', email: 'self@example.com', name: 'Self' },
    messageOrigin: 'https://emea.ng.msg.teams.microsoft.com', tokens: { chatsvcagg: { value: 'fixture-csa' } }, chatToken: { value: 'fixture-chat' } };
  const fetcher = async (value, options) => {
    const url = new URL(value); requests.push({ url, ...options });
    const response = value => new Response(JSON.stringify(value), { status: 200 });
    if (url.pathname.endsWith('/teams/users/me')) return response({ chats: data.chats, teams: data.teams, metadata: data.metadata ?? {} });
    const match = url.pathname.match(/\/conversations\/([^/]+)\/(messages|properties)$/);
    assert.ok(match, 'expected regional request'); const thread = decodeURIComponent(match[1]);
    if (match[2] === 'messages' && !options.method) {
      if (state.rateLimit === thread) return new Response(null, { status: 429 });
      return response({ messages: data.messages[thread] ?? [], _metadata: { backwardLink: links.get(thread) } });
    }
    if (options.method === 'POST') {
      state.sends++; if (state.failSend) throw new Error('connection lost');
      const payload = JSON.parse(options.body), time = fixtureData.boundary + 9000 + state.sends;
      data.messages[thread].push({ ...rawMessage(time, payload.content, '8:orgid:self'), messagetype: payload.messagetype, clientmessageid: payload.clientmessageid });
      state.afterSend?.(thread); return response({ OriginalArrivalTime: time });
    }
    if (options.method === 'PUT') {
      state.marks++; if (state.failRead) return new Response(null, { status: 503 });
      const marker = JSON.parse(options.body).consumptionhorizon;
      const conversation = [...data.chats, ...data.teams.flatMap(team => team.channels)].find(item => item.id === thread);
      conversation.userConsumptionHorizon = { originalArrivalTime: Number(marker.split(';')[0]) };
      return new Response(null, { status: 204 });
    }
    assert.fail('unexpected request');
  };
  const api = new TeamsAPI({ credentials }, { fetcher }), ledger = new SendLedger(directory), triage = new TeamsTriage(api, ledger, { protocols });
  return { api, ledger, triage, data, state, requests, links, protocols };
}
const reply = (group, key = 'triage-reply-001', text = 'Thanks!') => ({ target: group.replyTarget, text, idempotencyKey: key });

test('grouped inbox prioritizes mentions/direct chats, groups channel replies and never marks read', async t => {
  const f = await fixture(t), result = await f.triage.inbox();
  assert.deepEqual(result.items.map(item => item.title), ['Mention chat', 'Direct chat', 'Group chat', 'Test channel']);
  assert.equal(result.items[0].signals.mentionedMe, true);
  assert.equal(result.items[1].unread[0].authorId, '8:orgid:peer');
  assert.equal(result.items[3].parentMessageId, root);
  assert.equal(result.items[3].unread.length, 2);
  assert.equal(result.items[3].canReply, false);
  assert.equal(result.items[0].context.length, 0, 'unread messages are not duplicated in context');
  assert.equal(result.completeAccount, false); assert.equal(result.mayMarkRead, false);
  assert.equal(f.state.marks, 0); assert.equal(f.state.sends, 0);
});

test('own messages, deleted messages and system events are excluded from unread; repliedSince remains factual', async t => {
  const f = await fixture(t), messages = f.data.messages[f.data.chats[0].id];
  messages.push(rawMessage(fixtureData.boundary + 5000, 'My answer', '8:orgid:self'));
  messages.push({ ...rawMessage(fixtureData.boundary + 6000), properties: { deletetime: 'yes' } });
  messages.push({ ...rawMessage(fixtureData.boundary + 7000), messagetype: 'ThreadActivity/AddMember' });
  const result = await f.triage.inbox();
  assert.equal(result.items[0].unread.length, 1); assert.equal(result.items[0].unread[0].repliedSince, true);
  assert.equal(result.items[0].context[0].isOwn, true);
});

test('missing horizons and unknown unread flags are coverage issues, not guessed unread messages', async t => {
  const f = await fixture(t); delete f.data.chats[0].consumptionHorizon; delete f.data.teams[0].channels[0].isMessageRead;
  f.data.metadata = { isPartialData: true };
  const result = await f.triage.inbox();
  assert.equal(result.items.length, 2); assert.equal(result.coverage.discoveryPartial, true);
  assert.deepEqual(result.coverage.issues.map(item => item.reason).sort(), ['unknown_consumption_horizon', 'unknown_unread_flag']);
});

test('inbox cursor replay is stable, parameter/account-bound and expires; context refresh is read-only', async t => {
  const f = await fixture(t), first = await f.triage.inbox({ limit: 1 });
  const second = await f.triage.inbox({ limit: 1, cursor: first.nextCursor });
  assert.equal(second.items[0].title, 'Direct chat');
  assert.deepEqual(await f.triage.inbox({ limit: 1, cursor: first.nextCursor }), second);
  await assert.rejects(f.triage.inbox({ limit: 2, cursor: first.nextCursor }), { code: 'invalid_cursor' });
  const contexts = await f.triage.contexts({ targets: [first.items[0].replyTarget, 'invalid-target'] });
  assert.notEqual(contexts.items[0].replyTarget, first.items[0].replyTarget);
  assert.equal(contexts.items[1].error.code, 'invalid_reply_target'); assert.equal(f.state.marks, 0);
  f.api.session.credentials.identity.oid = 'other-account';
  await assert.rejects(f.triage.inbox({ limit: 1, cursor: first.nextCursor }), { code: 'invalid_cursor' });
  assert.throws(() => f.triage.target(first.items[0].replyTarget), { code: 'invalid_reply_target' });
});

test('quoted chat reply escapes incoming content and replies; unverified read updates defer explicitly', async t => {
  const f = await fixture(t), group = (await f.triage.inbox()).items[0];
  const result = await f.triage.replies({ replies: [reply(group, 'quote-reply-001', '<safe> & clear')] });
  assert.equal(result.items[0].send.status, 'accepted_by_api');
  assert.deepEqual(result.items[0].read, { status: 'deferred', reason: 'protocol_unverified' });
  const post = f.requests.find(request => request.method === 'POST'), payload = JSON.parse(post.body);
  assert.equal(payload.messagetype, 'RichText/Html');
  assert.ok(payload.content.includes(`itemtype="http://schema.skype.com/Reply" itemid="${group.unread[0].id}"`));
  assert.ok(payload.content.includes('Peer &amp; friend')); assert.ok(payload.content.includes('Please review &lt;this&gt;'));
  assert.ok(payload.content.endsWith('<p>&lt;safe&gt; &amp; clear</p>')); assert.equal(f.state.marks, 0);
});

test('context changes and target edits prevent sends; a failed item does not block unrelated valid replies', async t => {
  const f = await fixture(t), groups = (await f.triage.inbox()).items;
  f.data.messages[f.data.chats[0].id][0].version = '2';
  const result = await f.triage.replies({ replies: [reply(groups[0], 'stale-reply-001'), reply(groups[1], 'valid-reply-001')] });
  assert.equal(result.items[0].send.error.code, 'context_changed');
  assert.equal(result.items[1].send.status, 'accepted_by_api'); assert.equal(f.state.sends, 1);
  const fresh = (await f.triage.contexts({ targets: [groups[0].replyTarget] })).items[0];
  f.data.messages[f.data.chats[0].id].push(rawMessage(fixtureData.boundary + 11000));
  assert.equal((await f.triage.replies({ replies: [reply(fresh, 'new-arrival-001')] })).items[0].send.error.code, 'context_changed');
});

test('malformed batch content and duplicate keys fail before any write', async t => {
  const f = await fixture(t), group = (await f.triage.inbox()).items[0];
  await assert.rejects(f.triage.replies({ replies: [reply(group), reply(group, 'invalid-reply-002', '')] }), { code: 'invalid_message' });
  await assert.rejects(f.triage.replies({ replies: [reply(group), reply(group)] }), { code: 'invalid_body' });
  assert.equal(f.state.sends, 0);
});

test('uncertain sends stop the batch and remain quarantined across restart', async t => {
  const f = await fixture(t), groups = (await f.triage.inbox()).items; f.state.failSend = true;
  const input = { replies: [reply(groups[0]), reply(groups[1], 'second-reply-002')] };
  const result = await f.triage.replies(input);
  assert.equal(result.items[0].send.status, 'uncertain'); assert.equal(result.items[1].send.status, 'not_attempted');
  const restarted = new TeamsTriage(f.api, new SendLedger(path.dirname(f.ledger.file)));
  assert.equal((await restarted.replies(input)).items[0].send.status, 'uncertain'); assert.equal(f.state.sends, 1);
});

test('accepted reply replays after restart and target expiry; changed content conflicts', async t => {
  const f = await fixture(t), group = (await f.triage.inbox()).items[0], input = { replies: [reply(group)] };
  await f.triage.replies(input);
  const restarted = new TeamsTriage(f.api, new SendLedger(path.dirname(f.ledger.file)));
  const result = await restarted.replies(input);
  assert.equal(result.items[0].send.replayed, true); assert.equal(f.state.sends, 1);
  assert.equal((await restarted.replies({ replies: [{ ...reply(group), text: 'different' }] })).items[0].send.error.code, 'idempotency_conflict');
  assert.ok(!JSON.stringify(result).includes('_triageRead'));
});

test('read stage failure never hides accepted reply; replay retries read without sending', async t => {
  const f = await fixture(t, { protocols: enabled }), group = (await f.triage.inbox()).items[0]; f.state.failRead = true;
  const input = { replies: [reply(group)] }, result = await f.triage.replies(input);
  assert.equal(result.items[0].send.status, 'accepted_by_api'); assert.equal(result.items[0].read.status, 'failed');
  f.state.failRead = false;
  const restarted = new TeamsTriage(f.api, new SendLedger(path.dirname(f.ledger.file)), { protocols: enabled });
  const retry = await restarted.replies(input);
  assert.equal(retry.items[0].send.replayed, true); assert.equal(retry.items[0].read.status, 'marked_read');
  assert.equal(f.state.sends, 1); assert.equal(f.state.marks, 2);
});

test('read updates stop at observed content and preserve newer existing horizons', async t => {
  const f = await fixture(t, { protocols: enabled }), group = (await f.triage.inbox()).items[0];
  f.state.afterSend = thread => f.data.messages[thread].push(rawMessage(fixtureData.boundary + 20000));
  const result = await f.triage.replies({ replies: [reply(group)] });
  assert.equal(result.items[0].read.throughMessageId, group.unread.at(-1).id);
  assert.equal(f.data.chats[0].userConsumptionHorizon.originalArrivalTime, fixtureData.boundary + 2000);
  const other = await fixture(t, { protocols: enabled }), otherGroup = (await other.triage.inbox()).items[0];
  other.data.chats[0].userConsumptionHorizon = { originalArrivalTime: fixtureData.boundary + 40000 };
  assert.equal((await other.triage.replies({ replies: [reply(otherGroup)] })).items[0].read.status, 'already_read');
  assert.equal(other.state.marks, 0);
});

test('unhandled sibling channel thread defers marking; later reply makes advancement safe', async t => {
  const f = await fixture(t, { protocols: enabled }), sibling = rawMessage(fixtureData.boundary + 3500, 'Thread B');
  f.data.messages[channel].push(sibling); f.data.messages[`${channel};messageid=${sibling.id}`] = [];
  const groups = (await f.triage.inbox()).items.filter(group => group.kind === 'channel');
  const a = groups.find(group => group.parentMessageId === root), b = groups.find(group => group.parentMessageId === sibling.id);
  const first = await f.triage.replies({ replies: [reply(a, 'channel-reply-001')] });
  assert.equal(first.items[0].send.status, 'accepted_by_api'); assert.equal(first.items[0].read.status, 'deferred');
  assert.equal(f.state.marks, 0);
  const second = await f.triage.replies({ replies: [reply(b, 'channel-reply-002')] });
  assert.equal(second.items[0].read.status, 'marked_read'); assert.equal(f.state.marks, 1);
  assert.equal(f.data.teams[0].channels[0].userConsumptionHorizon.originalArrivalTime, fixtureData.boundary + 4000);
  const posts = f.requests.filter(request => request.method === 'POST');
  assert.ok(decodeURIComponent(posts[0].url.pathname).includes(`;messageid=${root}/messages`));
});

test('partial windows never mark read; production channel replies fail without dispatch', async t => {
  const f = await fixture(t, { protocols: enabled });
  f.links.set(f.data.chats[0].id, `${f.api.credentials.messageOrigin}/v1/users/ME/conversations/${encodeURIComponent(f.data.chats[0].id)}/messages?older=1`);
  const result = await f.triage.replies({ replies: [reply((await f.triage.inbox()).items[0])] });
  assert.equal(result.items[0].read.reason, 'incomplete_unread_window'); assert.equal(f.state.marks, 0);
  const production = await fixture(t), group = (await production.triage.inbox()).items.find(group => group.kind === 'channel');
  assert.equal((await production.triage.replies({ replies: [reply(group)] })).items[0].send.error.code, 'api_operation_unsupported');
  assert.equal(production.state.sends, 0);
});

test('own writes in a batch do not invalidate its subsequent replies', async t => {
  const f = await fixture(t), group = (await f.triage.inbox()).items[0];
  const result = await f.triage.replies({ replies: [reply(group, 'same-chat-001', 'First'), reply(group, 'same-chat-002', 'Second')] });
  assert.deepEqual(result.items.map(item => item.send.status), ['accepted_by_api', 'accepted_by_api']); assert.equal(f.state.sends, 2);
});

test('an old channel root cannot prove that a paginated reply window covers the horizon', async t => {
  const f = await fixture(t, { protocols: enabled }), oldRoot = rawMessage(fixtureData.boundary - 1000, 'Old root');
  f.data.messages[channel] = [oldRoot];
  const thread = `${channel};messageid=${oldRoot.id}`;
  f.data.messages[thread] = [rawMessage(fixtureData.boundary + 4000, 'Recent reply', '8:orgid:peer', oldRoot.id)];
  f.links.set(thread, `${f.api.credentials.messageOrigin}/v1/users/ME/conversations/${encodeURIComponent(thread)}/messages?older=1`);
  const group = (await f.triage.inbox()).items.find(group => group.kind === 'channel');
  assert.equal(group.coverage.unreadWindowComplete, false);
  const result = await f.triage.replies({ replies: [reply(group)] });
  assert.equal(result.items[0].read.reason, 'incomplete_unread_window'); assert.equal(f.state.marks, 0);
});

test('ledger failure after send leaves an uncertain intent; read-stage persistence failure preserves accepted result', async t => {
  const f = await fixture(t), group = (await f.triage.inbox()).items[0], write = f.ledger.write.bind(f.ledger); let writes = 0;
  f.ledger.write = records => ++writes === 2 ? Promise.reject(new Error('disk full')) : write(records);
  const input = { replies: [reply(group)] }, result = await f.triage.replies(input);
  assert.equal(result.items[0].send.status, 'uncertain'); assert.equal(f.state.sends, 1);
  const restarted = new TeamsTriage(f.api, new SendLedger(path.dirname(f.ledger.file)));
  assert.equal((await restarted.replies(input)).items[0].send.status, 'uncertain'); assert.equal(f.state.sends, 1);
  const other = await fixture(t, { protocols: enabled }), otherGroup = (await other.triage.inbox()).items[0], otherWrite = other.ledger.write.bind(other.ledger); let otherWrites = 0;
  other.ledger.write = records => ++otherWrites === 3 ? Promise.reject(new Error('disk full')) : otherWrite(records);
  const otherInput = { replies: [reply(otherGroup)] }, accepted = await other.triage.replies(otherInput);
  assert.equal(accepted.items[0].send.status, 'accepted_by_api'); assert.equal(accepted.items[0].read.reason, 'read_update_persistence_failed');
  const recovery = new TeamsTriage(other.api, new SendLedger(path.dirname(other.ledger.file)), { protocols: enabled });
  const replay = await recovery.replies(otherInput);
  assert.equal(replay.items[0].send.replayed, true); assert.equal(replay.items[0].read.status, 'already_read');
  assert.equal(other.state.sends, 1); assert.equal(other.state.marks, 1);
});

test('bounded source continuation can return an empty page and still reach later unread candidates', async t => {
  const f = await fixture(t);
  f.data.chats = Array.from({ length: 21 }, (_, index) => ({ ...structuredClone(fixtureData.chats[0]), id: `19:scan-${index}@thread.v2` }));
  f.data.teams = [];
  for (const chat of f.data.chats) f.data.messages[chat.id] = [];
  f.data.messages[f.data.chats[20].id] = [rawMessage(fixtureData.boundary + 1000)];
  const first = await f.triage.inbox(); assert.equal(first.items.length, 0); assert.equal(first.hasMore, true);
  const second = await f.triage.inbox({ cursor: first.nextCursor }); assert.equal(second.items.length, 1); assert.equal(second.hasMore, false);
  const cursor = f.triage.cursors.get(first.nextCursor); cursor.expires = 1;
  await assert.rejects(f.triage.inbox({ cursor: first.nextCursor }), { code: 'invalid_cursor' });
});

test('channel-root overflow continues without omitting or duplicating threads', async t => {
  const f = await fixture(t); f.data.chats = [];
  f.data.messages[channel] = Array.from({ length: 25 }, (_, index) => rawMessage(fixtureData.boundary + 1000 + index, `Root ${index}`));
  for (const parent of f.data.messages[channel]) f.data.messages[`${channel};messageid=${parent.id}`] = [];
  const ids = []; let cursor;
  do {
    const page = await f.triage.inbox({ limit: 20, cursor }); ids.push(...page.items.map(group => group.parentMessageId)); cursor = page.nextCursor;
  } while (cursor);
  assert.equal(ids.length, 25); assert.equal(new Set(ids).size, 25);
});

test('context refresh keeps the selected reply message even after a newer incoming message', async t => {
  const f = await fixture(t), first = (await f.triage.inbox()).items[0];
  f.data.messages[f.data.chats[0].id].push(rawMessage(fixtureData.boundary + 12000));
  const updated = (await f.triage.contexts({ targets: [first.replyTarget] })).items[0];
  assert.equal(f.triage.target(updated.replyTarget).selected.id, first.unread[0].id);
  assert.equal(updated.unread.length, 2);
});

test('channel root continuations validate server backward links and stop cycles', async t => {
  const f = await fixture(t); f.data.chats = [];
  f.links.set(channel, 'https://attacker.example/steal');
  const first = await f.triage.inbox(); assert.equal(first.hasMore, true);
  const second = await f.triage.inbox({ cursor: first.nextCursor });
  assert.equal(second.coverage.issues[0].code, 'api_cursor_rejected');
  assert.ok(f.requests.every(request => request.url.hostname !== 'attacker.example'));
  const other = await fixture(t); other.data.chats = [];
  other.links.set(channel, `${other.api.credentials.messageOrigin}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages?older=1`);
  const initial = await other.triage.inbox();
  const repeated = await other.triage.inbox({ cursor: initial.nextCursor });
  assert.equal(repeated.hasMore, false); assert.equal(repeated.items.length, 0);
  assert.ok(repeated.coverage.issues.some(issue => issue.reason === 'channel_root_traversal_stopped'));
});

test('preflight rate limits stop subsequent sends, and unsupported channel reads report partial coverage', async t => {
  const f = await fixture(t), groups = (await f.triage.inbox()).items; f.state.rateLimit = f.data.chats[0].id;
  const result = await f.triage.replies({ replies: [reply(groups[0]), reply(groups[1], 'after-rate-002')] });
  assert.equal(result.items[0].send.error.code, 'api_rate_limited'); assert.equal(result.items[1].send.status, 'not_attempted'); assert.equal(f.state.sends, 0);
  const other = await fixture(t, { protocols: { ...triageProtocols, channelThreads: false } });
  const inbox = await other.triage.inbox(); assert.equal(inbox.items.length, 3);
  assert.equal(inbox.coverage.issues[0].code, 'api_operation_unsupported');
});

test('caller abort reaches HTTP and batch cancellation prevents dispatch', async t => {
  const controller = new AbortController(); controller.abort();
  await assert.rejects(requestJSON('https://example.com', { signal: controller.signal }, async (_, options) => {
    assert.equal(options.signal.aborted, true); options.signal.throwIfAborted();
  }), { code: 'api_transport_failed' });
  const f = await fixture(t), group = (await f.triage.inbox()).items[0];
  const result = await f.triage.replies({ replies: [reply(group)] }, { signal: controller.signal });
  assert.equal(result.items[0].send.status, 'failed'); assert.equal(f.state.sends, 0);
});

test('agent-facing OpenAPI advertises three helpers with bounded batch input and retry guidance', () => {
  assert.equal(openapi.paths['/triage/inbox'].get.operationId, 'getTeamsTriageInbox');
  assert.equal(openapi.paths['/triage/contexts'].post.operationId, 'getTeamsReplyContexts');
  const operation = openapi.paths['/triage/replies'].post;
  assert.equal(operation.operationId, 'replyTeamsBatch');
  assert.equal(operation.requestBody.content['application/json'].schema.properties.replies.maxItems, 10);
  assert.match(operation.description, /SAME key/); assert.match(operation.description, /protocol_unverified/);
});
