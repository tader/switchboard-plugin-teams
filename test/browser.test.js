import test from 'node:test';
import assert from 'node:assert/strict';
import { TeamsBrowser, profilePath } from '../lib/browser.js';
import { selectorsFrom } from '../lib/dom.js';

test('profile queue serializes operations after failures and isolates profiles', async () => {
  const a = new TeamsBrowser('/tmp/a'), b = new TeamsBrowser('/tmp/b');
  const events = [];
  let unlock;
  const barrier = new Promise(resolve => unlock = resolve);
  const first = a.serial(async () => { events.push('a1'); await barrier; throw new Error('failure'); });
  const second = a.serial(async () => events.push('a2'));
  await b.serial(async () => events.push('b1'));
  assert.deepEqual(events, ['a1', 'b1']);
  unlock(); await assert.rejects(first); await second;
  assert.deepEqual(events, ['a1', 'b1', 'a2']);
});

test('profile IDs cannot traverse filesystem and chat IDs are validated', () => {
  assert.throws(() => profilePath('/tmp', '../escape'), { code: 'invalid_profile' });
  const browser = new TeamsBrowser('/tmp/a');
  for (const id of ['bad', '', Buffer.from(JSON.stringify({key:{attribute:'onclick', value:'x'},title:'x'})).toString('base64url')]) assert.throws(() => browser.decodeChat(id), { code: 'invalid_chat' });
  const chat = { key: { attribute: 'data-chat-id', value: 'a"b' }, title: 'A chat' };
  assert.deepEqual(browser.decodeChat(Buffer.from(JSON.stringify(chat)).toString('base64url')), chat);
});

test('selector overrides reject unsupported fields and empty selectors', () => {
  assert.equal(selectorsFrom('{"send":"button.send"}').send, 'button.send');
  for (const value of ['[]', '{"cookie":"x"}', '{"send":""}', '{"send":true}']) assert.throws(() => selectorsFrom(value));
});

test('unread reads snapshot flags before opening chats, cap work and separate fallback messages', async () => {
  const browser = new TeamsBrowser('/tmp/unread');
  const events = [];
  browser.listChats = async unreadOnly => {
    assert.equal(unreadOnly, true); events.push('snapshot');
    return { items: [{ id: 'a', title: 'Alice' }, { id: 'b', title: 'Bob' }, { id: 'c', title: 'Carol' }] };
  };
  browser.openChat = async id => { events.push('open-' + id); return { title: id }; };
  browser.dom = async (_action, chat) => chat.title === 'a' ? { items: [{ id: 'm1', text: 'Unread' }], boundaryFound: true } : { items: [], recentMessages: [{ id: 'old', text: 'Unconfirmed' }], boundaryFound: false };
  const result = await browser.unreadMessages(2, 10);
  assert.deepEqual(events, ['snapshot', 'open-a', 'open-b']);
  assert.equal(result.items.length, 1);
  assert.equal(result.chats[1].recentMessages[0].id, 'old');
  assert.equal(result.remainingChats, 1);
  assert.equal(result.mayMarkRead, true);
  assert.equal(result.completeAccount, false);
});

test('person conversations reselect by normalized email and never trust an encoded display name', async () => {
  const browser = new TeamsBrowser('/tmp/person');
  const person = { key: { attribute: 'recipient', value: 'alice@example.com' }, title: 'Verified Alice', threadId: null };
  browser.openRecipient = async email => { assert.equal(email, 'alice@example.com'); return person; };
  const forged = { ...person, title: 'Someone else', threadId: 'wrong-thread' };
  assert.deepEqual(await browser.openChat(browser.chatId(forged)), person);
  const opened = await browser.conversation('Alice@Example.COM');
  assert.equal(opened.messageSent, false);
  assert.equal(opened.email, 'alice@example.com');
  assert.equal(browser.decodeChat(opened.chatId).title, 'Verified Alice');
});
