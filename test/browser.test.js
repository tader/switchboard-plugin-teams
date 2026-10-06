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
