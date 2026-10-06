import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { SendLedger } from '../lib/ledger.js';

test('send intent survives restart, replays result, and binds key to content', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  let sends = 0;
  const result = { status: 'observed_in_chat', message: { id: 'm2' } };
  const ledger = new SendLedger(dir);
  await ledger.send('request-001', 'chat1', 'hello', async () => 'prepared', async prepared => {
    assert.equal(prepared, 'prepared');
    const records = await ledger.read();
    assert.equal(Object.values(records).length, 1);
    assert.equal(Object.values(records)[0].result, undefined);
    sends++; return result;
  });
  assert.deepEqual(await new SendLedger(dir).send('request-001', 'chat1', 'hello', () => assert.fail('prepare repeated'), () => assert.fail('send repeated')), { ...result, replayed: true });
  await assert.rejects(() => ledger.send('request-001', 'chat2', 'hello', () => {}, () => {}), { code: 'idempotency_conflict' });
  await assert.rejects(() => ledger.send('request-001', 'chat1', 'changed', () => {}, () => {}), { code: 'idempotency_conflict' });
  assert.equal(sends, 1);
});

test('uncertain attempts never send again, including after a crash', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir);
  await assert.rejects(() => ledger.send('request-002', 'chat1', 'hello', async () => ({}), async () => { throw new Error('browser crashed'); }), /browser crashed/);
  await assert.rejects(() => new SendLedger(dir).send('request-002', 'chat1', 'hello', () => assert.fail(), () => assert.fail()), { code: 'send_uncertain' });
});

test('preflight failures remain retryable and invalid payloads never reach browser', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir);
  await assert.rejects(() => ledger.send('request-003', 'chat1', 'hello', async () => { throw new Error('draft exists'); }, () => assert.fail()), /draft exists/);
  assert.deepEqual(await ledger.read(), {});
  for (const [key, text] of [['short', 'hello'], ['request-003', ' '], ['request-003', '\0'], ['request-003', 'a'.repeat(20001)]]) {
    await assert.rejects(() => ledger.send(key, 'chat1', text, () => assert.fail(), () => assert.fail()), error => error.status === 400);
  }
});

test('corrupted durable state stops sends rather than resetting duplicate protection', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir);
  await fs.writeFile(ledger.file, 'broken JSON');
  await assert.rejects(() => ledger.send('request-004', 'chat1', 'hello', () => assert.fail(), () => assert.fail()), SyntaxError);
});
