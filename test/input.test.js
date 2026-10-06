import test from 'node:test';
import assert from 'node:assert/strict';
import { recipientEmail, peopleQuery } from '../lib/input.js';

test('recipient identity accepts one email/UPN and rejects names, groups and URLs', () => {
  assert.equal(recipientEmail(' Alice+Teams@Example.COM '), 'alice+teams@example.com');
  for (const value of ['Alice Adams', 'alice@example.com,bob@example.com', 'alice@example.com bob@example.com', 'https://example.com', null]) assert.throws(() => recipientEmail(value), { code: 'invalid_recipient' });
});

test('people search is bounded and requires meaningful text', () => {
  assert.equal(peopleQuery(' Alice Adams '), 'Alice Adams');
  for (const value of ['a', '', 'a\0b', 'a'.repeat(151), null]) assert.throws(() => peopleQuery(value), { code: 'invalid_query' });
});
