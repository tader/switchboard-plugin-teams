import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { AdapterError } from './errors.js';

// Durable intent BEFORE any editor input or send click. A process crash must never
// turn a repeated request into a second message. Called inside the profile queue.
export class SendLedger {
  constructor(directory) { this.file = path.join(directory, 'send-ledger.json'); }
  async read() {
    try { return JSON.parse(await fs.readFile(this.file, 'utf8')); }
    catch (error) { if (error.code === 'ENOENT') return {}; throw error; }
  }
  async write(records) {
    await fs.mkdir(path.dirname(this.file), { recursive: true, mode: 0o700 });
    const temp = `${this.file}.${randomUUID()}.tmp`;
    const handle = await fs.open(temp, 'wx', 0o600);
    try { await handle.writeFile(JSON.stringify(records)); await handle.sync(); }
    finally { await handle.close(); }
    await fs.rename(temp, this.file);
    const directory = await fs.open(path.dirname(this.file), 'r');
    try { await directory.sync(); } finally { await directory.close(); }
  }
  async send(key, chatId, text, prepare, execute) {
    if (typeof key !== 'string' || !/^[A-Za-z0-9._:-]{8,128}$/.test(key)) throw new AdapterError('invalid_idempotency_key', 'idempotencyKey must be 8–128 letters, digits or . _ : - characters.', 400);
    if (typeof text !== 'string' || !text.trim() || text.length > 20_000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text)) throw new AdapterError('invalid_message', 'text must be nonempty plain text, at most 20,000 characters, without control characters.', 400);
    const records = await this.read();
    const index = createHash('sha256').update(key).digest('hex');
    const fingerprint = createHash('sha256').update(JSON.stringify([chatId, text])).digest('hex');
    const previous = records[index];
    if (previous) {
      if (previous.fingerprint !== fingerprint) throw new AdapterError('idempotency_conflict', 'This key was already used for a different chat or message.');
      if (previous.result) return { ...previous.result, replayed: true };
      throw new AdapterError('send_uncertain', 'This send was already attempted. Check Teams manually; do not retry with a new key.');
    }
    if (Object.keys(records).length >= 10_000) throw new AdapterError('ledger_full', 'Send ledger is full. Archive it manually after checking outstanding sends.', 503);
    const prepared = await prepare();
    records[index] = { fingerprint, attemptedAt: new Date().toISOString() };
    await this.write(records);
    const result = await execute(prepared);
    records[index].result = result;
    await this.write(records);
    return result;
  }
}
