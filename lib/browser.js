import fs from 'node:fs/promises';
import path from 'node:path';
import { CDP, findBrowser } from './cdp.js';
import { teamsDOM, selectorsFrom } from './dom.js';
import { AdapterError } from './errors.js';

const TEAMS_URL = 'https://teams.microsoft.com/v2/';
export class TeamsBrowser {
  constructor(profileDir, settings = {}, options = {}) {
    this.profileDir = profileDir;
    this.settings = settings;
    this.options = options;
    this.selectors = selectorsFrom(settings.selectors);
    this.queue = Promise.resolve();
    this.queued = 0;
    this.disposed = false;
  }
  serial(fn) {
    if (this.disposed) return Promise.reject(new AdapterError('disposed', 'Browser is being disposed.', 503));
    if (this.queued >= 20) return Promise.reject(new AdapterError('queue_full', 'Too many queued Teams operations. Try later.', 429));
    this.queued++;
    const result = this.queue.then(() => {
      if (this.disposed) throw new AdapterError('disposed', 'Browser has been disposed.', 503);
      return fn();
    }).finally(() => this.queued--);
    this.queue = result.catch(() => {});
    return result;
  }
  async start() {
    if (this.cdp && !this.cdp.closed) return;
    await fs.mkdir(this.profileDir, { recursive: true, mode: 0o700 });
    await fs.chmod(this.profileDir, 0o700);
    this.cdp = new CDP(await findBrowser(this.settings.browserPath), this.profileDir, this.options);
    try { this.page = await this.cdp.page(TEAMS_URL); }
    catch (error) { await this.cdp.close(); throw error; }
  }
  async dom(action, args = {}) {
    await this.start();
    const location = await this.page.evaluate(() => ({ host: location.hostname, protocol: location.protocol }));
    if (location.protocol !== 'https:' || !['teams.microsoft.com', 'teams.cloud.microsoft'].includes(location.host)) {
      if (action === 'status') return { ready: false, signInRequired: true };
      if (action === 'diagnostics') return { ready: false, signInRequired: true, host: location.host };
      throw new AdapterError('signin_required', 'Sign in in the dedicated Teams browser using the login operation.', 503);
    }
    return this.page.evaluate(teamsDOM, action, args, this.selectors);
  }
  async login() {
    await this.start();
    await this.page.call('Page.bringToFront');
    return { opened: true, instructions: 'Complete sign-in in this browser, then check status. Keep this dedicated Teams window open.' };
  }
  async listChats() {
    const result = await this.dom('chats');
    return { ...result, items: result.items.map(({ key, threadId, ...item }) => ({ ...item, id: Buffer.from(JSON.stringify({ key, title: item.title, threadId })).toString('base64url') })) };
  }
  decodeChat(id) {
    try {
      if (typeof id !== 'string' || id.length > 4096) throw new Error();
      const value = JSON.parse(Buffer.from(id, 'base64url').toString());
      if (!['data-fui-tree-item-value', 'data-chat-id', 'data-itemid', 'id', 'title'].includes(value.key?.attribute) || typeof value.key.value !== 'string' || !value.key.value || typeof value.title !== 'string' || !value.title || (value.threadId != null && typeof value.threadId !== 'string')) throw new Error();
      return value;
    } catch { throw new AdapterError('invalid_chat', 'Use a chat ID returned by listTeamsChats.', 400); }
  }
  async openChat(id) {
    const chat = this.decodeChat(id);
    await this.dom('open', chat);
    return chat;
  }
  async messages(id, limit = 50, olderPages = 0) {
    const chat = await this.openChat(id);
    const result = await this.dom('messages', { ...chat, limit: 200 });
    const items = result.items;
    for (let page = 0; page < olderPages; page++) {
      const older = await this.dom('older', chat);
      items.unshift(...older.items);
      if (!older.changed) break;
    }
    const seen = new Set();
    const unique = items.filter(m => {
      const key = m.id ?? JSON.stringify([m.text, m.author, m.timestamp]);
      if (seen.has(key)) return false;
      seen.add(key); return true;
    });
    return { chatId: id, title: chat.title, items: olderPages ? unique.slice(0, limit) : unique.slice(-limit), coverage: 'rendered_message_window', completeHistory: false };
  }
  async prepare(id) {
    const chat = await this.openChat(id);
    return { chat, ...await this.dom('prepare', chat) };
  }
  async send(prepared, text) {
    // Input.insertText dispatches normal editor input. Never inject HTML or press Enter
    // (Teams' Enter-to-send setting varies). Click only the verified Send button.
    await this.page.call('Input.insertText', { text });
    const point = await this.dom('sendPoint', { ...prepared.chat, text });
    await this.page.clickPoint(point);
    return this.dom('confirm', { ...prepared.chat, text, beforeIds: prepared.beforeIds });
  }
  async close() { this.disposed = true; await this.queue; await this.cdp?.close(); }
}

export const profilePath = (dataDir, profile) => {
  if (!/^[a-f0-9-]{36}$/.test(profile)) throw new AdapterError('invalid_profile', 'Invalid browser profile.', 400);
  return path.join(dataDir, 'profiles', profile);
};
