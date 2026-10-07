import fs from 'node:fs/promises';
import path from 'node:path';
import { CDP, findBrowser } from './cdp.js';
import { teamsActions } from './actions-dom.js';
import { teamsDOM, selectorsFrom } from './dom.js';
import { AdapterError } from './errors.js';
import { randomUUID } from 'node:crypto';
import { recipientEmail, peopleQuery, messageQuery, messageId, messageText, reactionInput } from './input.js';

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
    return this.page.evaluate(action.startsWith('action:') ? teamsActions : teamsDOM, action.replace(/^action:/, ''), args, this.selectors);
  }
  async login() {
    await this.start();
    await this.page.call('Page.bringToFront');
    return { opened: true, instructions: 'Complete sign-in in this browser, then check status. Keep this dedicated Teams window open.' };
  }
  chatId(chat) { return Buffer.from(JSON.stringify(chat)).toString('base64url'); }
  async listChats(unreadOnly = false) {
    const result = await this.dom('chats');
    return { ...result, items: result.items.filter(item => !unreadOnly || item.unread).map(({ key, threadId, ...item }) => ({ ...item, id: this.chatId({ key, title: item.title, threadId }) })) };
  }
  encodeChats(result, unreadOnly = false) {
    return { ...result, items: result.items.filter(item => !unreadOnly || item.unread).map(({ key, threadId, ...item }) => ({ ...item, id: this.chatId({ key, title: item.title, threadId }) })) };
  }
  async scanChats(unreadOnly = true, maxWindows = 100, target) {
    const token = randomUUID(), found = new Map();
    let complete = false, windows = 0, endSignature, matched, restoration;
    try {
      await this.dom('scanBegin', { token, unreadOnly });
      for (; windows < maxWindows; windows++) {
        const result = await this.dom(windows ? 'scanNext' : 'scanWindow', { token });
        for (const item of result.items) found.set(item.threadId || JSON.stringify(item.key), item);
        if (target) {
          matched = result.items.find(item => JSON.stringify(item.key) === JSON.stringify(target.key));
          if (matched) { await this.dom('open', target); break; }
        }
        const signature = JSON.stringify([result.position, result.height, [...found.keys()]]);
        // Observe a stable terminal window twice: lazy loading may append rows.
        if (result.atEnd && !result.collapsed && !result.more && signature === endSignature) { complete = true; windows++; break; }
        endSignature = result.atEnd ? signature : null;
      }
    } finally { restoration = await this.dom('scanEnd', { token }).catch(error => { if (error.code !== 'scan_changed') throw error; }); }
    if (target && !matched) throw new AdapterError('chat_unavailable', 'Chat was not found during sidebar discovery. Refresh chat IDs.', 404);
    return this.encodeChats({ items: [...found.values()], discoveryComplete: complete, truncated: !complete, windowsScanned: windows, coverage: 'native_unread_chat_list', completeAccount: false, uiRestored: restoration?.restored ?? false, unrestoredSections: restoration?.unrestoredSections ?? 0, excludes: ['channels', 'other_tenants', 'hidden_chats'] }, unreadOnly);
  }
  async searchMessages(query, limit = 50, maxPages = 3) {
    query = messageQuery(query);
    const focus = await this.dom('searchFocus');
    if (!focus.focused) throw new AdapterError('search_focus_changed', 'Could not focus the Teams search input.');
    await this.page.call('Input.insertText', { text: query });
    await this.dom('searchSubmit', { query });
    await this.page.call('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r', unmodifiedText: '\r' });
    await this.page.call('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
    const items = new Map(); let hasNext = false, pages = 0;
    for (; pages < maxPages; pages++) {
      const result = await this.dom('searchResults', { query }); hasNext = result.hasNext;
      for (const { key, ...item } of result.items) items.set(key, { ...item, resultId: this.chatId({ query, key }), messageId: null });
      if (!hasNext || items.size >= limit) { pages++; break; }
      if (pages + 1 < maxPages) await this.dom('searchNext', { query });
    }
    return { query, items: [...items.values()].slice(0, limit), pagesScanned: pages, hasMore: hasNext || items.size > limit, coverage: 'native_message_search', completeSearch: !hasNext && items.size <= limit };
  }
  decodeSearch(resultId) {
    try {
      if (typeof resultId !== 'string' || resultId.length > 4096) throw new Error();
      const value = JSON.parse(Buffer.from(resultId, 'base64url').toString());
      value.query = messageQuery(value.query);
      if (typeof value.key !== 'string' || !value.key.startsWith('serp-message-card-content-') || value.key.length > 2000) throw new Error();
      return value;
    } catch { throw new AdapterError('invalid_search_result', 'Use resultId returned by searchTeamsMessages.', 400); }
  }
  async openSearch(resultId) {
    const { query, key } = this.decodeSearch(resultId);
    // Re-run the native query, then locate this exact result key. Never open by
    // display name, ordinal, or snippet, which may change between requests.
    await this.searchMessages(query, 1, 1);
    for (let page = 0; page < 10; page++) {
      const result = await this.dom('searchResults', { query });
      if (result.items.some(item => item.key === key)) return this.dom('searchOpen', { query, resultKey: key });
      if (!result.hasNext) break;
      await this.dom('searchNext', { query });
    }
    throw new AdapterError('search_result_unavailable', 'The selected result was not found in the first ten search pages. Search again.', 404);
  }
  async searchContext(resultId) {
    const chat = await this.openSearch(resultId);
    return { chatId: this.chatId(chat), title: chat.title, ...await this.dom('messages', { ...chat, limit: 100 }), completeHistory: false, mayMarkRead: true };
  }
  async unreadMessages(maxChats = 3, limitPerChat = 50) {
    // Snapshot unread flags BEFORE opening anything: Teams marks opened chats read.
    const snapshot = await this.scanChats(true);
    const chats = [];
    for (const summary of snapshot.items.slice(0, maxChats)) {
      try {
        const chat = await this.openChat(summary.id);
        chats.push({ chatId: summary.id, title: summary.title, unreadAtSnapshot: true, ...await this.dom('unreadMessages', { ...chat, limit: limitPerChat }) });
      } catch (error) {
        chats.push({ chatId: summary.id, title: summary.title, items: [], error: { code: error.code ?? 'read_failed', message: error.message } });
      }
    }
    return {
      items: chats.flatMap(chat => chat.items.map(message => ({ chatId: chat.chatId, chatTitle: chat.title, ...message }))),
      chats, unreadChatsAtSnapshot: snapshot.items.length, remainingChats: Math.max(0, snapshot.items.length - maxChats),
      mayMarkRead: true, discoveryComplete: snapshot.discoveryComplete, completeAccount: false, coverage: 'native_unread_chat_list',
    };
  }
  async searchPeople(query) {
    query = peopleQuery(query);
    const before = await this.dom('peopleFocus');
    await this.page.call('Input.insertText', { text: query });
    return this.dom('peopleResults', { ...before, query });
  }
  async openRecipient(email) {
    email = recipientEmail(email);
    const result = await this.searchPeople(email);
    if (result.items.filter(person => person.email === email).length !== 1) throw new AdapterError('recipient_unavailable', 'No unique directory person matches that email. Search people and use a returned email.', 404);
    return this.dom('selectPerson', { email });
  }
  async conversation(email) {
    email = recipientEmail(email);
    const chat = await this.openRecipient(email);
    return { chatId: this.chatId(chat), title: chat.title, email, status: 'conversation_open', messageSent: false };
  }
  decodeChat(id) {
    try {
      if (typeof id !== 'string' || id.length > 4096) throw new Error();
      const value = JSON.parse(Buffer.from(id, 'base64url').toString());
      if (!['data-fui-tree-item-value', 'data-chat-id', 'data-itemid', 'id', 'title', 'recipient', 'search'].includes(value.key?.attribute) || typeof value.key.value !== 'string' || !value.key.value || typeof value.title !== 'string' || !value.title || (value.threadId != null && typeof value.threadId !== 'string')) throw new Error();
      if (value.key.attribute === 'recipient') value.key.value = recipientEmail(value.key.value);
      return value;
    } catch { throw new AdapterError('invalid_chat', 'Use a chat ID returned by listTeamsChats.', 400); }
  }
  async openChat(id) {
    const chat = this.decodeChat(id);
    if (chat.key.attribute === 'search') {
      let target; try { target = JSON.parse(chat.key.value); } catch { throw new AdapterError('invalid_chat', 'Invalid search conversation.', 400); }
      const opened = await this.openSearch(this.chatId({ query: target.query, key: target.resultKey }));
      if (opened.threadId !== chat.threadId) throw new AdapterError('chat_changed', 'The search result opened a different thread.');
      return opened;
    }
    if (chat.key.attribute === 'recipient') return this.openRecipient(chat.key.value);
    try { await this.dom('open', chat); }
    catch (error) { if (error.code !== 'chat_unavailable') throw error; await this.scanChats(false, 100, chat); }
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
  async prepareRecipient(email) {
    const chat = await this.openRecipient(email);
    return { chat, ...await this.dom('prepare', chat) };
  }
  async prepareQuote(id, replyToMessageId) {
    replyToMessageId = messageId(replyToMessageId);
    const chat = await this.openChat(id);
    let point;
    for (let page = 0; page <= 5; page++) {
      try { point = await this.dom('quoteTarget', { ...chat, replyToMessageId }); break; }
      catch (error) {
        if (error.code !== 'message_unavailable' || page === 5) throw error;
        const older = await this.dom('older', chat);
        if (!older.changed) throw error;
      }
    }
    await this.page.call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'right', clickCount: 1 });
    await this.page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'right', clickCount: 1 });
    return { chat, ...await this.dom('quoteSelect', { ...chat, replyToMessageId }) };
  }
  async contextClick(point) {
    await this.page.call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'right', clickCount: 1 });
    await this.page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'right', clickCount: 1 });
  }
  async prepareMessageAction(id, mid, kind, payload = {}) {
    mid = messageId(mid);
    if (['edit', 'delete'].includes(kind)) messageText(payload.expectedText, 'expectedText', true);
    if (kind === 'edit') messageText(payload.text);
    if (kind === 'reaction') payload = reactionInput(payload.reaction, payload.selected);
    const chat = await this.openChat(id), args = { ...chat, messageId: mid, token: randomUUID(), kind, ...payload };
    let point;
    const cleanup = () => this.dom('action:release', args);
    try {
      for (let page = 0; page <= 5; page++) {
        try { point = await this.dom('action:messageTarget', args); break; }
        catch (error) { if (error.code !== 'message_unavailable' || page === 5) throw error; if (!(await this.dom('older', chat)).changed) throw error; }
      }
      if (['reaction', 'reactions'].includes(kind)) {
        await this.page.call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...point });
        const result = await this.dom('action:reactionReady', args);
        if (kind === 'reactions') return { result, cleanup };
        await this.dom('action:reactionPrepare', args);
      } else {
        await this.contextClick(point); await this.dom('action:menuPrepare', args);
        if (kind === 'edit') await this.dom('action:editPrepare', args);
      }
      return { args, cleanup };
    } catch (error) { await cleanup(); throw error; }
  }
  async reactions(id, mid) {
    const prepared = await this.prepareMessageAction(id, mid, 'reactions');
    try { return prepared.result; } finally { await prepared.cleanup(); }
  }
  async executeMessageAction(prepared, checkCancelled = () => {}) {
    const { args } = prepared;
    if (args.kind === 'edit') {
      await this.dom('action:editSelect', args); checkCancelled();
      await this.page.call('Input.insertText', { text: args.text });
    }
    checkCancelled();
    return this.dom(`action:${args.kind}Commit`, args);
  }
  async prepareUnread(id) {
    await this.dom('status');
    let chat = this.decodeChat(id);
    if (['recipient', 'search'].includes(chat.key.attribute)) chat = await this.openChat(id);
    const args = { ...chat, token: randomUUID(), kind: 'unread' };
    let scanToken, point;
    const cleanup = async () => {
      try { await this.dom('action:release', args); }
      finally { if (scanToken) { const token = scanToken; scanToken = null; await this.dom('scanEnd', { token }); } }
    };
    try {
      try { point = await this.dom('action:unreadTarget', args); }
      catch (error) {
        if (error.code !== 'chat_unavailable') throw error;
        scanToken = randomUUID(); await this.dom('scanBegin', { token: scanToken, unreadOnly: false });
        for (let page = 0; page < 100; page++) {
          const window = await this.dom(page ? 'scanNext' : 'scanWindow', { token: scanToken });
          if (window.items.some(item => item.threadId === chat.threadId)) { point = await this.dom('action:unreadTarget', args); break; }
          if (window.atEnd && !window.more && !window.collapsed) break;
        }
        if (!point) throw new AdapterError('chat_unavailable', 'Could not find this chat during sidebar discovery.', 404);
      }
      if (!point.alreadyUnread) { await this.contextClick(point); await this.dom('action:unreadPrepare', args); }
      return { args, cleanup };
    } catch (error) { await cleanup(); throw error; }
  }
  async send(prepared, text) {
    // Input.insertText dispatches normal editor input. Never inject HTML or press Enter
    // (Teams' Enter-to-send setting varies). Click only the verified Send button.
    await this.page.call('Input.insertText', { text });
    const point = await this.dom('sendPoint', { ...prepared.chat, text, replyToMessageId: prepared.replyToMessageId });
    await this.page.clickPoint(point);
    return this.dom('confirm', { ...prepared.chat, text, replyToMessageId: prepared.replyToMessageId, quotePreview: prepared.quotePreview, beforeIds: prepared.beforeIds });
  }
  async close() { this.disposed = true; await this.queue; await this.cdp?.close(); }
}

export const profilePath = (dataDir, profile) => {
  if (!/^[a-f0-9-]{36}$/.test(profile)) throw new AdapterError('invalid_profile', 'Invalid browser profile.', 400);
  return path.join(dataDir, 'profiles', profile);
};
