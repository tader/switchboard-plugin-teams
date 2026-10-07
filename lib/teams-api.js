import { randomUUID, randomBytes } from 'node:crypto';
import { AdapterError } from './errors.js';
import { requestJSON } from './api-http.js';
import { messageId, recipientEmail, messageText, reactionInput } from './input.js';
import { richContent } from './content.js';

export const unsupported = () => { throw new AdapterError('api_operation_unsupported', 'This operation has not been migrated to the Teams API. See getTeamsCapabilities.', 501); };
export const capabilities = {
  transport: 'api', browserUse: 'authentication_only', authModes: ['browser', 'tokens'],
  supported: ['status', 'chats', 'unread_chats', 'unread_messages', 'messages', 'chat_history', 'channels', 'channel_history', 'exact_email_lookup', 'existing_one_to_one', 'send_chat', 'formatted_chat', 'edit_chat', 'delete_chat', 'chat_reactions'],
  unsupported: ['message_search', 'new_conversation', 'quoted_replies', 'mentions', 'mark_unread', 'channel_thread_replies', 'send_channel_thread', 'attachments'],
  liveValidated: ['oauth_tokens', 'profile_read', 'conversation_list', 'recent_messages'],
};
const encode = (kind, threadId) => Buffer.from(JSON.stringify({ version: 1, kind, threadId })).toString('base64url');
const escape = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const entities = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
const arrivalMillis = value => {
  if (typeof value === 'number') return value;
  if (typeof value === 'string' && /^\d{13}$/.test(value)) return Number(value);
  return typeof value === 'string' ? Date.parse(value) : NaN;
};
export function plainText(content, type = '') {
  if (!/html/i.test(type)) return String(content ?? '');
  return String(content ?? '').replace(/<br\s*\/?\s*>/gi, '\n').replace(/<\/(p|div|li|blockquote)>/gi, '\n').replace(/<[^>]*>/g, '')
    .replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, (full, entity) => {
      if (!entity.startsWith('#')) return entities[entity.toLowerCase()] ?? full;
      const number = entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
      return number > 0 && number <= 0x10ffff ? String.fromCodePoint(number) : full;
    }).trimEnd();
}
function arrayProperty(value) {
  if (typeof value === 'string') { try { value = JSON.parse(value); } catch { return []; } }
  return Array.isArray(value) ? value : [];
}
function normalize(raw) {
  return { id: String(raw.id), text: plainText(raw.content, raw.messagetype ?? raw.messageType),
    author: raw.imdisplayname ?? raw.imDisplayName ?? null, authorId: raw.from ?? null,
    time: raw.originalarrivaltime ?? raw.originalArrivalTime ?? raw.composetime ?? raw.composeTime ?? null,
    richText: /html/i.test(raw.messagetype ?? raw.messageType ?? '') ? String(raw.content ?? '') : null,
    deleted: Boolean(raw.properties?.deletetime ?? raw.properties?.deleteTime),
    reactions: arrayProperty(raw.properties?.emotions), mentions: arrayProperty(raw.properties?.mentions),
    parentMessageId: raw.properties?.parentMessageId == null ? null : String(raw.properties.parentMessageId) };
}
function order(items) { return items.sort((a, b) => String(a.time ?? a.id).localeCompare(String(b.time ?? b.id))); }
export function messagePayload(input, name) {
  if (input.replyToMessageId !== undefined) unsupported();
  let content, messagetype;
  if (input.content !== undefined) {
    const blocks = richContent(input.content);
    const runs = values => values.map(run => {
      if (run.mention) unsupported();
      let value = escape(run.text).replace(/\n/g, '<br>');
      for (const mark of run.marks) { const tag = { bold: 'strong', italic: 'em', underline: 'u', strike: 's', code: 'code' }[mark]; value = `<${tag}>${value}</${tag}>`; }
      return run.link ? `<a href="${escape(run.link)}">${value}</a>` : value;
    }).join('');
    content = blocks.map(block => {
      if (block.runs) { const tag = block.type === 'quote' ? 'blockquote' : 'p'; return `<${tag}>${runs(block.runs)}</${tag}>`; }
      const tag = block.type === 'numberedList' ? 'ol' : 'ul'; return `<${tag}>${block.items.map(item => `<li>${runs(item)}</li>`).join('')}</${tag}>`;
    }).join(''); messagetype = 'RichText/Html';
  } else { content = messageText(input.text); messagetype = 'Text'; }
  return { content, messagetype, contenttype: 'text', imdisplayname: name, properties: { importance: '', subject: null } };
}

export class TeamsAPI {
  constructor(session, { fetcher = fetch } = {}) { this.session = session; this.fetcher = fetcher; this.cursors = new Map(); }
  get credentials() {
    const c = this.session.credentials;
    if (!c) throw new AdapterError('signin_required', 'Reconnect this connection to enable API authentication.', 503);
    return c;
  }
  async request(service, suffix, options = {}) {
    const c = this.credentials;
    const base = service === 'messages' ? c.messageOrigin + '/v1/' : service === 'csa' ? 'https://teams.microsoft.com/api/csa/api/v1/' : `https://teams.microsoft.com/api/mt/${c.region}/beta/`;
    const url = new URL(suffix, base);
    if (!url.href.startsWith(base) || url.username || url.password) throw new AdapterError('api_url_rejected', 'Invalid Teams API target.', 400);
    const auth = service === 'messages' ? { Authentication: `skypetoken=${c.chatToken.value}` } : { Authorization: `Bearer ${c.tokens[service === 'csa' ? 'chatsvcagg' : 'skype'].value}` };
    return requestJSON(url.href, { ...options, headers: { ...auth, ...(options.body ? { 'content-type': 'application/json' } : {}) } }, this.fetcher);
  }
  async directory(force = false) {
    if (!force && this.directoryCache?.expires > Date.now()) return this.directoryCache.value;
    const value = await this.request('csa', 'teams/users/me?isPrefetch=false&enableMembershipSummary=true');
    if (!value || !Array.isArray(value.chats) || !Array.isArray(value.teams)) throw new AdapterError('api_schema_changed', 'Teams conversation response no longer matches the expected shape.', 502);
    this.directoryCache = { value, expires: Date.now() + 30_000 }; return value;
  }
  selfMri() { return `8:orgid:${this.credentials.identity.oid}`; }
  async profile(email = this.credentials.identity.email) {
    email = recipientEmail(email);
    const result = await this.request('mt', `users/${encodeURIComponent(email)}/?throwIfNotFound=false&isMailAddress=true&enableGuest=true&includeIBBarredUsers=true&skypeTeamsInfo=true`);
    const person = result?.value;
    const emails = [person?.email, person?.userPrincipalName, person?.upn].filter(value => typeof value === 'string').map(value => value.toLowerCase());
    if (!person || !emails.includes(email) || typeof person.mri !== 'string') throw new AdapterError('person_unavailable', 'Teams did not return an exact matching directory person.', 404);
    return { email, name: person.displayName || email, mri: person.mri };
  }
  async people(query) { const item = await this.profile(query); return { query, items: [item], coverage: 'exact_email_lookup' }; }
  async chats(unreadOnly = false) {
    const data = await this.directory(true);
    const items = data.chats.filter(chat => !chat.isConversationDeleted && (!unreadOnly || chat.isRead === false)).map(chat => ({
      id: encode('chat', chat.id), title: chat.title || chat.members?.filter(member => member.mri !== this.selfMri()).map(member => member.friendlyName || member.mri).join(', ') || 'Teams chat',
      unread: typeof chat.isRead === 'boolean' ? !chat.isRead : null, hidden: Boolean(chat.hidden),
      preview: chat.lastMessage ? plainText(chat.lastMessage.content, chat.lastMessage.messageType) : null,
      time: chat.lastMessage?.originalArrivalTime ?? null,
    }));
    const partial = Boolean(data.metadata?.isPartialData || data.metadata?.hasMoreChats || data.metadata?.forwardSyncToken);
    return { items, discoveryComplete: !partial, truncated: partial, completeAccount: false, coverage: 'api_conversation_snapshot', mayMarkRead: false };
  }
  async channels() {
    const data = await this.directory();
    return { items: data.teams.flatMap(team => (team.channels || []).filter(channel => !channel.isDeleted).map(channel => ({ id: encode('channel', channel.id), title: channel.displayName, teamTitle: team.displayName }))), coverage: 'api_conversation_snapshot', completeAccount: false, mayMarkRead: false };
  }
  async resolve(id, kind = 'chat') {
    let value;
    try { if (typeof id !== 'string' || id.length > 4096) throw new Error(); value = JSON.parse(Buffer.from(id, 'base64url')); }
    catch { throw new AdapterError('invalid_conversation', 'Use a conversation ID returned by the Teams listing.', 400); }
    // Old DOM IDs with an exact threadId can migrate after server membership lookup.
    if (typeof value.threadId !== 'string' || !value.threadId || value.threadId.length > 1024 || /[\u0000-\u001f/]/.test(value.threadId) || value.kind && value.kind !== kind || value.parentMessageId) throw new AdapterError('invalid_conversation', 'Refresh the conversation ID with listTeamsChats or listTeamsChannels.', 400);
    const data = await this.directory();
    const match = kind === 'chat' ? data.chats.find(chat => chat.id === value.threadId) : data.teams.flatMap(team => team.channels || []).find(channel => channel.id === value.threadId);
    if (!match || match.isDeleted || match.isConversationDeleted) throw new AdapterError('conversation_unavailable', 'Conversation was not returned for this account.', 404);
    return match;
  }
  async existingConversation(email) {
    const person = await this.profile(email), data = await this.directory(true);
    const matches = data.chats.filter(chat => chat.isOneOnOne && !chat.isConversationDeleted && chat.members?.some(member => member.mri === person.mri) && chat.members?.some(member => member.mri === this.selfMri()));
    if (!matches.length) unsupported();
    if (matches.length !== 1) throw new AdapterError('ambiguous_conversation', 'Multiple one-to-one chats match this person. Select a chat from listTeamsChats.', 409);
    return { email: person.email, title: person.name, chatId: encode('chat', matches[0].id), messageSent: false };
  }
  messagePath(threadId) { return `users/ME/conversations/${encodeURIComponent(threadId)}/messages`; }
  async page(threadId, limit = 50, backwardLink) {
    const expected = this.messagePath(threadId);
    let suffix = `${expected}?view=msnp24Equivalent%7CsupportsMessageProperties&pageSize=${limit}&startTime=1`;
    if (backwardLink) {
      let url;
      try { url = new URL(backwardLink, this.credentials.messageOrigin); }
      catch { throw new AdapterError('api_cursor_rejected', 'Teams returned an invalid history link.', 502); }
      if (url.origin !== this.credentials.messageOrigin || url.username || url.password || url.hash || decodeURIComponent(url.pathname) !== decodeURIComponent('/v1/' + expected)) throw new AdapterError('api_cursor_rejected', 'Teams history link changed host or conversation.', 502);
      suffix = url.href;
    }
    const response = await this.request('messages', suffix);
    if (!Array.isArray(response?.messages) || response.messages.some(item => typeof item?.id !== 'string' && typeof item?.id !== 'number')) throw new AdapterError('api_schema_changed', 'Teams message response no longer matches the expected shape.', 502);
    return { raw: response.messages, items: order(response.messages.map(normalize)), next: response._metadata?.backwardLink || null };
  }
  async messages(id, limit = 50, olderPages = 0, kind = 'chat') {
    const conversation = await this.resolve(id, kind), seen = new Map(); let page, next;
    for (let i = 0; i <= olderPages; i++) {
      page = await this.page(conversation.id, limit, next);
      for (const message of page.items) seen.set(message.id, message);
      if (!page.next || page.next === next) break; next = page.next;
    }
    const items = order([...seen.values()]);
    return { conversationId: id, items: olderPages ? items.slice(0, limit) : items.slice(-limit), hasMore: Boolean(page.next), completeHistory: false, coverage: 'api_messages', mayMarkRead: false };
  }
  async history(id, { limit = 100, cursor, format = 'json', kind = 'chat' } = {}) {
    for (const [key, state] of this.cursors) if (state.expires < Date.now()) this.cursors.delete(key);
    let state;
    if (cursor) {
      state = this.cursors.get(cursor);
      if (!state || state.id !== id || state.limit !== limit || state.format !== format || state.kind !== kind) throw new AdapterError('invalid_cursor', 'History cursor expired or belongs to different parameters.', 400);
      if (state.result) return state.result;
    } else { const conversation = await this.resolve(id, kind); state = { id, limit, format, kind, threadId: conversation.id, seen: [], links: [], next: null }; }
    if (this.cursors.size >= 1000 || state.seen.length >= 100_000) throw new AdapterError('history_capacity', 'History traversal capacity reached.', 429);
    const page = await this.page(state.threadId, limit, state.next);
    const seen = new Set(state.seen), items = page.items.filter(message => { if (seen.has(message.id)) return false; seen.add(message.id); return true; });
    const hasMore = Boolean(page.next && page.next !== state.next && !state.links.includes(page.next)), nextCursor = hasMore ? randomUUID() : null;
    const result = { conversationId: id, items, hasMore, nextCursor, completeHistory: false, coverage: 'api_history', mayMarkRead: false,
      ...(format === 'ndjson' ? { ndjson: items.map(item => JSON.stringify(item)).join('\n') + (items.length ? '\n' : '') } : {}) };
    if (nextCursor) this.cursors.set(nextCursor, { ...state, seen: [...seen], links: [...state.links, page.next], next: page.next, result: undefined, expires: Date.now() + 1800_000 });
    if (cursor) state.result = result;
    return result;
  }
  async unread(maxChats = 3, limit = 50) {
    const snapshot = await this.chats(true), chats = [], items = [];
    for (const chat of snapshot.items.slice(0, maxChats)) {
      const conversation = await this.resolve(chat.id), page = await this.messages(chat.id, limit);
      const horizon = conversation.userConsumptionHorizon || conversation.consumptionHorizon;
      const boundary = Number(horizon?.originalArrivalTime);
      const boundaryFound = Number.isFinite(boundary) && boundary > 0;
      const unread = boundaryFound ? page.items.filter(item => arrivalMillis(item.time) > boundary && item.authorId !== this.selfMri()) : [];
      chats.push({ chatId: chat.id, boundaryFound, items: unread, ...(boundaryFound ? {} : { recentMessages: page.items }) });
      items.push(...unread.map(item => ({ ...item, chatId: chat.id, chatTitle: chat.title })));
    }
    return { items, chats, remainingChats: Math.max(0, snapshot.items.length - maxChats), completeAccount: false, coverage: 'api_consumption_horizon_window', mayMarkRead: false };
  }
  async findMessage(threadId, mid) {
    mid = messageId(mid); let next;
    for (let i = 0; i < 6; i++) {
      const page = await this.page(threadId, 200, next), raw = page.raw.find(item => String(item.id) === mid);
      if (raw) return { raw, message: normalize(raw) };
      if (!page.next || page.next === next) break; next = page.next;
    }
    throw new AdapterError('message_unavailable', 'Message was not found in six recent API pages.', 404);
  }
  async prepareSend(id, input) {
    const conversation = await this.resolve(id);
    if (conversation.isMessagingDisabled || conversation.isDisabled) throw new AdapterError('api_forbidden', 'Messaging is disabled for this chat.', 403);
    const payload = messagePayload(input, this.credentials.identity.name || this.credentials.identity.email);
    return { threadId: conversation.id, payload: { ...payload, clientmessageid: String(BigInt('0x' + randomBytes(8).toString('hex')) >> 1n) } };
  }
  async send(prepared) {
    try {
      const result = await this.request('messages', this.messagePath(prepared.threadId), { method: 'POST', body: JSON.stringify(prepared.payload) });
      const id = result?.OriginalArrivalTime ?? result?.originalarrivaltime ?? result?.id;
      if (id == null) throw new Error();
      return { status: 'accepted_by_api', message: { id: String(id), clientMessageId: prepared.payload.clientmessageid }, deliveryConfirmed: false };
    } catch { throw new AdapterError('send_uncertain', 'The send was attempted. Inspect Teams before further action; reuse this idempotency key.', 409); }
  }
  async reactions(id, mid) {
    const conversation = await this.resolve(id), { message } = await this.findMessage(conversation.id, mid);
    const items = message.reactions.map(entry => ({ reaction: entry.key, count: entry.users?.length ?? 0, selected: entry.users?.some(user => user.mri === this.selfMri()) ?? false }));
    return { messageId: mid, items, available: ['like', 'heart', 'laugh', 'surprised', 'sad', 'angry'].map(reaction => ({ reaction, selected: items.some(item => item.reaction === reaction && item.selected) })), coverage: 'api_message_reactions' };
  }
  async prepareMutation(id, mid, kind, input) {
    const conversation = await this.resolve(id), found = await this.findMessage(conversation.id, mid);
    if (kind === 'reaction') return { threadId: conversation.id, mid, kind, payload: reactionInput(input.reaction, input.selected) };
    if (found.message.authorId !== this.selfMri()) throw new AdapterError('not_message_owner', 'Only your own messages can be changed.', 403);
    if (found.message.text !== input.expectedText) throw new AdapterError('message_changed', 'Message text changed; read it again before changing it.', 409);
    if (found.message.deleted) throw new AdapterError('message_deleted', 'This message was already deleted.', 409);
    if (kind === 'edit' && (found.message.mentions.length || found.raw.properties?.files && found.raw.properties.files !== '[]' || /schema\.skype\.com\/Reply/.test(found.message.richText || ''))) unsupported();
    const payload = kind === 'edit' ? { ...messagePayload(input, this.credentials.identity.name || this.credentials.identity.email), skypeeditedid: mid } : undefined;
    return { threadId: conversation.id, mid, kind, payload, expected: found.message.text, version: found.raw.version };
  }
  async mutate(prepared) {
    // Re-read immediately before a write so changed ownership/text is refused.
    // The private endpoint offers no conditional-write guarantee: a remote race remains possible.
    if (prepared.kind !== 'reaction') {
      const current = await this.findMessage(prepared.threadId, prepared.mid);
      if (current.message.authorId !== this.selfMri() || current.message.text !== prepared.expected || current.raw.version !== prepared.version) throw new AdapterError('message_changed', 'Message changed before the API write.', 409);
    }
    let method = prepared.kind === 'delete' ? 'DELETE' : 'PUT', suffix = this.messagePath(prepared.threadId) + '/' + encodeURIComponent(prepared.mid), payload = prepared.payload;
    if (prepared.kind === 'reaction') { suffix += '/properties?name=emotions'; method = payload.selected ? 'PUT' : 'DELETE'; payload = { emotions: { key: payload.reaction, value: prepared.mid } }; }
    try {
      await this.request('messages', suffix, { method, ...(payload ? { body: JSON.stringify(payload) } : {}) });
      return { status: 'accepted_by_api', operation: prepared.kind, messageId: prepared.mid };
    } catch { throw new AdapterError('mutation_uncertain', 'The change was attempted. Inspect Teams and reuse the same idempotency key.', 409); }
  }
}
