import { createHash, randomUUID } from 'node:crypto';
import { AdapterError } from './errors.js';
import { arrivalMillis, encodeConversation, messagePayload } from './teams-api.js';
import { messageId } from './input.js';

// Source-verified quote/read contracts: docs/triage-protocol.md. Never enable
// candidate writes merely because a mocked server accepts their payload.
export const triageProtocols = Object.freeze({ chatQuotes: true, channelThreads: true, channelReplies: false, readMarkers: false });
export const triageCapabilities = {
  inbox: true, replyContexts: true, batchReplies: true, chatQuotes: true,
  channelThreads: true, channelReplies: false, readAfterReply: false,
  liveValidated: ['inbox', 'replyContexts'], pendingValidation: ['chatQuotes', 'channelReplies', 'readAfterReply'],
};
const TTL = 1800_000, CAPACITY = 1000, SOURCES = 20;
const millis = message => arrivalMillis(message.time);
const chronological = messages => messages.sort((a, b) => millis(a) - millis(b) || a.id.localeCompare(b.id));
const unique = messages => chronological([...new Map(messages.map(message => [message.id, message])).values()]);
const stamp = message => createHash('sha256').update(JSON.stringify([message.id, message.text, message.richText, message.authorId, message.deleted, message.version])).digest('hex');
const escape = value => String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const stopError = error => ['api_unauthorized', 'signin_required', 'token_import_required', 'api_rate_limited', 'api_transport_failed', 'send_uncertain', 'client_disconnected'].includes(error.code);
const publicError = error => ({ code: error instanceof AdapterError ? error.code : 'adapter_error', message: error instanceof AdapterError ? error.message : 'Local Teams triage operation failed.' });
const range = (value, fallback, max, name) => {
  value ??= fallback;
  if (!Number.isInteger(value) || value < 1 || value > max) throw new AdapterError('invalid_body', `${name} must be between 1 and ${max}.`, 400);
  return value;
};
function horizon(conversation) {
  const value = conversation.userConsumptionHorizon ?? conversation.consumptionHorizon;
  const boundary = Number(value?.originalArrivalTime ?? value?.OriginalArrivalTime);
  return Number.isFinite(boundary) && boundary > 0 ? boundary : null;
}
function decorate(page) {
  const raw = new Map(page.raw.map(message => [String(message.id), message]));
  return page.items.map(message => ({ ...message, version: raw.get(message.id)?.version ?? null,
    messageType: raw.get(message.id)?.messagetype ?? raw.get(message.id)?.messageType ?? null,
    clientMessageId: raw.get(message.id)?.clientmessageid ?? raw.get(message.id)?.clientMessageId ?? null }));
}
const visible = message => !message.deleted && Boolean(message.authorId) && Number.isFinite(millis(message)) &&
  (message.messageType == null || /^(Text|RichText(?:\/.*)?)$/.test(message.messageType));

export class TeamsTriage {
  constructor(api, ledger, { protocols = triageProtocols } = {}) {
    Object.assign(this, { api, ledger, protocols });
    this.targets = new Map(); this.cursors = new Map(); this.localSends = new Map();
  }
  account() { const { tenant, oid } = this.api.credentials.identity; return `${tenant}:${oid}`; }
  prune() {
    for (const map of [this.targets, this.cursors]) for (const [key, entry] of map) if (entry.expires <= Date.now() || entry.account !== this.account()) map.delete(key);
    for (const [key, expires] of this.localSends) if (expires <= Date.now()) this.localSends.delete(key);
  }
  put(map, value) {
    this.prune();
    // Evict oldest handles under a bounded memory budget, never accept an evicted one.
    if (map.size >= CAPACITY) map.delete(map.keys().next().value);
    const key = randomUUID(); map.set(key, { ...value, account: this.account(), expires: Date.now() + TTL }); return key;
  }
  target(key) {
    this.prune(); const value = this.targets.get(key);
    if (!value) throw new AdapterError('invalid_reply_target', 'Reply target expired or belongs to another account. Fetch fresh context.', 400);
    return value;
  }
  mentionsMe(message) {
    const self = this.api.selfMri(), oid = this.api.credentials.identity.oid;
    return message.mentions.some(mention => [mention?.mri, mention?.id, mention?.mentioned?.user?.id].some(id => id === self || id === oid));
  }
  async snapshot(source, signal, limit = 50) {
    const threadId = source.parentMessageId ? `${source.threadId};messageid=${messageId(source.parentMessageId)}` : source.threadId;
    const page = await this.api.page(threadId, limit, undefined, signal);
    const pageMessages = decorate(page), messages = unique([...(source.root ? [source.root] : []), ...pageMessages]), boundary = horizon(source.conversation);
    // An old root added to a recent reply window is not proof that all unread
    // replies were traversed. Only the actual server page can cover the horizon.
    const valid = pageMessages.filter(visible);
    const complete = boundary !== null && messages.every(message => Number.isFinite(millis(message))) &&
      (!page.next || valid.length > 0 && millis(valid[0]) <= boundary);
    return { ...source, messages, boundary, complete, hasMore: Boolean(page.next) };
  }
  async channelSnapshots(source, signal, traverse = false) {
    if (!this.protocols.channelThreads) throw new AdapterError('api_operation_unsupported', 'Channel thread reads are not verified.', 501);
    const roots = source.rootQueue?.length ? null : await this.api.page(source.threadId, traverse ? 20 : 50, source.rootNext, signal);
    const seen = new Set(source.rootSeen ?? []);
    const parents = roots ? decorate(roots).filter(message => !message.parentMessageId && visible(message) && !seen.has(message.id)) : source.rootQueue;
    const selected = parents.slice(-20);
    const items = [];
    // Sources run in parallel; keep thread traversal serial within each channel
    // so the aggregate never exceeds four concurrent message reads.
    for (const root of selected) {
      signal?.throwIfAborted();
      try { items.push(await this.snapshot({ ...source, parentMessageId: root.id, root }, signal)); }
      catch (error) { if (stopError(error) || signal?.aborted) throw error; items.push({ error: publicError(error) }); }
    }
    let nextSource;
    const rootNext = roots?.next ?? (roots ? null : source.rootNext), links = [...(source.rootLinks ?? [])];
    if (roots && source.rootNext) links.push(source.rootNext);
    const rootQueue = parents.slice(0, -20);
    for (const parent of selected) seen.add(parent.id);
    const repeated = rootNext && (rootNext === source.rootNext && Boolean(roots) || links.includes(rootNext));
    if (traverse && (rootQueue.length || rootNext && !repeated) && seen.size < 100_000) {
      nextSource = { ...source, rootQueue, rootNext, rootSeen: [...seen], rootLinks: links };
    }
    return { items, nextSource, complete: !rootQueue.length && !rootNext,
      ...(repeated || seen.size >= 100_000 ? { issues: [{ conversationId: source.conversationId, reason: 'channel_root_traversal_stopped' }] } : {}) };
  }
  async parallel(values, fn, signal) {
    const results = new Array(values.length); let index = 0, failure;
    const controller = new AbortController();
    const shared = signal ? AbortSignal.any([signal, controller.signal]) : controller.signal;
    await Promise.allSettled(Array.from({ length: Math.min(4, values.length) }, async () => {
      while (index < values.length) {
        if (shared.aborted) { failure ??= shared.reason; break; } const next = index++;
        try { results[next] = await fn(values[next], shared); }
        catch (error) {
          if (stopError(error) || shared.aborted) { failure ??= error; controller.abort(error); break; }
          results[next] = { error: publicError(error) };
        }
      }
    }));
    if (failure) throw failure;
    return results;
  }
  view(snapshot, contextLimit, selectedId) {
    const valid = snapshot.messages.filter(visible), self = this.api.selfMri();
    const unread = snapshot.boundary === null ? [] : valid.filter(message => millis(message) > snapshot.boundary && message.authorId !== self);
    const context = unique([...valid.slice(-contextLimit), ...(selectedId ? valid.filter(message => message.id === selectedId) : []), ...(snapshot.root ? [snapshot.root] : [])]);
    const shown = unique([...unread, ...context]);
    const latestOwn = Math.max(0, ...valid.filter(message => message.authorId === self).map(millis));
    const incoming = valid.filter(message => message.authorId !== self);
    const targets = new Map(), activity = snapshot.messages.map(item => ({ id: item.id, stamp: stamp(item) }));
    for (const message of shown) {
      const target = this.put(this.targets, { snapshot, selected: message, seen: shown, activity });
      targets.set(message.id, target);
    }
    const output = message => {
      const { version, clientMessageId, ...publicMessage } = message;
      return { ...publicMessage, isOwn: message.authorId === self, mentionedMe: this.mentionsMe(message),
        repliedSince: latestOwn > millis(message), replyTarget: targets.get(message.id) };
    };
    const latestIncoming = unread.at(-1) ?? incoming.at(-1);
    const defaultMessage = selectedId ? shown.find(message => message.id === selectedId) : latestIncoming;
    return { kind: snapshot.kind, conversationId: snapshot.conversationId, title: snapshot.title,
      ...(snapshot.parentMessageId ? { parentMessageId: snapshot.parentMessageId, teamTitle: snapshot.teamTitle } : {}),
      unread: unread.map(output), context: context.filter(message => !unread.some(item => item.id === message.id)).map(output),
      ...(defaultMessage && targets.has(defaultMessage.id) ? { replyTarget: targets.get(defaultMessage.id) } : {}),
      canReply: snapshot.kind === 'chat' ? this.protocols.chatQuotes : this.protocols.channelReplies,
      signals: { mentionedMe: unread.some(message => this.mentionsMe(message)), directChat: snapshot.kind === 'chat' && snapshot.conversation.isOneOnOne === true,
        oldestUnreadAt: unread[0]?.time ?? null },
      coverage: { unreadWindowComplete: snapshot.complete, hasMoreHistory: snapshot.hasMore, readStateKnown: snapshot.boundary !== null }, mayMarkRead: false };
  }
  async inbox({ limit = 10, contextLimit = 5, cursor, signal } = {}) {
    limit = range(limit, 10, 20, 'limit'); contextLimit = range(contextLimit, 5, 20, 'contextLimit');
    signal = signal ? AbortSignal.any([signal, AbortSignal.timeout(60_000)]) : AbortSignal.timeout(60_000);
    this.prune(); let state;
    if (cursor) {
      state = this.cursors.get(cursor);
      if (!state || state.limit !== limit || state.contextLimit !== contextLimit) throw new AdapterError('invalid_cursor', 'Triage cursor expired or parameters changed.', 400);
      if (state.result) return state.result;
    } else {
      const data = await this.api.directory(true, signal), sources = [], issues = [];
      for (const chat of data.chats) if (!chat.isConversationDeleted && typeof chat.isRead !== 'boolean') issues.push({ kind: 'chat', conversationId: encodeConversation('chat', chat.id), reason: 'unknown_unread_flag' });
      for (const chat of data.chats.filter(chat => !chat.isConversationDeleted && chat.isRead === false)) {
        sources.push({ kind: 'chat', threadId: chat.id, conversationId: encodeConversation('chat', chat.id), title: chat.title || 'Teams chat', conversation: chat });
      }
      for (const team of data.teams) for (const channel of team.channels ?? []) {
        if (channel.isDeleted || channel.isMember === false) continue;
        const isRead = channel.isMessageRead ?? channel.isRead;
        if (isRead === true) continue;
        if (isRead !== false) { issues.push({ kind: 'channel', conversationId: encodeConversation('channel', channel.id), reason: 'unknown_unread_flag' }); continue; }
        sources.push({ kind: 'channel', threadId: channel.id, conversationId: encodeConversation('channel', channel.id), title: channel.displayName || 'Teams channel', teamTitle: team.displayName, conversation: channel });
      }
      state = { limit, contextLimit, sources, groups: [], issues, discoveryPartial: Boolean(data.metadata?.isPartialData || data.metadata?.hasMoreChats || data.metadata?.forwardSyncToken) };
    }
    // A source scan is bounded; priority is explicit within discovered coverage.
    let nextState = { ...state, sources: [...state.sources], groups: [...state.groups], issues: [...state.issues] };
    if (!nextState.groups.length && nextState.sources.length) {
      const sources = nextState.sources.splice(0, SOURCES);
      const scanned = await this.parallel(sources, async (source, shared) => {
        if (horizon(source.conversation) === null) return { issues: [{ conversationId: source.conversationId, reason: 'unknown_consumption_horizon' }], items: [] };
        if (source.kind === 'chat') return { items: [await this.snapshot(source, shared)], complete: true };
        return this.channelSnapshots(source, shared, true);
      }, signal);
      scanned.forEach((result, index) => {
        if (result.error) nextState.issues.push({ conversationId: sources[index].conversationId, ...result.error });
        nextState.issues.push(...(result.issues ?? []));
        if (result.nextSource) nextState.sources.push(result.nextSource);
        else if (result.complete === false) nextState.issues.push({ conversationId: sources[index].conversationId, reason: 'channel_roots_window' });
        for (const snapshot of result.items ?? []) {
          if (snapshot.error) { nextState.issues.push({ conversationId: sources[index].conversationId, ...snapshot.error }); continue; }
          if (snapshot.messages.some(message => visible(message) && millis(message) > snapshot.boundary && message.authorId !== this.api.selfMri())) nextState.groups.push(snapshot);
        }
      });
      const priority = snapshot => snapshot.messages.some(message => visible(message) && millis(message) > snapshot.boundary && message.authorId !== this.api.selfMri() && this.mentionsMe(message)) ? 0 : snapshot.kind === 'chat' && snapshot.conversation.isOneOnOne ? 1 : 2;
      const oldest = snapshot => Math.min(...snapshot.messages.filter(message => visible(message) && millis(message) > snapshot.boundary && message.authorId !== this.api.selfMri()).map(millis));
      nextState.groups.sort((a, b) => priority(a) - priority(b) || oldest(a) - oldest(b) || a.threadId.localeCompare(b.threadId) || String(a.parentMessageId).localeCompare(String(b.parentMessageId)));
    }
    const items = nextState.groups.splice(0, limit).map(snapshot => this.view(snapshot, contextLimit));
    const hasMore = Boolean(nextState.groups.length || nextState.sources.length);
    const nextCursor = hasMore ? this.put(this.cursors, { ...nextState, result: undefined }) : null;
    const result = { items, hasMore, nextCursor, remainingSources: nextState.sources.length, completeAccount: false,
      coverage: { priorityScope: 'scanned_sources', discoveryPartial: nextState.discoveryPartial, issues: nextState.issues }, mayMarkRead: false };
    if (cursor) state.result = result;
    return result;
  }
  async contexts(input, signal) {
    if (!input || Object.keys(input).some(key => !['targets', 'contextLimit'].includes(key)) || !Array.isArray(input.targets) || input.targets.length < 1 || input.targets.length > 10 || input.targets.some(target => typeof target !== 'string')) throw new AdapterError('invalid_body', 'Provide 1–10 reply targets and optional contextLimit.', 400);
    const limit = range(input.contextLimit, 20, 50, 'contextLimit');
    signal = signal ? AbortSignal.any([signal, AbortSignal.timeout(60_000)]) : AbortSignal.timeout(60_000);
    await this.api.directory(true, signal);
    const items = await this.parallel(input.targets, async (target, shared) => {
      const previous = this.target(target), source = await this.freshSource(previous.snapshot, shared);
      const snapshot = await this.snapshot(source, shared, Math.max(50, limit));
      if (!snapshot.messages.some(message => message.id === previous.selected.id && visible(message))) throw new AdapterError('message_unavailable', 'Selected message is no longer in the current context window.', 404);
      return { target, ...this.view(snapshot, limit, previous.selected.id) };
    }, signal);
    return { items: items.map((item, index) => ({ target: input.targets[index], ...item })), mayMarkRead: false };
  }
  async freshSource(source, signal) {
    const conversation = await this.api.resolve(source.conversationId, source.kind);
    let root = source.root;
    if (source.parentMessageId) {
      const found = await this.api.findMessage(source.threadId, source.parentMessageId, signal);
      root = { ...found.message, version: found.raw.version ?? null, clientMessageId: found.raw.clientmessageid ?? found.raw.clientMessageId ?? null,
        messageType: found.raw.messagetype ?? found.raw.messageType ?? null };
      if (!visible(root)) throw new AdapterError('message_unavailable', 'The channel thread root is unavailable.', 404);
    }
    return { ...source, conversation, ...(root ? { root } : {}) };
  }
  validateReplies(input) {
    if (!input || Object.keys(input).some(key => key !== 'replies') || !Array.isArray(input.replies) || input.replies.length < 1 || input.replies.length > 10) throw new AdapterError('invalid_body', 'Provide 1–10 explicit replies.', 400);
    const keys = new Set();
    return input.replies.map(reply => {
      if (!reply || Array.isArray(reply) || Object.keys(reply).some(key => !['target', 'text', 'content', 'idempotencyKey'].includes(key)) || typeof reply.target !== 'string' || !/^[a-f0-9-]{36}$/.test(reply.target) || typeof reply.idempotencyKey !== 'string' || !/^[A-Za-z0-9._:-]{8,128}$/.test(reply.idempotencyKey) || (reply.text === undefined) === (reply.content === undefined)) throw new AdapterError('invalid_body', 'Each reply requires target, exactly one of text/content, and an 8–128 character idempotencyKey.', 400);
      // Validate all content before the first batch write; canonicalize fingerprints.
      messagePayload(reply, '');
      if (keys.has(reply.idempotencyKey)) throw new AdapterError('invalid_body', 'Use a different idempotencyKey for each reply in a batch.', 400);
      keys.add(reply.idempotencyKey);
      return { target: reply.target, ...(reply.content === undefined ? { text: reply.text } : { content: reply.content }), idempotencyKey: reply.idempotencyKey };
    });
  }
  async prepareReply(reply, signal) {
    const target = this.target(reply.target), previous = target.snapshot;
    if (!(previous.kind === 'chat' ? this.protocols.chatQuotes : this.protocols.channelReplies)) throw new AdapterError('api_operation_unsupported', 'Native replies for this conversation kind await protocol validation.', 501);
    await this.api.directory(true, signal);
    const source = await this.freshSource(previous, signal);
    if (source.conversation.isMessagingDisabled || source.conversation.isDisabled || source.conversation.isArchived) throw new AdapterError('api_forbidden', 'Messaging is disabled for this conversation.', 403);
    const current = await this.snapshot(source, signal);
    const selected = current.messages.find(message => message.id === target.selected.id);
    const known = new Map(target.activity.map(message => [message.id, message]));
    const locallySent = message => message.authorId === this.api.selfMri() && [message.id, message.clientMessageId].some(id => this.localSends.has(`${this.account()}:${id}`));
    if (!selected || !visible(selected) || stamp(selected) !== stamp(target.selected) || current.messages.some(message => !locallySent(message) &&
      (!known.has(message.id) || known.get(message.id).stamp !== stamp(message)))) throw new AdapterError('context_changed', 'Context or the selected message changed. Fetch fresh reply context before sending.', 409);
    if (!selected.authorId) throw new AdapterError('api_schema_changed', 'The selected sender identity is unavailable.', 502);
    const prepared = await this.api.prepareSendForTriage(source, reply);
    if (source.kind === 'chat') {
      const preview = selected.text.length > 200 ? `${selected.text.slice(0, 197)}...` : selected.text;
      const quote = `<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="${escape(selected.id)}"><strong itemprop="mri" itemid="${escape(selected.authorId)}">${escape(selected.author || selected.authorId)}</strong><p itemprop="preview">${escape(preview)}</p></blockquote>`;
      const content = prepared.payload.messagetype === 'Text' ? `<p>${escape(prepared.payload.content).replace(/\n/g, '<br>')}</p>` : prepared.payload.content;
      prepared.payload.content = quote + content; prepared.payload.messagetype = 'RichText/Html';
    }
    const cutoff = unique(target.seen.filter(visible)).at(-1);
    prepared.readIntent = { account: this.account(), source: { kind: source.kind, threadId: source.threadId, conversationId: source.conversationId, parentMessageId: source.parentMessageId },
      complete: previous.complete, cutoff: cutoff && { id: cutoff.id, time: millis(cutoff), clientMessageId: cutoff.clientMessageId },
      seen: target.seen.filter(visible).map(message => ({ id: message.id, time: millis(message), stamp: stamp(message) })) };
    return prepared;
  }
  async replies(input, { signal, checkCancelled = () => {} } = {}) {
    const replies = this.validateReplies(input), items = []; let stopped = false;
    signal = signal ? AbortSignal.any([signal, AbortSignal.timeout(90_000)]) : AbortSignal.timeout(90_000);
    for (const reply of replies) {
      if (stopped) { items.push({ target: reply.target, idempotencyKey: reply.idempotencyKey, send: { status: 'not_attempted', reason: 'batch_stopped' }, read: { status: 'not_attempted' } }); continue; }
      let dispatching = false;
      try {
        checkCancelled(); signal.throwIfAborted();
        const result = await this.ledger.run(reply.idempotencyKey, ['triage-reply', this.account(), reply], () => this.prepareReply(reply, signal), async prepared => {
          checkCancelled(); signal.throwIfAborted(); dispatching = true;
          return { ...await this.api.send(prepared, signal), _triageRead: prepared.readIntent };
        }, 'send_uncertain');
        const { _triageRead, ...send } = result;
        for (const id of [send.message.id, send.message.clientMessageId]) {
          if (this.localSends.size >= CAPACITY) this.localSends.delete(this.localSends.keys().next().value);
          this.localSends.set(`${this.account()}:${id}`, Date.now() + TTL);
        }
        items.push({ target: reply.target, idempotencyKey: reply.idempotencyKey, send, read: { status: 'pending' } });
      } catch (error) {
        if (dispatching && error.code !== 'send_uncertain') error = new AdapterError('send_uncertain', 'Reply may have been accepted. Reuse this key and inspect Teams before further action.', 409);
        items.push({ target: reply.target, idempotencyKey: reply.idempotencyKey, send: { status: error.code === 'send_uncertain' ? 'uncertain' : 'failed', error: publicError(error) }, read: { status: 'not_attempted' } });
        stopped ||= stopError(error) || signal.aborted || !(error instanceof AdapterError);
      }
    }
    // The accepted result was fsynced first. Read failures cannot hide that result.
    let reads;
    try { reads = await this.ledger.updateTriageReads(this.account(), entries => this.updateReads(entries, signal)); }
    catch { reads = null; }
    for (const item of items) if (item.send.status === 'accepted_by_api') item.read = reads?.get(createHash('sha256').update(item.idempotencyKey).digest('hex')) ?? { status: 'failed', reason: 'read_update_persistence_failed' };
    return { items, stopped, deliveryConfirmed: false };
  }
  async updateReads(entries, signal) {
    const results = new Map(), groups = new Map();
    for (const entry of entries) {
      if (['marked_read', 'already_read'].includes(entry.read?.status)) { results.set(entry.key, entry.read); continue; }
      if (!this.protocols.readMarkers) { results.set(entry.key, { status: 'deferred', reason: 'protocol_unverified' }); continue; }
      const id = entry.intent.source.conversationId;
      if (!groups.has(id)) groups.set(id, []); groups.get(id).push(entry);
    }
    // Completed grants remain evidence of handled channel threads for later updates.
    for (const [id, pending] of groups) {
      const handled = entries.filter(entry => entry.intent.source.conversationId === id), source = pending[0].intent.source;
      try {
        signal?.throwIfAborted(); await this.api.directory(true, signal);
        const fresh = await this.freshSource(source, signal), boundary = horizon(fresh.conversation);
        const newest = pending.filter(entry => entry.intent.cutoff).sort((a, b) => b.intent.cutoff.time - a.intent.cutoff.time)[0];
        const desired = newest?.intent.cutoff;
        let outcome;
        if (!desired || boundary === null) outcome = { status: 'deferred', reason: 'unknown_consumption_horizon' };
        else if (boundary >= desired.time) outcome = { status: 'already_read' };
        else if (!newest.intent.complete) outcome = { status: 'deferred', reason: 'incomplete_unread_window' };
        else {
          const scan = source.kind === 'channel' ? await this.channelSnapshots(fresh, signal) : { items: [await this.snapshot(fresh, signal)], complete: true };
          const seen = new Set(handled.flatMap(entry => entry.intent.seen.map(message => message.stamp)));
          const unsafe = !scan.complete || scan.items.some(snapshot => snapshot.error || !snapshot.complete || snapshot.messages.some(message => visible(message) &&
            millis(message) > boundary && millis(message) <= desired.time && message.authorId !== this.api.selfMri() && !seen.has(stamp(message))));
          if (unsafe) outcome = { status: 'deferred', reason: source.kind === 'channel' ? 'unhandled_channel_threads_or_incomplete_coverage' : 'unseen_or_changed_messages' };
          else if (typeof desired.clientMessageId !== 'string' || !/^\d+$/.test(desired.clientMessageId) || desired.id !== String(desired.time)) outcome = { status: 'deferred', reason: 'read_marker_fields_unavailable' };
          else {
            // Refresh immediately before a marker write, including after a slow
            // channel scan. The candidate API offers no atomic compare-and-set.
            await this.api.directory(true, signal);
            const latestBoundary = horizon(await this.api.resolve(id, source.kind));
            if (latestBoundary === null || latestBoundary < boundary) outcome = { status: 'deferred', reason: 'read_state_changed' };
            else if (latestBoundary >= desired.time) outcome = { status: 'already_read' };
            else {
              await this.api.request('messages', `users/ME/conversations/${encodeURIComponent(source.threadId)}/properties?name=consumptionhorizon`, {
                method: 'PUT', body: JSON.stringify({ consumptionhorizon: `${desired.id};${Date.now()};${desired.clientMessageId}` }), signal });
              this.api.directoryCache = null; outcome = { status: 'marked_read', throughMessageId: desired.id };
            }
          }
        }
        for (const entry of pending) results.set(entry.key, outcome);
      } catch (error) { for (const entry of pending) results.set(entry.key, { status: 'failed', error: publicError(error) }); }
    }
    return results;
  }
}
