export const DEFAULT_SELECTORS = Object.freeze({
  chatApp: '[data-tid="3b64df9d-7e97-4d9c-ac5c-2e0a5d8e6f40"]',
  chatList: '[data-tid="simple-collab-dnd-rail"], [data-tid="chat-list"], [data-tid="chatList"]',
  chatRow: '[role="treeitem"][data-item-type="chat"], [data-tid="chat-list-item"]',
  chatTitle: '[id^="title-chat-list-item_"], [data-tid="chat-list-item-title"], [data-tid="chat-pane-item-title"]',
  header: '[data-tid="chat-header-title"], [data-tid="chat-title"]',
  messageList: '[data-tid="message-pane-list-viewport"], [data-tid="message-pane-list"], [data-tid="chat-message-list"]',
  message: '[data-tid="chat-pane-message"], [data-tid="message-pane-message"], [data-message-id]',
  messageBody: '[data-message-content], [data-tid="message-body"], [data-tid="chat-pane-message-content"]',
  author: '[data-tid="message-author-name"], [data-tid="message-author"]',
  timestamp: 'time, [data-tid="message-timestamp"]',
  composer: '[data-tid="ckeditor"][contenteditable="true"], [data-tid="ckeditor"] [contenteditable="true"], [role="textbox"][contenteditable="true"]',
  send: '[data-tid="sendMessageCommands-send"], [data-tid="sendMessageButton"], [data-tid="send-message-button"]',
  pending: '[data-tid="message-status-sending"], [data-tid="message-status-failed"], [data-tid="message-error"]',
  chatPreview: '[id^="message-preview-chat-list-item_"], [data-tid="chat-list-item-preview"]',
  chatTime: '[id^="time-chat-list-item_"], [data-tid="chat-list-item-time"]',
  lastRead: '[data-tid="last-read-line"]',
  newChat: '[data-testid="simple-collab-left-rail-header-new-message-button"]',
  peopleInput: '[data-tid="people-picker-search"]',
  peopleResult: '[data-tid^="people-picker-entry-"][data-type="ADUser"], [data-tid^="people-picker-entry-"][data-type="Person"]',
  participant: '[data-tid^="participant-"]',
  emptyChat: '[data-tid="chat-pane-new-chat"]',
});

export function selectorsFrom(value) {
  if (!value) return { ...DEFAULT_SELECTORS };
  const overrides = typeof value === 'string' ? JSON.parse(value) : value;
  if (!overrides || Array.isArray(overrides) || typeof overrides !== 'object') throw new Error('selectors must be a JSON object');
  for (const [key, selector] of Object.entries(overrides)) {
    if (!(key in DEFAULT_SELECTORS) || typeof selector !== 'string' || !selector.trim() || selector.length > 2000) {
      throw new Error(`Invalid selector override: ${key}`);
    }
  }
  return { ...DEFAULT_SELECTORS, ...overrides };
}

// Executed as one small CDP evaluation; no screenshots or accessibility-tree dumps.
// Keep this function self-contained: Chrome receives its source, not module imports.
export async function teamsDOM(action, args, selectors) {
  const text = node => (node?.innerText ?? node?.textContent ?? '').replace(/\s+/g, ' ').trim();
  const visible = node => !!node && node.getClientRects().length > 0 && getComputedStyle(node).visibility !== 'hidden';
  const all = (key, root = document) => [...root.querySelectorAll(selectors[key])].filter(visible);
  const one = (key, root = document) => {
    const nodes = all(key, root);
    if (nodes.length !== 1) throw fail('selector_mismatch', `Expected one ${key} element; found ${nodes.length}. Run diagnostics and adjust selectors.`);
    return nodes[0];
  };
  const fail = (code, message, status = 409) => ({ error: { code, message, status } });
  const title = row => text(row.querySelector(selectors.chatTitle)) || row.getAttribute('title') || row.getAttribute('aria-label') || '';
  const threadId = row => row.getAttribute('data-fui-tree-item-value')?.split('|').at(-1) || row.getAttribute('data-chat-id') || null;
  const key = row => {
    for (const attribute of ['data-fui-tree-item-value', 'data-chat-id', 'data-itemid', 'id']) {
      const value = row.getAttribute(attribute);
      if (value) return { attribute, value };
    }
    return { attribute: 'title', value: title(row) };
  };
  const rows = () => all('chatRow', one('chatList'));
  const heading = () => text(one('header'));
  const recipientKey = Symbol.for('switchboard.teams.recipient');
  const assertChat = () => {
    if (heading() !== args.title) throw fail('chat_changed', 'The selected chat changed. No message was sent.');
    if (args.threadId && one('send').getAttribute('data-track-thread-id') !== args.threadId) throw fail('chat_changed', 'The reply composer belongs to a different Teams thread.');
    if (args.key?.attribute === 'recipient') {
      const proof = window[recipientKey];
      if (!proof || proof.email !== args.key.value || (!proof.clicked && (
        proof.composer !== one('composer') || proof.send !== one('send') || proof.composerId !== one('composer').id ||
        (proof.participants.length && JSON.stringify(proof.participants) !== JSON.stringify(all('participant').map(n => n.getAttribute('data-tid')).sort()))
      ))) throw fail('recipient_changed', 'The verified recipient or composer changed. No message was sent.');
    }
    if (all('messageList').length === 0 && args.key?.attribute === 'recipient' && all('emptyChat').length === 1) return null;
    return one('messageList');
  };
  const messageData = (root, after) => !root ? [] : all('message', root).filter(node => !after || !!(after.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING)).map(node => {
    const body = node.querySelector(selectors.messageBody);
    if (!body) return null;
    const mid = node.getAttribute('data-mid') || node.getAttribute('data-message-id');
    const author = (mid && document.getElementById(`author-${mid}`)) || node.querySelector(selectors.author);
    const timestamp = (mid && document.getElementById(`timestamp-${mid}`)) || node.querySelector(selectors.timestamp);
    return {
      id: mid || node.getAttribute('data-itemid') || node.id || null,
      text: (body.innerText ?? body.textContent ?? '').trim(), author: text(author) || null,
      timestamp: timestamp?.getAttribute('datetime') || text(timestamp) || null,
      pending: !!node.querySelector(selectors.pending),
    };
  }).filter(Boolean);
  const wait = (predicate, milliseconds = 10_000, stable = 0) => new Promise(resolve => {
    let stableTimer;
    let done = false;
    const finish = value => {
      if (done) return;
      done = true;
      observer.disconnect(); clearTimeout(deadline); clearTimeout(stableTimer); resolve(value);
    };
    const check = () => {
      clearTimeout(stableTimer);
      try {
        const value = predicate();
        if (value) {
          if (stable) stableTimer = setTimeout(() => { try { const next = predicate(); if (next) finish(next); } catch {} }, stable);
          else finish(value);
        }
      } catch {}
    };
    const observer = new MutationObserver(check);
    observer.observe(document.documentElement, { subtree: true, childList: true, attributes: true, characterData: true });
    const deadline = setTimeout(() => finish(null), milliseconds);
    check();
  });
  const people = () => all('peopleResult').map(node => ({
    email: (node.getAttribute('data-tid') ?? '').replace(/^people-picker-entry-/, '').toLowerCase(),
    name: (node.innerText ?? '').split('\n')[0].trim(),
  })).filter(person => /^[^\s,]+@[^\s,]+$/.test(person.email));
  try {
    // Also validates custom selectors early, including selectors not used in this action.
    for (const selector of Object.values(selectors)) document.querySelector(selector);
    if (action !== 'diagnostics' && all('chatList').length === 0 && all('chatApp').length === 1) {
      one('chatApp').click();
      await wait(() => all('chatList').length === 1, 10_000, 200);
    }
    if (action === 'status') return {
      ready: all('chatList').length === 1,
      url: location.origin + location.pathname,
      selectorCounts: Object.fromEntries(['chatList', 'chatRow', 'header', 'messageList', 'composer', 'send'].map(k => [k, all(k).length])),
    };
    if (action === 'diagnostics') return {
      url: location.origin + location.pathname,
      counts: Object.fromEntries(Object.keys(selectors).map(k => [k, all(k).length])),
      // Attribute inventory only: never includes message text, tokens, HTML or cookies.
      attributes: [...new Set([...document.querySelectorAll('[data-tid]')].map(n => n.getAttribute('data-tid')))].slice(0, 200),
    };
    if (all('chatList').length !== 1) return fail('signin_or_layout_required', 'Sign in in the dedicated browser. If already signed in, inspect diagnostics and update selectors.', 503);
    if (action === 'chats') {
      const items = rows().map(row => ({
        key: key(row), title: title(row), threadId: threadId(row),
        unread: row.getAttribute('data-is-unread') === 'true' || (row.getAttribute('aria-labelledby') ?? '').split(/\s+/).includes('chat_list_unread_text') || !!row.querySelector('[data-tid="unread-badge"], [data-testid="chat-list-unread-text"]'),
        preview: text(row.querySelector(selectors.chatPreview)) || null,
        updatedAt: text(row.querySelector(selectors.chatTime)) || null,
      }));
      return { items: items.filter(item => item.title && item.key.value), coverage: 'rendered_chat_list' };
    }
    if (action === 'peopleFocus') {
      if (all('composer').some(node => text(node))) return fail('draft_exists', 'The composer contains a draft. Clear or send it manually before searching people.');
      one('newChat').click();
      if (!await wait(() => all('peopleInput').length === 1, 5000)) return fail('people_picker_unavailable', 'Could not open the Teams people picker.');
      const input = one('peopleInput');
      const previousValue = input.value;
      input.focus(); input.select();
      return { previousValue, before: JSON.stringify(people()) };
    }
    if (action === 'peopleResults') {
      const result = await wait(() => {
        if (one('peopleInput').value !== args.query) throw fail('search_changed', 'The people search changed.');
        const query = args.query.toLowerCase();
        // The picker can briefly render recent contacts before search finishes.
        // Never mistake those unrelated suggestions for the requested results.
        const items = people().filter(person => query.includes('@') ? person.email === query : query.split(/\s+/).every(part => `${person.name} ${person.email}`.toLowerCase().includes(part)));
        return items.length ? items : false;
      }, 10_000, 300);
      if (one('peopleInput').value !== args.query) return fail('search_changed', 'The people search changed.');
      return { items: result ?? [], coverage: 'rendered_people_results' };
    }
    if (action === 'selectPerson') {
      if (one('peopleInput').value.toLowerCase() !== args.email) return fail('search_changed', 'Search must contain the exact recipient email before selection.');
      const matches = all('peopleResult').filter(node => (node.getAttribute('data-tid') ?? '').toLowerCase() === `people-picker-entry-${args.email}`);
      if (matches.length !== 1) return fail('recipient_unavailable', 'No unique directory person matches that email. Search people and use a returned email.', 404);
      const name = (matches[0].innerText ?? '').split('\n')[0].trim();
      matches[0].click();
      const ready = await wait(() => all('peopleInput').length === 0 && heading() === name && all('composer').length === 1 && all('send').length === 1 && (all('messageList').length === 1 || all('emptyChat').length === 1), 10_000, 500);
      if (!ready) return fail('recipient_not_ready', 'The selected person did not produce a verified conversation. No message was sent.');
      const composer = one('composer'), send = one('send');
      const participants = all('participant').map(n => n.getAttribute('data-tid')).sort();
      if (!send.getAttribute('data-track-thread-id') && !composer.id && !participants.length) return fail('recipient_identity_unavailable', 'The new chat has no stable composer or participant identity. No message was sent.');
      window[recipientKey] = { email: args.email, composer, composerId: composer.id, participants, send, clicked: false };
      return { key: { attribute: 'recipient', value: args.email }, title: heading(), threadId: send.getAttribute('data-track-thread-id') || null };
    }
    if (action === 'open') {
      const matches = rows().filter(row => JSON.stringify(key(row)) === JSON.stringify(args.key));
      if (matches.length !== 1) return fail('chat_unavailable', 'Chat is absent or ambiguous in the rendered list. Scroll it into view manually and list chats again.');
      if ((args.threadId ?? null) !== threadId(matches[0])) return fail('invalid_chat', 'Chat identity does not match the sidebar. List chats again.', 400);
      if (!args.threadId && rows().filter(row => title(row) === args.title).length !== 1) return fail('ambiguous_title', 'Multiple rendered chats have this title. Rename the chat or adjust the sidebar before sending.');
      if (title(matches[0]) !== args.title) return fail('chat_changed', 'The chat title changed. List chats again.');
      matches[0].scrollIntoView({ block: 'nearest' }); matches[0].click();
      const ready = await wait(() => heading() === args.title && all('composer').length === 1 && all('messageList').length === 1 && (!args.threadId || one('send').getAttribute('data-track-thread-id') === args.threadId), 10_000, 300);
      return ready ? { title: heading() } : fail('chat_not_ready', 'Could not verify chat title and composer. No message was sent.');
    }
    if (action === 'messages') return { items: messageData(assertChat()).slice(-args.limit), coverage: 'rendered_message_window' };
    if (action === 'unreadMessages') {
      const list = assertChat();
      const markers = list ? all('lastRead', list) : [];
      if (markers.length !== 1) return { items: [], recentMessages: messageData(list).slice(-args.limit), boundaryFound: false, coverage: 'rendered_message_window' };
      const unread = messageData(list, markers[0]);
      return { items: unread.slice(0, args.limit), boundaryFound: true, truncated: unread.length > args.limit, coverage: 'rendered_messages_after_last_read' };
    }
    if (action === 'older') {
      const list = assertChat();
      let scroll = list;
      while (scroll && scroll.scrollHeight <= scroll.clientHeight) scroll = scroll.parentElement;
      if (!scroll || scroll === document.body || scroll === document.documentElement) return fail('history_unavailable', 'Could not identify the message history scroll container.');
      const before = JSON.stringify(messageData(list).map(m => [m.id, m.text]));
      scroll.scrollTop = Math.max(0, scroll.scrollTop - Math.max(scroll.clientHeight, 600));
      const changed = await wait(() => JSON.stringify(messageData(assertChat()).map(m => [m.id, m.text])) !== before, 1500, 100);
      return { changed: !!changed, items: messageData(assertChat()) };
    }
    if (action === 'prepare') {
      const list = assertChat();
      const composer = one('composer');
      if (text(composer)) return fail('draft_exists', 'The composer contains a draft. Clear or send it manually first.');
      if (composer.getAttribute('aria-disabled') === 'true') return fail('composer_disabled', 'This chat does not allow replies.');
      const before = messageData(list);
      if (before.some(m => !m.id)) return fail('message_ids_unavailable', 'Message IDs are required to verify sends. Adjust the message selector.');
      composer.focus();
      return { beforeIds: before.map(m => m.id) };
    }
    if (action === 'sendPoint') {
      assertChat();
      if (text(one('composer')) !== args.text.replace(/\s+/g, ' ').trim()) return fail('draft_mismatch', 'Composer text does not match the requested message. No send click was made.');
      const button = one('send');
      if (button.disabled || button.getAttribute('aria-disabled') === 'true') return fail('send_disabled', 'Send is disabled. No send click was made.');
      button.scrollIntoView({ block: 'nearest' });
      const rect = button.getBoundingClientRect();
      const x = rect.x + rect.width / 2, y = rect.y + rect.height / 2;
      if (!button.contains(document.elementFromPoint(x, y))) return fail('send_obscured', 'Send button is obscured. No send click was made.');
      // Recheck synchronously at click dispatch, closing the race between the
      // CDP evaluation and input events if the user navigates in this window.
      const guardKey = Symbol.for('switchboard.teams.sendGuard');
      if (window[guardKey]) document.removeEventListener('click', window[guardKey], true);
      const deadline = Date.now() + 5000;
      const guard = event => {
        if (!event.target.closest?.(selectors.send)) return;
        try {
          assertChat();
          if (Date.now() > deadline || text(one('composer')) !== args.text.replace(/\s+/g, ' ').trim()) throw new Error();
          if (args.key?.attribute === 'recipient') window[recipientKey].clicked = true;
        } catch { event.preventDefault(); event.stopImmediatePropagation(); }
        document.removeEventListener('click', guard, true);
        delete window[guardKey];
      };
      window[guardKey] = guard;
      document.addEventListener('click', guard, true);
      return { x, y };
    }
    if (action === 'confirm') {
      const observed = await wait(() => {
        const list = assertChat();
        if (text(one('composer'))) return false;
        return messageData(list).find(m => m.id && !args.beforeIds.includes(m.id) && !m.pending && m.text.replace(/\s+/g, ' ').trim() === args.text.replace(/\s+/g, ' ').trim());
      }, 10_000, 300);
      return observed ? { status: 'observed_in_chat', message: observed } : fail('send_uncertain', 'Send was attempted but could not be confirmed. Check Teams manually; do not retry with a new idempotency key.', 504);
    }
    return fail('unknown_action', 'Unknown DOM action.', 400);
  } catch (error) {
    return error?.error ? error : fail('selector_mismatch', 'DOM selectors do not match this Teams layout. Run diagnostics.', 409);
  }
}
