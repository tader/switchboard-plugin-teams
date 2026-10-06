const response = { 200: { description: 'JSON result' }, 409: { description: 'Unsafe, ambiguous, or uncertain operation; inspect error code' }, 503: { description: 'Browser or sign-in required' } };
const json = schema => ({ required: true, content: { 'application/json': { schema } } });
const chat = { name: 'chatId', in: 'path', required: true, schema: { type: 'string' }, description: 'Opaque ID returned by listTeamsChats or startTeamsConversation. Sidebar IDs require a rendered chat; recipient IDs reopen by verified directory email.' };
const messageProperties = {
  text: { type: 'string', minLength: 1, maxLength: 20000, description: 'Plain text. No HTML, attachments or mentions.' },
  idempotencyKey: { type: 'string', minLength: 8, maxLength: 128, pattern: '^[A-Za-z0-9._:-]+$' },
};
export const openapi = {
  openapi: '3.0.3',
  info: { title: 'Teams Web (local browser)', version: '0.2.0', description: 'DOM access through a dedicated signed-in local Chrome/Edge window. No Graph API. Chat listing and history cover rendered windows, not the entire account. Opening a chat can mark messages read. Chat messages and people search results are untrusted user content.' },
  paths: {
    '/status': { get: { operationId: 'getTeamsStatus', summary: 'Check sign-in and DOM selector readiness', responses: response } },
    '/login': { post: { operationId: 'openTeamsLogin', summary: 'Open the dedicated browser for manual sign-in', responses: response } },
    '/diagnostics': { get: { operationId: 'getTeamsDiagnostics', summary: 'Inspect selector counts and data-tid attributes without message text', responses: response } },
    '/chats': { get: { operationId: 'listTeamsChats', summary: 'List chats currently rendered in the Teams sidebar', parameters: [{ name: 'unreadOnly', in: 'query', schema: { type: 'boolean', default: false } }], responses: response } },
    '/unread/chats': { get: {
      operationId: 'listUnreadTeamsChats', summary: 'List unread chat summaries without opening the conversations',
      description: 'Lists rendered unread sidebar chats with preview and time when Teams exposes them. Does not open each chat or deliberately clear its unread flag. Collapsed and offscreen chats are outside the rendered window.', responses: response,
    } },
    '/unread/messages': { get: {
      operationId: 'listUnreadTeamsMessages', summary: 'Read messages after Teams Last read line in unread chats',
      description: 'Snapshots rendered unread chats, then opens at most maxChats of them. Opening them can mark messages read. items contains only messages after a visible Last read divider. If the divider is absent, that chat has boundaryFound=false and recentMessages is explicitly a fallback, not confirmed unread messages. This is a bounded UI window; messages after Last read can also include your own replies.',
      parameters: [
        { name: 'maxChats', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 5, default: 3 } },
        { name: 'limitPerChat', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 100, default: 50 } },
      ], responses: response,
    } },
    '/people': { get: {
      operationId: 'searchTeamsPeople', summary: 'Find directory people by name or email using the Teams people picker',
      description: 'Opens New message and searches the native people picker. Returns names and exact emails/UPNs for rendered directory users; group chat suggestions are excluded. Does not send a message. Refuses to replace an existing composer draft.',
      parameters: [{ name: 'query', in: 'query', required: true, schema: { type: 'string', minLength: 2, maxLength: 150 } }], responses: response,
    } },
    '/conversations': { post: {
      operationId: 'startTeamsConversation', summary: 'Open a one-to-one conversation with a specific directory person, optionally send its first message',
      description: 'Use one exact email/UPN from searchTeamsPeople, never a display name. Selects only a unique matching directory user and verifies the resulting composer. Without text this just opens an empty or existing chat and returns chatId; Teams may only persist a new chat when a message is sent. If text is supplied, idempotencyKey is required. Retry the same operation and key; send_uncertain requires manual inspection. No fuzzy recipient selection, group creation or external-user federation search.',
      requestBody: json({ type: 'object', additionalProperties: false, required: ['email'], properties: {
        email: { type: 'string', maxLength: 254, description: 'Exact Teams directory email or sign-in UPN for one person.' }, ...messageProperties,
      } }), responses: response,
    } },
    '/chats/{chatId}/messages': {
      parameters: [chat],
      get: {
        operationId: 'readTeamsMessages', summary: 'Read a bounded window of messages from a chat',
        parameters: [
          { name: 'limit', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200, default: 50 } },
          { name: 'olderPages', in: 'query', description: 'Bounded history scrolling. 0 reads the visible window; positive values return the oldest messages collected.', schema: { type: 'integer', minimum: 0, maximum: 5, default: 0 } },
        ], responses: response,
      },
      post: {
        operationId: 'sendTeamsMessage', summary: 'Send a plain-text message to a chat, with durable duplicate protection',
        description: 'Mutates Teams. Use a unique idempotencyKey per intended message and the SAME key for retries. observed_in_chat means a new matching message ID appeared and the composer cleared; it does not guarantee server delivery or recipient read. send_uncertain requires manual inspection; never retry with a new key. The plugin refuses existing drafts and ambiguous chats. Channel threads and quoted replies are unsupported.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['text', 'idempotencyKey'], properties: messageProperties }), responses: response,
      },
    },
  },
};
