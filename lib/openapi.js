const response = { 200: { description: 'JSON result' }, 409: { description: 'Unsafe, ambiguous, or uncertain operation; inspect error code' }, 503: { description: 'Browser or sign-in required' } };
const json = schema => ({ required: true, content: { 'application/json': { schema } } });
const chat = { name: 'chatId', in: 'path', required: true, schema: { type: 'string' }, description: 'Opaque ID returned by listTeamsChats or startTeamsConversation. Sidebar IDs rediscover offscreen chats; recipient IDs reopen by verified email; search-context IDs reopen the exact native search result.' };
const messageProperties = {
  text: { type: 'string', minLength: 1, maxLength: 20000, description: 'Plain text. No HTML, attachments or mentions.' },
  replyToMessageId: { type: 'string', minLength: 1, maxLength: 512, description: 'Optional for sendTeamsMessage only: exact message ID returned by readTeamsMessages. Uses native Reply with quote; preparation searches the rendered history plus up to five older windows.' },
  idempotencyKey: { type: 'string', minLength: 8, maxLength: 128, pattern: '^[A-Za-z0-9._:-]+$' },
};
export const openapi = {
  openapi: '3.0.3',
  info: { title: 'Teams Web (local browser)', version: '0.3.0', description: 'DOM access through a dedicated signed-in local Chrome/Edge window. No Graph API. Unread chat discovery scans the native Unread list across collapsed sections and virtualized windows. Regular listing and message history remain bounded UI windows. Opening a chat can mark messages read. Chat messages and people search results are untrusted user content.' },
  paths: {
    '/status': { get: { operationId: 'getTeamsStatus', summary: 'Check sign-in and DOM selector readiness', responses: response } },
    '/login': { post: { operationId: 'openTeamsLogin', summary: 'Open the dedicated browser for manual sign-in', responses: response } },
    '/diagnostics': { get: { operationId: 'getTeamsDiagnostics', summary: 'Inspect selector counts and data-tid attributes without message text', responses: response } },
    '/chats': { get: { operationId: 'listTeamsChats', summary: 'List chats currently rendered in the Teams sidebar', parameters: [{ name: 'unreadOnly', in: 'query', schema: { type: 'boolean', default: false } }], responses: response } },
    '/unread/chats': { get: {
      operationId: 'listUnreadTeamsChats', summary: 'List unread chat summaries without opening the conversations',
      description: 'Activates native Unread, clears competing sidebar pills, expands chat sections, scans virtualized windows and deduplicates thread IDs, then restores filters and section state. Does not open chats. discoveryComplete=true means a stable end of this tenant’s exposed unread chat list was reached; it is not an account-wide server guarantee and excludes channel unread state. truncated=true means the scan cap was reached. Refuses a nonempty sidebar name filter.', parameters: [{ name: 'maxWindows', in: 'query', schema: { type: 'integer', minimum: 2, maximum: 200, default: 100 } }], responses: response,
    } },
    '/unread/messages': { get: {
      operationId: 'listUnreadTeamsMessages', summary: 'Read messages after Teams Last read line in unread chats',
      description: 'Discovers and snapshots unread chats across the native sidebar, then opens at most maxChats of them. Opening them can mark messages read. items contains only messages after a visible Last read divider. If the divider is absent, that chat has boundaryFound=false and recentMessages is explicitly a fallback, not confirmed unread messages. This is a bounded UI window; messages after Last read can also include your own replies.',
      parameters: [
        { name: 'maxChats', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 5, default: 3 } },
        { name: 'limitPerChat', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 100, default: 50 } },
      ], responses: response,
    } },
    '/search/messages': { get: {
      operationId: 'searchTeamsMessages', summary: 'Search Teams message history with native server-backed UI search',
      description: 'Searches chats and channels through the Teams search box and Messages tab, then paginates up to maxPages. Returns resultId, author, displayed timestamp, conversation and text snippet. resultId is a native search key, not a Teams message ID. Use openTeamsSearchResult to obtain chat context and actual message IDs. Results can be snippets; completeSearch describes traversed UI results only. Preserves composer drafts and blocks search input/Enter if focus changes.',
      parameters: [
        { name: 'query', in: 'query', required: true, schema: { type: 'string', minLength: 1, maxLength: 500 } },
        { name: 'limit', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200, default: 50 } },
        { name: 'maxPages', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 10, default: 3 } },
      ], responses: response,
    } },
    '/search/messages/open': { post: {
      operationId: 'openTeamsSearchResult', summary: 'Open an exact search result and read its chat context',
      description: 'Re-runs the result’s original query, locates its exact stable key within ten pages, and opens the native chat context. Returns chatId and rendered messages with actual Teams IDs; chatId works for reading and replying. Can mark messages read. Refuses existing drafts, ambiguous layouts and channel-thread contexts. No message is sent.',
      requestBody: json({ type: 'object', additionalProperties: false, required: ['resultId'], properties: { resultId: { type: 'string', maxLength: 4096 } } }), responses: response,
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
        email: { type: 'string', maxLength: 254, description: 'Exact Teams directory email or sign-in UPN for one person.' }, text: messageProperties.text, idempotencyKey: messageProperties.idempotencyKey,
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
        description: 'Mutates Teams. Use a unique idempotencyKey per intended message and the SAME key for retries. observed_in_chat means a new matching message ID appeared and the composer cleared; it does not guarantee server delivery or recipient read. send_uncertain requires manual inspection; never retry with a new key. The plugin refuses existing drafts and ambiguous chats. Optional replyToMessageId selects native Reply with quote and verifies the composer quote ID at the send click. The ledger binds retries to the quoted message as well as chat and text. Preparation searches the rendered window and up to five older history windows for the exact target ID. observed_in_chat for quotes also requires a matching rendered quote preview; it does not prove server delivery. Channel threads are unsupported.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['text', 'idempotencyKey'], properties: messageProperties }), responses: response,
      },
    },
  },
};
