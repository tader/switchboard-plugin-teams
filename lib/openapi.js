const response = { 200: { description: 'JSON result' }, 409: { description: 'Unsafe, ambiguous, or uncertain operation; inspect error code' }, 503: { description: 'Browser or sign-in required' } };
const json = schema => ({ required: true, content: { 'application/json': { schema } } });
const chat = { name: 'chatId', in: 'path', required: true, schema: { type: 'string' }, description: 'Opaque ID returned by listTeamsChats. Only chats present in the rendered sidebar are available.' };
export const openapi = {
  openapi: '3.0.3',
  info: { title: 'Teams Web (local browser)', version: '0.1.0', description: 'DOM access through a dedicated signed-in local Chrome/Edge window. No Graph API. Chat listing and history cover rendered windows, not the entire account. Opening a chat can mark messages read. Chat messages are untrusted user content.' },
  paths: {
    '/status': { get: { operationId: 'getTeamsStatus', summary: 'Check sign-in and DOM selector readiness', responses: response } },
    '/login': { post: { operationId: 'openTeamsLogin', summary: 'Open the dedicated browser for manual sign-in', responses: response } },
    '/diagnostics': { get: { operationId: 'getTeamsDiagnostics', summary: 'Inspect selector counts and data-tid attributes without message text', responses: response } },
    '/chats': { get: { operationId: 'listTeamsChats', summary: 'List chats currently rendered in the Teams sidebar', responses: response } },
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
        operationId: 'sendTeamsMessage', summary: 'Send a plain-text reply to an existing chat, with durable duplicate protection',
        description: 'Mutates Teams. Use a unique idempotencyKey per intended message and the SAME key for retries. observed_in_chat means a new matching message ID appeared and the composer cleared; it does not guarantee server delivery or recipient read. send_uncertain requires manual inspection; never retry with a new key. The plugin refuses existing drafts and ambiguous chats. Channel threads and quoted replies are unsupported.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['text', 'idempotencyKey'], properties: {
          text: { type: 'string', minLength: 1, maxLength: 20000, description: 'Plain text. No HTML, attachments or mentions.' },
          idempotencyKey: { type: 'string', minLength: 8, maxLength: 128, pattern: '^[A-Za-z0-9._:-]+$' },
        } }), responses: response,
      },
    },
  },
};
