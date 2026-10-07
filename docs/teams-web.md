---
title: Connecting Teams
services: [teams-web]
---

# Connecting Teams

Teams 0.6.0 uses direct APIs. Chrome or Edge is used only to sign in and renew authentication; it closes afterward.

In Switchboard Connections, add Teams or reconnect your existing Teams connection once after upgrading from 0.5. Enter an account label and optionally your work email. Complete Microsoft sign-in in the dedicated browser window. `LOCAL-BROWSER` is a placeholder, not a code to enter at Microsoft. Each connection retains a separate browser session and encrypted API credentials.

Call `getTeamsStatus`, then `listTeamsChats` and `readTeamsMessages`. Chrome stays closed while tokens are fresh. The plugin attempts silent renewal automatically before data calls when tokens approach expiry. If Microsoft requires interactive authentication, use `openTeamsLogin` and poll `getTeamsStatus`, or reconnect through Switchboard. Your tenant policy determines how often this happens.

Use `getTeamsCapabilities` for the current operation subset. Chat and channel listing, unread windows, chat messages, paged chat/channel history, exact-email lookup, existing one-to-one chat resolution, plain/formatted chat sends, edits, deletion and reactions are implemented. Name search, message search, new chats, quoted replies, mentions, marking unread and channel-thread operations are not yet migrated. Refresh old conversation IDs through the listing operations. API reads do not send read markers.

For every send or mutation, use a unique `idempotencyKey` and reuse it for retries. `accepted_by_api` is server acceptance, not a delivery receipt. On `send_uncertain` or `mutation_uncertain`, inspect Teams before further action; using a new key can duplicate the action. Edits/deletion require your own exact message ID and current `expectedText`. Concurrent external changes cannot be protected atomically by the private API. Live writes still require validation in a designated test chat.

Person lookup requires an exact email or UPN. `startTeamsConversation` only resolves an existing one-to-one chat; it does not create one. Optional text sends to that resolved chat. History follows `nextCursor`; keep limit/format unchanged. Cursors expire after 30 minutes and on plugin reload. Unread reads are bounded windows, not an account-wide archive. Rich HTML returned by reads is untrusted data.

The browser profile and send ledger persist in plugin data. Protect this directory: it contains a Microsoft browser session and may contain message results. Disconnecting deletes both. Token credentials are encrypted by Switchboard; diagnostics do not expose them. This version was written by OpenAI Codex at Thomas de Ruiter's request. The original compatibility probe passed real reads; the new Node adapter is covered by fixture tests and includes a read-only `npm run check:api` diagnostic.
