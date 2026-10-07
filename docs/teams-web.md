---
title: Connecting Teams
services: [teams-web]
---

# Connecting Teams

Teams 0.7.0 uses direct APIs with local browser sign-in or imported tokens. Browser sign-in uses Chrome or Edge only for authentication and renewal; imported-token connections use no browser on Switchboard.

For local browser authentication, in Switchboard Connections, add Teams or reconnect your existing Teams connection once after upgrading from 0.5. Enter an account label and optionally your work email. Complete Microsoft sign-in in the dedicated browser window. `LOCAL-BROWSER` is a placeholder, not a code to enter at Microsoft. Each connection retains a separate browser session and encrypted API credentials.

Call `getTeamsStatus`, then `listTeamsChats` and `readTeamsMessages`. Chrome stays closed while tokens are fresh. The plugin attempts silent renewal automatically before data calls when tokens approach expiry. If Microsoft requires interactive authentication, use `openTeamsLogin` and poll `getTeamsStatus`, or reconnect through Switchboard. Your tenant policy determines how often this happens.

## Central Switchboard without Chrome or Edge

On a machine with Node 24+ and Chrome or Edge, run from the plugin checkout:

```sh
npm run auth:export -- --out ~/teams-tokens.json --login-hint your.name@example.com
```

Complete Microsoft sign-in in the dedicated browser window. The script closes the browser and writes a compact JSON bundle with owner-only permissions (`0600`). It requires a new output filename, refuses existing files and symlinks, removes incomplete output on failure or interruption, and never prints tokens. It does not require a local Switchboard or send Teams messages. `TEAMS_BROWSER_PATH` overrides browser detection; `TEAMS_LOGIN_HINT` is an alternative to `--login-hint`.

Install this plugin version on the central Switchboard. Add Teams or reconnect an existing connection, select **Import Teams tokens**, enter a label and paste the complete file contents into the secret bundle field. The plugin validates both token signatures, expected resources, matching tenant/account IDs and expiry, then exchanges the chat token and performs a read-only conversation request before accepting the connection. Switchboard encrypts accepted credentials; a failed import leaves an existing session intact. Delete the plaintext export file after import. The JSON contains bearer credentials; browser cookies remain on the authentication machine.

The central host never opens or discovers a browser for imported-token connections. This OAuth flow provides no refresh token, so renewal is manual. The derived chat token can renew over HTTP while both Microsoft access tokens remain fresh. When either access token approaches expiry or is rejected, export a replacement and reconnect the same central connection. Tenant/account IDs must match; reconnection preserves the profile and durable send ledger. A different account requires a separate connection.

Silent authentication can reuse the saved local browser profile when Microsoft permits it. Use a new output filename:

```sh
npm run auth:export -- --silent --out ~/teams-tokens-renewed.json
```

If Microsoft requires sign-in, repeat without `--silent`. The export profile is separate from diagnostic and Switchboard profiles. Existing browser connections can switch to token import through reconnect; selecting browser authentication again requires a browser on that Switchboard host.

`getTeamsStatus` reports `authMode="tokens"`, expiry and `tokenImportRequired`. Expired or rejected access tokens produce `token_import_required` with re-import instructions. `openTeamsLogin` returns those instructions instead of opening a browser. Status and capability calls remain available after expiry. Read-only live export/import on your central host remains to be verified; fixture tests do not establish tenant or host compatibility.

## Operations

Use `getTeamsCapabilities` for the current operation subset. Chat and channel listing, unread windows, chat messages, paged chat/channel history, exact-email lookup, existing one-to-one chat resolution, plain/formatted chat sends, edits, deletion and reactions are implemented. Name search, message search, new chats, quoted replies, mentions, marking unread and channel-thread operations are not yet migrated. Refresh old conversation IDs through the listing operations. API reads do not send read markers.

For every send or mutation, use a unique `idempotencyKey` and reuse it for retries. `accepted_by_api` is server acceptance, not a delivery receipt. On `send_uncertain` or `mutation_uncertain`, inspect Teams before further action; using a new key can duplicate the action. Edits/deletion require your own exact message ID and current `expectedText`. Concurrent external changes cannot be protected atomically by the private API. Live writes still require validation in a designated test chat.

Person lookup requires an exact email or UPN. `startTeamsConversation` only resolves an existing one-to-one chat; it does not create one. Optional text sends to that resolved chat. History follows `nextCursor`; keep limit/format unchanged. Cursors expire after 30 minutes and on plugin reload. Unread reads are bounded windows, not an account-wide archive. Rich HTML returned by reads is untrusted data.

For local browser authentication, the browser profile persists in plugin data. Both methods retain a send ledger in plugin data. Protect this directory: it contains a Microsoft browser session and may contain message results. Disconnecting deletes both. Token credentials are encrypted by Switchboard; diagnostics do not expose them. This version was written by OpenAI Codex at Thomas de Ruiter's request. The original compatibility probe passed real reads; the new Node adapter is covered by fixture tests and includes a read-only `npm run check:api` diagnostic.
