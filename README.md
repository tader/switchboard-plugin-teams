# Teams API plugin for Switchboard

Version 0.7.0 supports direct Teams private API calls with local browser sign-in or imported tokens. For browser sign-in, Chrome or Edge opens only for authentication and renewal, then closes. A central Switchboard can import tokens from another machine and run without Chrome or Edge. No Go, Electron, Microsoft Graph registration, npm dependencies, or persistent Teams tab is required. Node 24+ is required.

The service ID remains `teams-web`, and the existing browser authentication method remains `browser`, so existing Switchboard connections can be reconnected in place. The previous DOM adapter is retained in `legacy-index.js` for regression testing; production does not invoke it.

## Install locally

From this checkout:

```sh
npm test
npm run install:local -- ~/switchboard/.data
```

The installer stages code outside the plugin watcher, replaces `.data/plugins/teams-web`, and saves the previous code under `.data/plugin-backups/`. It does not change connection records, encrypted credentials, or `.data/plugin-data/teams-web`. Reload or restart Switchboard after installing. This branch has not been published as a GitHub release.

In Switchboard, reconnect the existing Teams connection once (or add a new Teams connection). Choose Microsoft sign-in through local Chrome, or Import Teams tokens for a browser-free host (see below). For browser sign-in, enter your work email as the optional Microsoft work email hint. Complete sign-in in the dedicated browser window; it closes automatically when authentication finishes. `LOCAL-BROWSER` is only a placeholder in the connection dialog.

Old DOM-only connections report `signin_required` until reconnected. Refresh chat/channel IDs by listing them again; exact legacy thread IDs can migrate after a membership check, but UI search, person and history cursors cannot. Existing send ledger entries remain preserved; a reused key with changed request semantics is rejected.

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

## Test without sending messages

```sh
npm test
npm run smoke
TEAMS_LOGIN_HINT=Thomas.De-Ruiter@dsm-firmenich.com npm run check:api
TEAMS_LOGIN_HINT=Thomas.De-Ruiter@dsm-firmenich.com npm run check:api -- --silent
```

`npm test` runs unit and adapter tests, including the retained DOM regression tests. `npm run smoke` runs the new API adapter and authentication tests with fixtures; neither command contacts Teams or sends real messages.

`check:api` performs real Microsoft authentication, closes Chrome, exchanges the chat token, then reads your profile, conversation list and five recent messages from one chat. It prints only PASS markers and token expiry, never tokens or message content. The second command checks whether the saved session can obtain tokens without interactive sign-in. Both use a separate diagnostic profile at `.local-profile/api/auth-browser`; they do not create a Switchboard connection or persist API tokens. Set `TEAMS_BROWSER_PATH` to override browser detection. A successful silent check demonstrates current session reuse, not a guarantee about future company sign-in policy.

After installing, run `getTeamsStatus`, `getTeamsCapabilities`, `listTeamsChats` and `readTeamsMessages` in Switchboard. Chrome should remain closed for data calls with fresh tokens. Verify an intended write manually in a designated test chat before depending on write operations. No real messages were sent during this implementation.

## Authentication lifetime

Tokens have individual expiry timestamps. Before a data call, the plugin renews when less than two minutes remain. An expired chat-service token can be exchanged over HTTP while its source token remains valid. In browser mode, expired OAuth access tokens require briefly launching headless Chrome with the same profile and `prompt=none`; no refresh token is obtained by this flow. The browser closes after each attempt. Imported-token connections instead require manual replacement of Microsoft access tokens.

There is no fixed manual-login interval: Microsoft session expiry, MFA, revocation and tenant policy decide when silent renewal stops working. `getTeamsStatus` reports `tokenExpiresAt`, readiness and sign-in errors. Use `openTeamsLogin` and poll status, or reconnect in Switchboard, when interactive sign-in is required. Renewal occurs on demand, not on a background schedule. Control/status calls remain available when tokens expire.

## Operations and migration limits

| Operation | API version |
| --- | --- |
| Chat/channel listing, recent chat messages | Supported |
| Unread chats/messages | Supported, bounded conversation snapshot and consumption horizon |
| Chat/channel history, NDJSON | Supported, server backward links and bounded cursors |
| Person lookup | Exact email/UPN only |
| Start conversation operation | Resolves an existing one-to-one chat; cannot create a new one |
| Chat sends and edits | Plain text or structured formatting and links |
| Delete own chat message; add/remove reaction | Supported |
| Message search, quoted replies, mentions, mark unread | Not yet migrated |
| Channel thread reads/writes, attachments, new conversations | Not yet migrated |

Unsupported operations are removed from the advertised OpenAPI schema; saved calls fail explicitly. `getTeamsCapabilities` reports the supported subset. Conversation data can be partial, and retained history is not a guaranteed complete archive. API reads do not send read markers.

## Writes and history

Every write requires an `idempotencyKey`. Use one new key per intended action and reuse it for retries. The durable ledger records intent before the HTTP write; there are no automatic write retries. `accepted_by_api` means server acceptance, not delivery or read confirmation. If `send_uncertain` or `mutation_uncertain` occurs, inspect Teams before further action; retrying with a new key can duplicate the action.

Edits and deletion require `expectedText`, exact message IDs and ownership. The adapter rechecks text and version immediately before writing, but the private endpoint does not provide an atomic conditional-write guarantee. Edits refuse existing mentions, quoted replies and attachments. Mutation target lookup searches at most six pages of 200 messages. Reaction IDs are the supported common Teams reactions.

For formatted sends/edits, use `content` instead of `text`, for example:

```json
{"content":[{"type":"paragraph","runs":[{"text":"Hello","marks":["bold"]}]}],"idempotencyKey":"example-unique-key"}
```

Supported blocks: paragraph, quote, bulletedList and numberedList. Marks: bold, italic, underline, strike and code; links must use HTTP(S). Structured quotes are formatting, not replies to another message. No arbitrary HTML or mentions are accepted. Message HTML returned by reads is untrusted data.

History cursors are bound to the connection, conversation, page size and format. They expire after 30 minutes and reset on plugin reload. Reusing a cursor repeats its page. Keep `limit` and `format` unchanged while following `nextCursor`; pages proceed from newer to older history with chronological records inside each page. Capacity is 1,000 cursor records and 100,000 observed IDs per traversal. `format=ndjson` adds an export string to each JSON response.

Unread message reads select messages newer than the conversation consumption horizon and exclude your own messages. If the horizon is absent, `items` is empty and a separate `recentMessages` fallback is returned. These are bounded recent windows; `completeAccount=false` remains explicit.

## Storage and validation

Switchboard encrypts API tokens as connection credentials for both authentication methods. The plugin holds decrypted tokens only in memory. For browser authentication, Chrome retains Microsoft session cookies in its separate per-connection `auth-browser` profile under the plugin data directory. Protect that directory like a logged-in browser profile. Disconnecting a connection deletes its profile and ledger; installation preserves them. Switchboard's normal call audit still applies. The ledger can contain request results and message text.

OAuth state, nonce, ID-token signature, issuer, audience, tenant and account binding are checked. HTTP calls use fixed Teams service hosts and the auth-service regional chat host; history links cannot change host or conversation. The loopback adapter exposes short-lived single-use invocation capabilities, never Microsoft tokens. Diagnostics omit tokens and message contents.

The earlier Go compatibility probe was live-validated for OAuth, profile, conversation and recent-message reads. This Node implementation has automated fixture coverage for authentication, renewal, portable token export/import, restart recovery, account isolation, request routing, pagination, token redaction, write payloads and duplicate protection. The Node live check, silent renewal under your tenant policy, central token import, and live writes still need validation. Teams private endpoints can change independently of this plugin. See [API sources](THIRD_PARTY.md).

This implementation and its tests were written by OpenAI Codex at Thomas de Ruiter's request. Independent code review remains outstanding.
