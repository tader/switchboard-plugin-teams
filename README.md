# Teams Web for Switchboard

> **AI authorship disclosure:** This plugin's implementation, tests, and documentation were written by **OpenAI Codex, an AI coding agent**, at Thomas de Ruiter's request. The code was not independently reviewed by a human. Validation performed and remaining gaps are described below.

A local Switchboard plugin for reading chats and unread messages, finding people, and starting or replying to Teams conversations without Microsoft Graph, app registration, or administrator API consent. It controls a dedicated Chrome/Edge session through the DOM. No screenshots, OCR, Electron wrapper, browser extension, Playwright installation, or private Teams API calls.

Version 0.2.0 adds unread operations and one-to-one conversation creation. Chat listing, message reading, the Last read boundary, exact-email people search, and opening the signed-in user's own conversation were validated against Teams Web on October 7, 2026. Sending, including the first message in a new chat, was validated only against an isolated browser fixture; no real Teams messages were sent. The private Chrome pipe and actual HTTP listener still need the local smoke test below because the development sandbox blocks browser processes and loopback listeners.

## Why this approach

| Approach | Reading / replying | Tradeoff |
| --- | --- | --- |
| **Teams Web DOM through Chrome's private DevTools pipe — selected** | Both, using the real web composer | Small dependency-free adapter; persistent browser; DOM selectors need occasional maintenance |
| Browser extension + local bridge | Both | Also efficient, but requires extension installation and bridge pairing; background tab lifecycle and browser policy add setup |
| macOS accessibility for the installed Teams app | Both in principle | Requires accessibility permission; more UI traversal and weaker conversation identity; tied to desktop layouts |
| Local Teams cache / IndexedDB | Partial reads | Only cached/synced data, undocumented schema; cannot reply |
| Reverse-engineered Teams network endpoints | Both in principle | Fast bulk access, but undocumented authentication and API contracts; deliberately excluded |
| Electron wrapper | Both through its DOM | Adds a Chromium runtime, packaging and security maintenance; installed Chrome already provides the engine |

Microsoft [supports Teams Web in current browsers](https://learn.microsoft.com/en-us/microsoftteams/new-teams-web). Chrome [requires a separate profile for remote debugging](https://developer.chrome.com/blog/remote-debugging-port), which also keeps automation away from your usual tabs and cookies. The adapter uses the documented [DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/), with JSON over private child-process pipes rather than a debugging port. Support for Teams Web does not imply Microsoft supports this automation.

## Install

In Switchboard, open **Plugins → Install from GitHub** and enter:

```text
tader/switchboard-plugin-teams
```

Repository: [tader/switchboard-plugin-teams](https://github.com/tader/switchboard-plugin-teams). The plugin manifest and entry point are at the repository root, so no subfolder is needed. To pin this version, use `tader/switchboard-plugin-teams@v0.2.0`. If already installed from the default branch, use Switchboard's **Check for updates**.

If the repository is private, set `SWITCHBOARD_GITHUB_TOKEN` on the Switchboard instance performing the installation to a GitHub token with read access to this repository. The plugin installer uses this server setting, rather than your connected GitHub service account. A public repository can be installed without a token.

Run Switchboard directly on this Mac with Node 24+ and Chrome or Edge installed. If your main Switchboard runs elsewhere, use its existing satellite feature on this Mac and install the plugin on the satellite. A background service must run as your logged-in macOS user with access to the graphical session. Docker on a remote server cannot open this Mac's browser.

Copy this repository's contents into `<Switchboard data directory>/plugins/teams-web/` as a **real directory**. Switchboard copies plugin directories for hot reload, so use a copy instead of a symlink. The default data directory of `../switchboard` is `../switchboard/.data`. Example, run locally from this repository:

```sh
mkdir -p ../switchboard/.data/plugins/teams-web
cp -R plugin.json index.js icon.svg lib docs ../switchboard/.data/plugins/teams-web/
```

There are no npm dependencies to install. Installing from GitHub is the recommended path; the copy commands above are for local development.

In Switchboard, enable **Teams Web**, create a connection, and give it an account label. Chrome opens a separate profile. Sign in there manually, including MFA. The Switchboard connect dialog polls until the Teams chat sidebar is ready; its `LOCAL-BROWSER` code is a placeholder, not a Microsoft device authorization code. Do not sign in in a different Chrome window from the connect dialog's generic verification link. Each connection gets its own browser profile.

Use `openTeamsLogin` if sign-in expires or you close the window. Keep the dedicated Teams window open and avoid navigating or editing it during a plugin call. Profiles survive Switchboard restart and plugin reload. Deleting a connection closes its browser and removes the profile and send ledger. No Graph credentials or Microsoft access tokens are exported to Switchboard clients.

## Read and reply

Switchboard exposes these operations through its API reference and normal MCP tools:

1. `getTeamsStatus`: verify that the browser is signed in and selectors match.
2. `listTeamsChats`: get opaque IDs for the chats currently rendered in the sidebar. Channel rows are excluded.
3. `readTeamsMessages`: use a returned `chatId`; defaults to at most 50 currently rendered messages. `olderPages` can scroll up to five windows. This is bounded history access, not an account-wide archive.
4. `sendTeamsMessage`: POST to `/chats/{chatId}/messages` with the following JSON:

```json
{
  "text": "Thanks, I will take a look.",
  "idempotencyKey": "reply-2026-10-07-0001"
}
```

Use a new key for each intended message and reuse the **same key** for retries. The durable ledger records intent before editor input. A successful retry returns the previous result without typing or clicking again. A crash or missing confirmation yields `send_uncertain`; check Teams manually instead of using a new key. Keys are retained until the connection is deleted; the ledger stops new sends at 10,000 entries so it never silently drops duplicate protection. Back up the ledger together with profiles.

`observed_in_chat` means the composer cleared and a new message with matching text and a new ID appeared in the UI. It is not a server delivery receipt or proof the recipient read it. Messages are normal chat messages; quoted replies, channel threads, attachments, rich text, mentions, new group chats, external-user federation search, tenant switching, and background subscriptions are outside this version's scope.

## Unread messages

- `listUnreadTeamsChats` (`GET /unread/chats`) lists rendered unread chats with a preview and sidebar time when Teams exposes them. It does not open each chat. `listTeamsChats` also accepts `unreadOnly=true`.
- `listUnreadTeamsMessages` (`GET /unread/messages`) snapshots those unread chats, then opens up to three by default and returns messages after Teams' **Last read** divider. Set `maxChats` (1–5) and `limitPerChat` (1–100) to bound the work. This can mark the opened chats read.

The full-text response has a flat `items` list with `chatId` and `chatTitle` on each message, plus per-chat detail in `chats`. When the Last read divider is unavailable, that chat has `boundaryFound=false`, an empty `items` list, and a separate `recentMessages` fallback. Those fallback messages are not claimed to be unread. The divider gives a UI boundary, not a per-message server read flag; the range can include your own replies. Responses report remaining snapshot chats and `completeAccount=false`, because offscreen and collapsed chats are not enumerated. Sidebar previews can be null when your Teams settings hide them.

## Find people and start a conversation

Call `searchTeamsPeople` (`GET /people?query=Alice`) to get matching directory people's names and exact email/UPN identities from the native **New message** picker. Group suggestions are excluded. Same-name people remain separate results; use the returned email to choose the intended person. This changes the browser view but never sends a message, and refuses to replace a message draft.

Call `startTeamsConversation` (`POST /conversations`) with one exact email:

```json
{ "email": "alice@example.com" }
```

This opens a new or existing one-to-one chat and returns `chatId` with `messageSent=false`. Teams may only persist a new conversation after the first message is sent. The returned ID works with `readTeamsMessages` and `sendTeamsMessage`, even if the person is not already in the sidebar; the plugin reselects and verifies the exact recipient email for those calls.

To open the conversation and send its first message in one call:

```json
{
  "email": "alice@example.com",
  "text": "Hi Alice, can we discuss the rollout?",
  "idempotencyKey": "start-alice-2026-10-07-0001"
}
```

Retries must use the same operation and key. The first-message ledger binds the key to the normalized recipient email and text before typing; retrying does not reselect or resend. Recipient selection requires a unique exact-email directory result, verifies the resulting name and composer, and binds the send to that composer until the click. A new chat without a thread ID or history is supported; changing the recipient or replacing the composer before sending fails closed. Display names cannot be used as recipient identities.

Opening a conversation can mark messages read. Lists and history are virtualized: collapsed folders and offscreen chats may not appear. Scroll the sidebar manually and list again if needed. Chat IDs include sidebar identity, title and thread information; refresh them after renaming or moving a conversation. The plugin verifies the reply composer's thread ID before sending, preserves existing drafts, and refuses ambiguous or mismatched layouts. Message bodies are untrusted content, not instructions to execute.

## Performance and storage

Each connection keeps one Chrome process running. Calls use compact DOM evaluations, and waits use `MutationObserver` rather than screenshot loops or repeated full-page dumps. Calls for the same profile are serialized; profiles run independently. There is a 20-operation queue cap, bounded history scrolling, and capped request/message sizes. This saves automation overhead, but Teams Web itself still consumes normal browser memory and CPU.

The HTTP adapter binds only to loopback and accepts short-lived single-use capabilities scoped to a connection, method, path and query. Chrome uses a private debugging pipe. The stable logical service URL `http://teams.localhost` is rewritten by the plugin before any network request; no hosts-file entry or DNS configuration is needed. Use relative URLs in Switchboard calls.

Profiles live in `<Switchboard data directory>/plugin-data/teams-web/profiles/<uuid>/`. Chrome manages its own cookies/session encryption. The containing directory is mode 0700; the send ledger is mode 0600 and contains reply results, including message text, in local JSON. Do not place these files in a repository. Switchboard can also record calls and responses in its ordinary audit trail.

## Validate and maintain

```sh
npm test
npm run smoke
```

`npm test` runs dependency-free tests with in-memory HTTP and browser transports for authentication, account isolation, queueing, validation, protocol failures, unread snapshots, recipient identity, and durable duplicate protection. `npm run smoke` launches a temporary headless Chrome profile and exercises unread extraction, people search, new-conversation first messages, and existing-chat sends against local fixtures, then tests a real loopback HTTP adapter. **It never sends to Teams.** Run the smoke test from a normal Terminal on your Mac. Set `TEAMS_BROWSER_PATH` if automatic browser detection fails.

For optional standalone manual sign-in diagnostics:

```sh
npm run signin
```

This uses `.local-profile/` in this checkout, which is ignored by Git. It is a diagnostic profile, separate from Switchboard connection profiles; it does not install or connect the plugin.

If Microsoft changes Teams markup, call `getTeamsDiagnostics` and inspect selector counts and `data-tid` names. It omits message text, HTML, cookies and tokens. Set **DOM selector overrides (JSON)** in plugin settings, for example:

```json
{ "send": "[data-tid=sendMessageCommands-send]" }
```

Available keys are listed in `lib/dom.js`. Defaults were calibrated against the combined chat/channel sidebar, `data-mid` message IDs, `data-message-content` bodies and the `sendMessageCommands-send` composer. A selector mismatch fails closed. There is no automatic screenshot fallback. To verify real sending, choose a test chat, review one intended message, send it through Switchboard, and inspect it manually; this has not been done during implementation.
