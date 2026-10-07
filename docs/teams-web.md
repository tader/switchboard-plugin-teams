---
title: Connecting Teams Web
services: [teams-web]
---

# Connecting Teams Web

This plugin was coded by **OpenAI Codex, an AI coding agent**, at Thomas de Ruiter's request. Its implementation, tests, and documentation were AI-written; independent human code review and real Teams send validation remain outstanding.

Install it using **Plugins → Install from GitHub** with `tader/switchboard-plugin-teams`, or `tader/switchboard-plugin-teams@v0.3.0` to pin this release. Existing default-branch installations can use **Check for updates**.

For a private repository, the Switchboard administrator must configure `SWITCHBOARD_GITHUB_TOKEN` with a token that can read this repository. Connecting a GitHub account in Switchboard does not configure the plugin installer's credentials.

Run this plugin on the Mac where Chrome or Edge can open a window, either on your main Switchboard or on a local satellite. In **Connections**, select **Teams Web**, enter an account label, and connect. Sign in manually in the dedicated browser window opened by the plugin. Switchboard polls until the Teams sidebar is ready. `LOCAL-BROWSER` is a placeholder shown by the connect dialog; you do not enter it into Microsoft.

Keep this browser window open. When your session expires, call `openTeamsLogin` and sign in again in that window. Each connection has a separate browser profile. Avoid navigating or typing in the dedicated window during plugin calls.

Call `listTeamsChats` for chat IDs, `readTeamsMessages` to read a rendered window, and `sendTeamsMessage` for a plain-text message. Opening a chat may mark messages read. Unread discovery scans offscreen chats and collapsed chat sections; history remains bounded. Channel threads, attachments, group-chat creation, external-user federation search, and mentions are unsupported.

For unread summaries, use `listUnreadTeamsChats`, or `listTeamsChats` with `unreadOnly=true`. These use the native Unread filter, expand chat sections and scan virtualized sidebar windows without opening each conversation, then restore the prior filters. Set maxWindows (2–200, default 100) to cap discovery; discoveryComplete reports a stable end and truncated reports the cap. Clear any sidebar name filter first. This discovers exposed unread chats in the current tenant; channel unread state is excluded. Previews and times appear when Teams exposes them. For full text, use `listUnreadTeamsMessages` with optional `maxChats` (1–5, default 3) and `limitPerChat` (1–100, default 50). It opens a snapshot of unread chats, which can mark them read. Only messages after a visible **Last read** divider appear in `items`; a missing divider yields `boundaryFound=false` and a separate `recentMessages` fallback. The divider's range may include your own replies. This is not a complete account-wide unread archive.

Use `searchTeamsMessages` with `query`, optional `limit` (1–200) and `maxPages` (1–10) to search native Teams history. Results contain snippets and `resultId`; call `openTeamsSearchResult` with that resultId to get chat context, chatId and actual message IDs. Opening can mark messages read. Search can find channel messages, but opening channel-thread contexts for replies is unsupported. Search preserves drafts and guards input and Enter against focus races.

For a quoted reply, add `replyToMessageId` to `sendTeamsMessage`, using an exact message ID from reading chat context. Preparation searches the current window and up to five older windows for the exact target ID. The plugin verifies the native quote ID before sending; retries bind to that target as well as chat and text. A changed quote blocks sending. Failed preparation may leave a quote draft to inspect manually.

To start a conversation with a specific person, call `searchTeamsPeople` with their name or email and choose the correct returned email/UPN. Then call `startTeamsConversation` with `{ "email": "alice@example.com" }`. It returns a usable `chatId` without sending. Add `text` and `idempotencyKey` to send an initial message in the same call. A new chat may only be persisted by Teams after its first message. The plugin selects only a unique exact-email directory person and refuses display-name-only recipient selection or an existing manual message draft.

Sending requires `text` and `idempotencyKey`. Use a new key per intended message, and the same key for all retries. If the result is `send_uncertain`, inspect Teams manually; using a new key could duplicate the message. `observed_in_chat` confirms the message appeared in the UI; it is not a delivery receipt.

If sign-in works but operations fail with `selector_mismatch`, use `getTeamsDiagnostics` to inspect selector counts. An administrator can update DOM selector overrides in plugin settings. Diagnostics exclude message bodies and credentials. The plugin never uses screenshots.

Browser profiles and a local send ledger persist in the plugin data directory. Deleting the connection erases both. Switchboard's normal call audit still applies. The ledger contains reply results and message text, so keep the plugin data directory private.
