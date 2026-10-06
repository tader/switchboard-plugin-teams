---
title: Connecting Teams Web
services: [teams-web]
---

# Connecting Teams Web

This plugin was coded by **OpenAI Codex, an AI coding agent**, at Thomas de Ruiter's request. Its implementation, tests, and documentation were AI-written; independent human code review and real Teams send validation remain outstanding.

Install it using **Plugins → Install from GitHub** with `tader/switchboard-plugin-teams`, or `tader/switchboard-plugin-teams@v0.1.0` to pin the initial release.

For a private repository, the Switchboard administrator must configure `SWITCHBOARD_GITHUB_TOKEN` with a token that can read this repository. Connecting a GitHub account in Switchboard does not configure the plugin installer's credentials.

Run this plugin on the Mac where Chrome or Edge can open a window, either on your main Switchboard or on a local satellite. In **Connections**, select **Teams Web**, enter an account label, and connect. Sign in manually in the dedicated browser window opened by the plugin. Switchboard polls until the Teams sidebar is ready. `LOCAL-BROWSER` is a placeholder shown by the connect dialog; you do not enter it into Microsoft.

Keep this browser window open. When your session expires, call `openTeamsLogin` and sign in again in that window. Each connection has a separate browser profile. Avoid navigating or typing in the dedicated window during plugin calls.

Call `listTeamsChats` for chat IDs, `readTeamsMessages` to read a rendered window, and `sendTeamsMessage` for a plain-text reply to an existing chat. Opening a chat may mark messages read. Offscreen chats, collapsed folders, and older messages are not automatically enumerated. Channel threads, quoted replies, attachments, new chats and mentions are unsupported.

Sending requires `text` and `idempotencyKey`. Use a new key per intended message, and the same key for all retries. If the result is `send_uncertain`, inspect Teams manually; using a new key could duplicate the message. `observed_in_chat` confirms the message appeared in the UI; it is not a delivery receipt.

If sign-in works but operations fail with `selector_mismatch`, use `getTeamsDiagnostics` to inspect selector counts. An administrator can update DOM selector overrides in plugin settings. Diagnostics exclude message bodies and credentials. The plugin never uses screenshots.

Browser profiles and a local send ledger persist in the plugin data directory. Deleting the connection erases both. Switchboard's normal call audit still applies. The ledger contains reply results and message text, so keep the plugin data directory private.
