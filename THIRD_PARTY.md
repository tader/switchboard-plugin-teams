# API contract references

This implementation uses Node built-ins and has no third-party runtime dependencies. Its HTTP contracts were examined in these primary sources:

- [fossteams/teams-api](https://github.com/fossteams/teams-api), pinned revision `dbbdc3681f32` (2022-06-04): authentication exchange, regional hosts, conversation snapshots, profile lookup and message reads. This is the library used by the local teams-cli compatibility experiment.
- The user's local `teams-token` tool: OAuth client/resource IDs and the Microsoft Teams redirect flow, independently implemented with Chrome CDP and signed ID-token validation.
- [Maxim-Mazurok/teams-api chat service](https://github.com/Maxim-Mazurok/teams-api/blob/main/src/api/chat-service.ts): HTTP contracts for chat sends, edits, deletion and reactions. These write contracts have fixture coverage here, not live validation.

These are private Teams API contracts, not an official Microsoft support commitment. No upstream Go library or Electron runtime is shipped with the plugin. The isolated Go probe retains its own pinned module metadata in experiments/teams-api/probe.
