# API migration follow-up

The production adapter uses Teams APIs with either local browser authentication or imported tokens for a browser-free host. Token export/import, signature checks, manual renewal, restart recovery and ledger retention have fixture coverage; a local export and read-only import on the central host still need live validation. The earlier Go probe validated real token acquisition and three read operations. The new Node adapter needs its read-only live check and tenant-specific silent-renewal check run on the user's Mac.

Remaining API migrations are message search, new-chat creation, quoted replies, exact-person mentions, unread marking, and channel thread reads/writes. Attachments remain unsupported. Unsupported operations are excluded from OpenAPI and reported by getTeamsCapabilities. They must acquire verified API contracts and meaningful tests before being advertised.

Further validation: designated-chat live writes, pagination on large accounts, token renewal across tenant sign-in policy expiry, and independent code review. No live write should be inferred from fixture test success.
