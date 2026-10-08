# API parity protocol evidence

The 0.11.0 feature branch restores server message/people search, exact result
opening, direct exact-ID quotes, person mentions, virtual one-to-one chats and
channel root/thread/history routes. Implementations use Node built-ins. No
upstream runtime or browser is used for data operations.

## Authentication and search

Search uses `https://substrate.office.com` access tokens, acquired on demand
with the existing Teams refresh token and the fixed Microsoft token endpoint.
Base Skype/CSA tokens remain intact. Rotations are returned to Switchboard for
encrypted persistence during authorization, before dispatch. Local diagnostics
persist rotations through the shared owner-only session cache. An old bundle
without renewal credentials reports `search_token_required`; the adapter does
not open a browser merely to obtain a search token.

People suggestions use POST `/search/api/v1/suggestions`, a People entity query,
Mailbox/Directory provenances, the `peoplepicker.newChat` scenario and a fresh
`Cvid`. Message search uses POST `/search/api/v2/query`, a Message entity request,
Teams content source, Optimized property set, bounded `from`/`size`, a fresh
`cvid` and the `powerbar` scenario. References are listed in [THIRD_PARTY](../THIRD_PARTY.md).
Initial simplified requests returned HTTP 400 on this account; corrected
requests passed live on 2026-10-08. Fixture acceptance alone did not establish
the private protocol.

Live message results supplied `ClientConversationId`, `ClientThreadId` and
numeric `InternetMessageId`. The adapter checks consistent conversation/root
fields and stores only exact locators in account-bound expiring handles. It also
accepts exact message deep links on the two fixed Teams HTTPS origins. It never
fetches a supplied result URL. Membership is revalidated before reading; result
snippets do not select a recipient or authorize a send. Unsupported/missing
locators report `canOpen=false`, and removed messages fail explicitly.

## Exact messages, quotes and mentions

The regional exact-message GET is
`/v1/users/ME/conversations/<encoded conversation>/messages/<encoded message ID>`.
A live exact-message read and opening a chat search hit passed with the browser
closed. Direct quotes use this lookup, share the escaped quote builder with
triage and recheck selected content/version before dispatch. The private API
provides no atomic guard against a change after that check.

Person mention runs resolve an exact directory email again. Server profile MRIs
and display names supply escaped `schema.skype.com/Mention` spans and serialized
`properties.mentions` entries with matching numeric item IDs and `mentionType=person`.
Caller names do not select identities. At most 20 distinct people are resolved
per message. Structured chat sends, triage replies and edits share this builder.
Existing quotes/attachments remain protected from editing.

Quotes and mentions have source and fixture evidence. No quote or mention was
sent live; native rendering and notifications remain designated-chat checks.

## Virtual one-to-one conversations

The primary client reference resolves a new organization one-to-one chat as
`19:<self oid>_<peer oid>@unq.gbl.spaces`, persisted by its first message. The
adapter reuses existing directory chats first. A virtual opaque ID carries the
exact email and thread ID; resolution re-fetches that profile and verifies both
account and peer before accepting the virtual ID after reload. Self-chat,
group creation and arbitrary foreign virtual IDs are refused.

First sends use the existing durable message ledger. Recipient-bound operation
fingerprints replay after restart; uncertainty never triggers an automatic
resend. Fixture checks establish these local guarantees. Actual new-chat
persistence has not been live-verified; no arbitrary recipient was contacted.

## Channel reads and history

Membership-bound exact root reads use the regional message GET. Replies are
read from `<channel>;messageid=<numeric root ID>`, the source-verified composite
conversation. Reply-as-root and explicit sibling-thread mismatches are refused.
Root listing traverses the base channel with backward links and filters replies.

History validates every backward link against the fixed regional origin and
exact conversation path. Cursors bind account, channel, root, limit and format;
replay repeats the same page. The verified root is included once, overlap IDs
are deduplicated and NDJSON is available. A first thread page can include its
root in addition to the bounded reply page. Retention/unavailable history remain
outside `completeHistory=false` coverage.

Live checks on 2026-10-08 passed channel root listing and continuation, exact
thread reads and a thread NDJSON export. That sampled thread had no older reply
page. Deep reply continuation and opening a channel search hit remain live
validation work; cross-thread/cross-account cursor protection, overlap, replay
and redirect rejection have fixture coverage. Channel sends/replies, channel
read markers and automatic read-after-reply remain disabled.

## Shared-session and central renewal checks

One initial local session capture was followed by a separate silent API check,
forced HTTP renewal, and multiple read-only feature diagnostics. They reused the
shared saved credentials without repeated sign-in windows. Central `auth:sync`
returned `server_renewed` and verified fresh expiry/readiness on the existing
imported-token connection without a browser or token upload. These results do
not establish indefinite token lifetime, cookie-only recovery after expiry,
central reconnect upload or durable ledger recovery.
