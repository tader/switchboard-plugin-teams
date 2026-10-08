# Triage protocol verification

The three triage helpers use the existing regional Teams API transport. Private
contracts are verified independently of fixture success. `getTeamsCapabilities`
reports the production gates; callers also receive `canReply` and separate send
and read outcomes.

Live read check, 2026-10-08: after Microsoft Authenticator approval, ten inbox
groups and three batched reply-context reads succeeded with the browser closed.
Coverage issues remained explicit. This validates bounded chat reads, not channel
completeness or writes.

| Contract | Evidence | Production status |
| --- | --- | --- |
| Chat quote | [Inline quote builder](https://github.com/Maxim-Mazurok/teams-api/blob/e109d58122160c59cc210ec3284cd5a3d0f5ac8c/src/html-utils.ts), [chat send](https://github.com/Maxim-Mazurok/teams-api/blob/e109d58122160c59cc210ec3284cd5a3d0f5ac8c/src/api/chat-service.ts) | Source-verified and fixture-tested, not live-validated |
| Channel thread reads | [Thread conversation IDs and reads](https://github.com/Maxim-Mazurok/teams-api/blob/e109d58122160c59cc210ec3284cd5a3d0f5ac8c/src/teams-client.ts) | Source-verified and fixture-tested, bounded reads |
| Channel unread metadata | [CSA channel and horizon models](https://github.com/fossteams/teams-api/blob/dbbdc3681f32/pkg/models/teams.go) | Require explicit unread flag and valid horizon; unknown data is a coverage issue |
| Channel reply | Candidate composite conversation POST | **Disabled** pending a captured native Teams request and designated-thread validation |
| Read update | Candidate regional consumption-horizon PUT; [Skype predecessor](https://github.com/Terrance/SkPy/blob/master/skpy/chat.py) is not sufficient Teams verification | **Disabled**; successful replies report `read.status=deferred`, `reason=protocol_unverified` |

Chat quotes independently construct escaped microdata identifying the exact
message and sender, include at most 200 characters of preview, and use the
existing rich-text send contract. No upstream implementation is bundled.

The channel reply and read-update implementations are candidate contracts behind
immutable production gates in `lib/triage.js`. Tests inject enabled contracts only
to exercise queueing, durable read stages and conservative advancement. Those
tests establish local behavior, not server compatibility. Synthetic redacted
examples in `test/fixtures/triage.json` are labeled accordingly.

Before enabling either write gate, capture the native Teams request in an
explicitly designated test chat/channel. Record the endpoint, non-secret body,
read-marker scope and successful response; replace identifiers, people and text
with synthetic values before adding fixtures. Verify the actual channel reply
appears beneath its intended root. Verify the marker advances only through the
observed cutoff, preserves later arrivals/newer markers, and cannot clear an
unhandled sibling thread. If scope cannot be established, leave the gate disabled.
Do not send to an arbitrary unread conversation to obtain evidence.

`npm run check:triage -- --silent` authenticates through the saved export profile,
closes the browser and exercises bounded inbox/context reads. Without `--silent`
it permits interactive sign-in. It prints counts and sanitized errors, never
tokens or message contents, and performs no sends or read-marker writes.
