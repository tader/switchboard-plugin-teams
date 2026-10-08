import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { BrowserAuthenticator, exchangeChatToken } from '../lib/api-auth.js';
import { TeamsAPI } from '../lib/teams-api.js';
import { TeamsTriage } from '../lib/triage.js';
import { AdapterError } from '../lib/errors.js';

// Diagnostic authentication reuses the export profile; tokens stay in memory.
// No send ledger is needed: only inbox and context reads are exercised.
let auth;
const controller = new AbortController();
const abort = () => { controller.abort(); auth?.close().catch(() => {}); };
process.once('SIGINT', abort); process.once('SIGTERM', abort);
try {
  let values;
  try { ({ values } = parseArgs({ options: { silent: { type: 'boolean' }, 'login-hint': { type: 'string' } }, strict: true, allowPositionals: false })); }
  catch { throw new AdapterError('invalid_arguments', 'Usage: npm run check:triage -- [--silent] [--login-hint <email>]', 400); }
  console.log(values.silent ? 'Checking triage through the saved authentication session.' : 'Complete sign-in in the dedicated Teams export browser. This diagnostic only reads messages.');
  auth = new BrowserAuthenticator(path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../.local-profile/token-export'), { browserPath: process.env.TEAMS_BROWSER_PATH });
  const bundle = await auth.acquire({ interactive: !values.silent, loginHint: values['login-hint'] || process.env.TEAMS_LOGIN_HINT });
  controller.signal.throwIfAborted();
  const api = new TeamsAPI({ credentials: await exchangeChatToken(bundle) });
  const triage = new TeamsTriage(api, null);
  const inbox = await triage.inbox({ signal: controller.signal });
  console.log(`PASS: API inbox; ${inbox.items.length} groups, ${inbox.items.filter(item => item.kind === 'channel').length} channel threads, ${inbox.coverage.issues.length} coverage issues; hasMore=${inbox.hasMore}.`);
  if (inbox.items.some(item => item.replyTarget)) {
    const contexts = await triage.contexts({ targets: inbox.items.filter(item => item.replyTarget).slice(0, 3).map(item => item.replyTarget) }, controller.signal);
    const errors = contexts.items.filter(item => item.error);
    console.log(`Context results: ${contexts.items.length - errors.length} successful; ${errors.length} unavailable.`);
    if (errors.length) console.log('Context error codes:', errors.map(item => item.error.code).join(', '));
  }
  console.log('No messages sent or read markers written. No tokens or message contents printed or exported.');
  console.log('Native chat quotes are source-verified; channel replies/read-marker writes still require designated live validation.');
} catch (error) {
  console.error(controller.signal.aborted ? 'Triage check cancelled.' : error instanceof AdapterError ? `${error.code}: ${error.message}` : 'Triage check failed.');
  process.exitCode = controller.signal.aborted ? 130 : 1;
} finally { await auth?.close(); process.removeListener('SIGINT', abort); process.removeListener('SIGTERM', abort); }
