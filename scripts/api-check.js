import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { BrowserAuthenticator, exchangeChatToken } from '../lib/api-auth.js';
import { TeamsAPI } from '../lib/teams-api.js';
import { AdapterError } from '../lib/errors.js';

const silent = process.argv.includes('--silent');
const directory = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../.local-profile/api');
const auth = new BrowserAuthenticator(directory, { browserPath: process.env.TEAMS_BROWSER_PATH });
process.once('SIGINT', () => { auth.close().finally(() => { process.exitCode = 130; }); });
try {
  console.log(silent ? 'Testing silent authentication with the saved diagnostic Chrome profile.' : 'Complete Microsoft sign-in in the dedicated Chrome window. This check only reads Teams.');
  const bundle = await auth.acquire({ interactive: !silent, loginHint: process.env.TEAMS_LOGIN_HINT });
  console.log('PASS: OAuth authentication; Chrome closed');
  const credentials = await exchangeChatToken(bundle);
  const api = new TeamsAPI({ credentials });
  await api.profile(process.env.TEAMS_LOGIN_HINT || credentials.identity.email);
  console.log('PASS: direct API profile read');
  const chats = await api.chats();
  console.log('PASS: direct API conversation list');
  if (chats.items.length) {
    await api.messages(chats.items[0].id, 5);
    console.log('PASS: direct API recent messages read');
  } else console.log('SKIP: recent messages (no chats returned)');
  console.log('Earliest token expiry:', new Date(Math.min(credentials.tokens.skype.expiresAt, credentials.tokens.chatsvcagg.expiresAt, credentials.chatToken.expiresAt)).toISOString());
  console.log('API check complete. No messages were sent. API tokens were held only in memory.');
} catch (error) {
  console.error('API check failed:', error instanceof AdapterError ? `${error.code}: ${error.message}` : 'runtime_or_network_failure');
  process.exitCode = 1;
} finally { await auth.close(); }
