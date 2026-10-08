import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { checkRefreshCompatibility } from '../lib/refresh-probe.js';
import { AdapterError } from '../lib/errors.js';

export async function main() {
  const controller = new AbortController();
  const abort = () => controller.abort();
  process.once('SIGINT', abort); process.once('SIGTERM', abort);
  try {
    const { values } = parseArgs({ options: { silent: { type: 'boolean' }, 'login-hint': { type: 'string' } }, strict: true, allowPositionals: false });
    console.log(values.silent ? 'Capturing Teams renewal credentials using the saved diagnostic browser session.' : 'Complete Teams sign-in in the dedicated diagnostic browser. It closes after refresh-token capture.');
    const result = await checkRefreshCompatibility({
      directory: path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../.local-profile/refresh-check'),
      silent: values.silent, browserPath: process.env.TEAMS_BROWSER_PATH,
      loginHint: values['login-hint'] || process.env.TEAMS_LOGIN_HINT, signal: controller.signal,
      report: message => console.log(message),
    });
    console.log('Earliest renewed access/chat-token expiry:', new Date(result.expiresAt).toISOString());
    console.log('Captured source grant:', result.sourceGrant);
    console.log('PASS: browser-closed HTTP renewal and read-only Teams validation');
    console.log('No captured tokens were exported or saved by the diagnostic. No messages were sent.');
    console.log('The dedicated browser profile retains Microsoft session data and Teams app caches.');
    console.log('This does not establish refresh-token lifetime, cookie-only recovery, or central-host compatibility.');
  } catch (error) {
    console.error(controller.signal.aborted ? 'Refresh compatibility check cancelled.' : error instanceof AdapterError ? `${error.code}: ${error.message}` : 'Refresh compatibility check failed. No credentials were exported.');
    process.exitCode = controller.signal.aborted ? 130 : 1;
  } finally {
    process.removeListener('SIGINT', abort); process.removeListener('SIGTERM', abort);
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) await main();
