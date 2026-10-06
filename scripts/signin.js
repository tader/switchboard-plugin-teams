import path from 'node:path';
import { TeamsBrowser } from '../lib/browser.js';

const profile = path.resolve(process.argv[2] ?? '.local-profile');
const browser = new TeamsBrowser(profile);
await browser.login();
console.log('Sign in in the dedicated Chrome window. Press Ctrl+C when done.');
console.log('Profile:', profile);
process.once('SIGINT', async () => { await browser.close(); process.exit(0); });
let last;
const timer = setInterval(async () => {
  try {
    const status = await browser.serial(() => browser.dom('status'));
    const encoded = JSON.stringify(status);
    if (encoded !== last) { console.log(status); last = encoded; }
  } catch (error) { console.log(error.code ?? error.message); }
}, 5000);
process.once('SIGINT', () => clearInterval(timer));
