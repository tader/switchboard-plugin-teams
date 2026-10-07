import { generateKeyPairSync, sign } from 'node:crypto';
import { resources } from '../lib/oauth.js';

export const tenant = '11111111-1111-1111-1111-111111111111';
export const oid = '22222222-2222-2222-2222-222222222222';
const { publicKey, privateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
export const key = { ...publicKey.export({ format: 'jwk' }), kid: 'import-test', use: 'sig' };
export function signedToken(service, changes = {}, headerChanges = {}) {
  const claims = { tid: tenant, oid, aud: resources[service], iss: `https://sts.windows.net/${tenant}/`, exp: Math.floor(Date.now() / 1000) + 3600, upn: 'test@example.com', ...changes };
  const encoded = [{ alg: 'RS256', kid: key.kid, ...headerChanges }, claims].map(value => Buffer.from(JSON.stringify(value)).toString('base64url')).join('.');
  return encoded + '.' + sign('RSA-SHA256', Buffer.from(encoded), privateKey).toString('base64url');
}
export function signedBundle(changes = {}) {
  const tokens = Object.fromEntries(Object.keys(resources).map(service => [service, { value: signedToken(service, changes), expiresAt: Date.now() + 3600_000 }]));
  return { authVersion: 1, identity: { tenant: changes.tid ?? tenant, oid: changes.oid ?? oid, email: 'test@example.com', name: 'Test user' }, tokens };
}
export const keysFetcher = async url => {
  if (url !== 'https://login.microsoftonline.com/common/discovery/keys') throw new Error('Unexpected signing key URL');
  return new Response(JSON.stringify({ keys: [key] }));
};
