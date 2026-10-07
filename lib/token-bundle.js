import { AdapterError } from './errors.js';
import { fresh, importRequired } from './api-auth.js';
import { resources, validateMicrosoftClaims } from './oauth.js';
import { recipientEmail } from './input.js';

const format = 'switchboard-teams-tokens';
const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const invalid = () => new AdapterError('invalid_token_bundle', 'Invalid Teams token bundle. Export a new bundle with auth:export.', 400);
function object(value, keys) {
  return value && typeof value === 'object' && !Array.isArray(value) && Object.keys(value).every(key => keys.includes(key));
}
function shape(value) {
  if (!object(value, ['format', 'version', 'identity', 'tokens']) || value.format !== format || value.version !== 1 ||
      !object(value.identity, ['tenant', 'oid', 'email', 'name']) || !uuid.test(value.identity.tenant ?? '') || !uuid.test(value.identity.oid ?? '') ||
      typeof value.identity.email !== 'string' || typeof value.identity.name !== 'string' || value.identity.name.length > 200 ||
      !object(value.tokens, ['skype', 'chatsvcagg'])) throw invalid();
  try { recipientEmail(value.identity.email); } catch { throw invalid(); }
  for (const service of Object.keys(resources)) {
    const token = value.tokens[service];
    if (!object(token, ['value', 'expiresAt']) || typeof token.value !== 'string' || token.value.length > 100_000 || !Number.isFinite(token.expiresAt)) throw invalid();
    if (!fresh(token)) throw importRequired();
  }
  return value;
}

export function encodeTokenBundle(credentials) {
  const { tenant, oid, email, name } = credentials.identity;
  const tokens = Object.fromEntries(Object.keys(resources).map(service => [service, { value: credentials.tokens[service].value, expiresAt: credentials.tokens[service].expiresAt }]));
  const encoded = JSON.stringify(shape({ format, version: 1, identity: { tenant, oid, email, name }, tokens }));
  if (Buffer.byteLength(encoded) > 128 * 1024) throw invalid();
  return encoded;
}

export async function validateTokenBundle(raw, previous, fetcher = fetch) {
  if (typeof raw !== 'string' || Buffer.byteLength(raw) > 128 * 1024) throw invalid();
  let value;
  try { value = JSON.parse(raw); } catch { throw invalid(); }
  shape(value);
  const identity = { ...value.identity, email: recipientEmail(value.identity.email) };
  const tokens = {};
  for (const service of Object.keys(resources)) {
    let claims;
    try { claims = await validateMicrosoftClaims(value.tokens[service].value, { audience: resources[service], tenant: identity.tenant }, fetcher); }
    catch { throw invalid(); }
    if (claims.oid !== identity.oid) throw new AdapterError('account_mismatch', 'All tokens must belong to the same Teams account.', 403);
    tokens[service] = { value: value.tokens[service].value, expiresAt: Math.min(value.tokens[service].expiresAt, claims.exp * 1000) };
    if (!fresh(tokens[service])) throw importRequired();
  }
  if (previous?.identity && (identity.tenant !== previous.identity.tenant || identity.oid !== previous.identity.oid)) {
    throw new AdapterError('account_mismatch', 'Import tokens for the same Teams account as this connection.', 403);
  }
  return { authVersion: 1, authMode: 'tokens', identity, tokens };
}
