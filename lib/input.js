import { AdapterError } from './errors.js';

export function recipientEmail(value) {
  if (typeof value !== 'string' || value.length > 254 || !/^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?$/.test(value.trim()) || value.includes(',')) {
    throw new AdapterError('invalid_recipient', 'email must be one exact Teams sign-in email/UPN, not a display name or a list of people.', 400);
  }
  return value.trim().toLowerCase();
}

export function peopleQuery(value) {
  if (typeof value !== 'string' || value.trim().length < 2 || value.length > 150 || /[\u0000-\u001f]/.test(value)) throw new AdapterError('invalid_query', 'query must contain 2–150 characters without control characters.', 400);
  return value.trim();
}
