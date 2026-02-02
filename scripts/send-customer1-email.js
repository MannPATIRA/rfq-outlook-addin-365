#!/usr/bin/env node
/**
 * Sends an email from customer-1 to mannpatira via Microsoft Graph so you can
 * open it in Outlook and start the add-in flow. Uses app-only auth (no app passwords).
 *
 * Usage:  node scripts/send-customer1-email.js   (or:  npm run send-customer1-email)
 *
 * REQUIRES (set in .env yourself; this script does not modify .env):
 *   AZURE_CLIENT_ID      - Application (client) ID of an Azure AD app
 *   AZURE_CLIENT_SECRET  - Client secret for that app
 *   AZURE_TENANT_ID      - Tenant ID or domain (e.g. hexa729.onmicrosoft.com)
 *   TO_EMAIL             - Recipient (e.g. mannpatira@hexa729.onmicrosoft.com)
 *
 * Azure app must have Mail.Send (Application permission) and admin consent.
 * See scripts/SEND-CUSTOMER1-SETUP.md for step-by-step.
 */

const path = require('path');
try {
  require('dotenv').config({ path: path.join(__dirname, '..', '.env') });
} catch (_) {}

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const TENANT_ID = process.env.AZURE_TENANT_ID || 'hexa729.onmicrosoft.com';
const FROM_EMAIL = 'customer-1@hexa729.onmicrosoft.com';
const TO_EMAIL = process.env.TO_EMAIL || 'mannpatira@hexa729.onmicrosoft.com';

const SUBJECT = 'RFQ – Add-in test ' + new Date().toISOString().slice(0, 19).replace('T', ' ');
const BODY = 'This email was sent by the send-customer1-email script so you can open it in Outlook and use the add-in (Notify Engineering).';

function decodeJwtPayload(token) {
  try {
    const payload = token.split('.')[1];
    if (!payload) return {};
    return JSON.parse(Buffer.from(payload, 'base64url').toString());
  } catch (_) {
    return {};
  }
}

async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error('Token request failed: ' + res.status + ' ' + err);
  }
  const json = await res.json();
  return json.access_token;
}

async function sendMail(accessToken) {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(FROM_EMAIL)}/sendMail`;
  const payload = {
    message: {
      subject: SUBJECT,
      body: { contentType: 'Text', content: BODY },
      toRecipients: [{ emailAddress: { address: TO_EMAIL } }],
    },
    saveToSentItems: true,
  };
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + accessToken,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const errText = await res.text();
    let errJson;
    try {
      errJson = errText ? JSON.parse(errText) : {};
    } catch (_) {
      errJson = { error: { message: errText } };
    }
    const msg = errJson.error?.message || errText || res.statusText;
    const err = new Error('Send mail failed: ' + res.status + ' – ' + msg);
    err.status = res.status;
    err.body = errText;
    throw err;
  }
}

async function main() {
  if (!CLIENT_ID || !CLIENT_SECRET) {
    console.error('Missing AZURE_CLIENT_ID or AZURE_CLIENT_SECRET.');
    console.error('Set them in .env (see scripts/SEND-CUSTOMER1-SETUP.md).');
    process.exit(1);
  }

  console.log('Using app Client ID: %s...', (CLIENT_ID || '').slice(0, 8) + '...');
  console.log('Sending from %s to %s...', FROM_EMAIL, TO_EMAIL);

  let tokenPayload = null;
  try {
    const token = await getAccessToken();
    tokenPayload = decodeJwtPayload(token);
    const roles = tokenPayload.roles || [];
    if (roles.length === 0) {
      console.warn('Warning: token has no application roles. Ensure Mail.Send (Application) is added and admin consent granted for the app with Client ID above.');
    } else if (!roles.includes('Mail.Send')) {
      console.warn('Warning: token roles are [%s]. Mail.Send is missing – use the app that has Mail.Send (Application) in .env.', roles.join(', '));
    }
    await sendMail(token);
    console.log('Sent. Open this email in Outlook (as %s) to start the add-in flow.', TO_EMAIL);
  } catch (err) {
    console.error('Error:', err.message);
    if (err.body) console.error('Response:', err.body);
    if (err.status === 401) {
      console.error('');
      const roles = (tokenPayload && tokenPayload.roles) || [];
      console.error('401 = Unauthorized. Token was for app %s and had roles: [%s]', (tokenPayload && tokenPayload.appid) || '?', roles.join(', ') || 'none');
      console.error('→ .env AZURE_CLIENT_ID must be the app where you added Mail.Send (Application) and granted consent (e.g. HEXA-ENGIONIC-DEMO-365).');
      console.error('→ In Azure, open that app → Overview and copy "Application (client) ID" into .env as AZURE_CLIENT_ID; use that app’s Client secret as AZURE_CLIENT_SECRET.');
    }
    if (err.status === 403 || err.message.includes('Access')) {
      console.error('Ensure the app has Mail.Send (Application) and admin consent.');
    }
    if (err.status === 404 || err.message.includes('not found')) {
      console.error('Check AZURE_TENANT_ID and that user %s exists.', FROM_EMAIL);
    }
    process.exit(1);
  }
}

main();
