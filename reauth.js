/**
 * Re-auth script: generates a new OAuth refresh token that includes Gmail scopes.
 *
 * Run:  node reauth.js
 *
 * Two flows:
 *   A) Port 8080 is free  → browser callback is captured automatically.
 *   B) Port 8080 is busy  → browser shows "connection refused" after consent;
 *                           copy the full URL from the address bar and paste it
 *                           (or just the code= value) into the terminal prompt.
 */
import { google } from 'googleapis';
import http from 'http';
import fs from 'fs';
import { createInterface } from 'readline';
import 'dotenv/config';

const CLIENT_ID     = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI  = 'http://localhost:8080/oauth2callback';

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.file',
];

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('❌ GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET missing from .env');
  process.exit(1);
}

const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

const authUrl = oauth2Client.generateAuthUrl({
  access_type: 'offline',
  scope: SCOPES,
  prompt: 'consent',   // force new refresh_token even if previously authorised
});

console.log('\n🔐 Re-authorisation required to add Gmail scope.\n');
console.log('📋 Scopes being requested:');
SCOPES.forEach(s => console.log('   •', s));
console.log('\n🌐 Opening browser …');
console.log('\nAuthorisation URL (open manually if browser does not open):');
console.log(authUrl, '\n');

// Try to open browser
try {
  const open = (await import('open')).default;
  await open(authUrl);
} catch { /* ignore */ }

// ── Try to start a one-shot callback server on port 8080 ────────────────────
function tryBind(port) {
  return new Promise((resolve) => {
    const srv = http.createServer();
    srv.once('error', () => resolve(null));   // port busy → null
    srv.listen(port, () => resolve(srv));     // port free → server
  });
}

const callbackServer = await tryBind(8080);
let code = null;

if (callbackServer) {
  console.log('👂 Callback server listening on http://localhost:8080/oauth2callback');
  console.log('   Complete the auth in your browser — this terminal will continue automatically.\n');

  code = await new Promise((resolve) => {
    const timer = setTimeout(() => { callbackServer.close(); resolve(null); }, 5 * 60 * 1000);
    callbackServer.on('request', (req, res) => {
      clearTimeout(timer);
      const url = new URL(req.url, 'http://localhost:8080');
      const c   = url.searchParams.get('code');
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(c
        ? '<h2 style="font-family:sans-serif">✅ Authorised! You can close this tab and return to the terminal.</h2>'
        : '<h2 style="font-family:sans-serif">❌ No code received — please try again.</h2>'
      );
      callbackServer.close();
      resolve(c || null);
    });
  });
} else {
  // Port 8080 is in use — guide the user to paste manually
  console.log('⚠️  Port 8080 is already in use — automatic callback capture unavailable.');
  console.log('');
  console.log('   After clicking "Allow" in the browser, it will show a');
  console.log('   "This site can\'t be reached" page. That is fine.');
  console.log('   Look at the ADDRESS BAR — it will contain:');
  console.log('');
  console.log('     http://localhost:8080/oauth2callback?code=4/0ABCD...&scope=...');
  console.log('');
  console.log('   Copy the ENTIRE URL (or just the value after code=) and paste below.\n');
}

// ── Manual paste fallback ────────────────────────────────────────────────────
if (!code) {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  const answer = await new Promise(resolve => {
    rl.question('🔑 Paste the full callback URL  (or just the code= value): ', a => {
      rl.close();
      resolve(a.trim());
    });
  });

  if (answer.startsWith('http')) {
    try { code = new URL(answer).searchParams.get('code'); } catch { code = answer; }
  } else {
    code = answer;
  }
}

if (!code) {
  console.error('❌ No authorisation code obtained. Exiting.');
  process.exit(1);
}

// ── Exchange code for tokens ─────────────────────────────────────────────────
console.log('\n🔄 Exchanging code for tokens …');
let tokens;
try {
  ({ tokens } = await oauth2Client.getToken(code));
} catch (err) {
  console.error('❌ Token exchange failed:', err.message);
  if (err.message.includes('redirect_uri_mismatch')) {
    console.error('   Ensure http://localhost:8080/oauth2callback is listed as an');
    console.error('   Authorized Redirect URI in your Google Cloud Console OAuth Client.');
  }
  process.exit(1);
}

if (!tokens.refresh_token) {
  console.error('\n❌ No refresh_token returned.');
  console.error('   Google only issues a refresh_token on first consent or after revocation.');
  console.error('   1. Go to https://myaccount.google.com/permissions');
  console.error('   2. Revoke access for your app');
  console.error('   3. Re-run:  npm run reauth');
  process.exit(1);
}

console.log('✅ New refresh token obtained!');

// ── Update .env ──────────────────────────────────────────────────────────────
const envPath = new URL('.env', import.meta.url).pathname;
let envContent = fs.readFileSync(envPath, 'utf8');

if (envContent.includes('GOOGLE_REFRESH_TOKEN=')) {
  envContent = envContent.replace(/GOOGLE_REFRESH_TOKEN=.*/, `GOOGLE_REFRESH_TOKEN=${tokens.refresh_token}`);
} else {
  envContent += `\nGOOGLE_REFRESH_TOKEN=${tokens.refresh_token}\n`;
}

fs.writeFileSync(envPath, envContent);
console.log('💾 .env updated with new GOOGLE_REFRESH_TOKEN');
console.log('\n✨ Done! Run  npm run update  now.\n');
