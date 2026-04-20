#!/usr/bin/env node

import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import * as readline from 'readline';
import * as fs from 'fs';
import * as path from 'path';
import * as http from 'http';
import { resolveAuthStorePath, saveSecureAuthStore } from './auth-store.js';

const SCOPES = [
  'https://www.googleapis.com/auth/classroom.courses.readonly',
  'https://www.googleapis.com/auth/classroom.course-work.readonly',
  'https://www.googleapis.com/auth/classroom.student-submissions.me.readonly',
  'https://www.googleapis.com/auth/classroom.announcements.readonly',
  'https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly',
  'https://www.googleapis.com/auth/classroom.topics.readonly',
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/forms.body.readonly',
];

function buildLoopbackRedirectUri(redirectUris: string[] = []): string {
  const localhostUri = redirectUris.find((uri) => {
    try {
      const parsed = new URL(uri);
      return parsed.hostname === 'localhost' || parsed.hostname === '127.0.0.1';
    } catch {
      return false;
    }
  });

  if (!localhostUri) {
    return 'http://127.0.0.1:3000/oauth2callback';
  }

  const parsed = new URL(localhostUri);
  const port = parsed.port || '3000';
  const pathname = parsed.pathname === '/' ? '/oauth2callback' : parsed.pathname;

  return `${parsed.protocol}//${parsed.hostname}:${port}${pathname}`;
}

async function waitForAuthorizationCode(redirectUri: string, timeoutMs = 120000): Promise<string | null> {
  const redirectUrl = new URL(redirectUri);

  return new Promise<string | null>((resolve, reject) => {
    let settled = false;
    const server = http.createServer((req, res) => {
      const requestUrl = new URL(req.url || '/', redirectUri);
      const code = requestUrl.searchParams.get('code');
      const error = requestUrl.searchParams.get('error');

      if (error) {
        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(`<h1>Authorization failed</h1><p>${error}</p>`);
        settled = true;
        server.close(() => reject(new Error(`Authorization failed: ${error}`)));
        return;
      }

      if (!code) {
        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<h1>Missing authorization code</h1><p>You can close this tab and try again.</p>');
        return;
      }

      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end('<h1>Authentication successful</h1><p>You can close this tab and return to the terminal.</p>');
      settled = true;
      server.close(() => resolve(code));
    });

    server.on('error', reject);
    server.listen(Number(redirectUrl.port), redirectUrl.hostname);
    setTimeout(() => {
      if (settled) {
        return;
      }

      settled = true;
      server.close(() => resolve(null));
    }, timeoutMs);
  });
}

async function promptForAuthorizationCode(): Promise<string> {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise<string>((resolve) => {
    rl.question('Paste the full redirected URL or just the authorization code: ', (answer) => {
      rl.close();
      const trimmed = answer.trim();

      try {
        const maybeUrl = new URL(trimmed);
        const code = maybeUrl.searchParams.get('code');
        resolve(code || trimmed);
      } catch {
        resolve(trimmed);
      }
    });
  });
}

async function setupAuthentication() {
  console.log('Google Classroom MCP Server Authentication Setup');
  console.log('==================================================\n');

  // Check if credentials file exists
  const credentialsPath = path.join(process.cwd(), 'credentials.json');
  if (!fs.existsSync(credentialsPath)) {
    console.error('credentials.json not found!');
    console.log('\nTo set up authentication:');
    console.log('1. Go to https://console.cloud.google.com/');
    console.log('2. Create a new project or select an existing one');
    console.log('3. Enable the Google Classroom API');
    console.log('4. Enable the Google Forms API (for quiz/test extraction support)');
    console.log('5. Go to "Credentials" and create an OAuth 2.0 Client ID');
    console.log('6. Choose "Desktop application" as the application type');
    console.log('7. Download the credentials and save as "credentials.json" in this directory');
    console.log('8. Run this setup script again\n');
    process.exit(1);
  }

  // Load credentials
  const credentialsContent = fs.readFileSync(credentialsPath, 'utf8');
  const credentials = JSON.parse(credentialsContent);
  
  const { client_id, client_secret, redirect_uris } = credentials.installed || credentials.web;
  
  if (!client_id || !client_secret) {
    console.error('Invalid credentials.json format');
    console.error('Make sure you downloaded the correct OAuth 2.0 Client ID credentials');
    process.exit(1);
  }

  // Create OAuth2 client
  const redirectUri = buildLoopbackRedirectUri(redirect_uris);
  const oauth2Client = new OAuth2Client(client_id, client_secret, redirectUri);

  // Generate authorization URL
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent', // Force consent screen to get refresh token
  });

  console.log('Please visit this URL to authorize the application:');
  console.log(`\n${authUrl}\n`);
  console.log(`Listening for the Google redirect on ${redirectUri}`);
  console.log('If the browser still shows a localhost error, copy the full redirected URL and paste it here.\n');

  let code: string | null = null;

  try {
    code = await waitForAuthorizationCode(redirectUri, 120000);
  } catch (error) {
    console.error('Failed to receive authorization code automatically:', error);
  }

  if (!code) {
    console.log('Automatic redirect was not received within 2 minutes.');
    code = await promptForAuthorizationCode();
  }

  if (!code) {
    console.error('No authorization code received.');
    process.exit(1);
  }

  try {
    // Exchange authorization code for tokens
    const { tokens } = await oauth2Client.getToken(code);
    
    if (!tokens.refresh_token) {
      console.error('No refresh token received. Please try again and make sure to grant all permissions.');
      process.exit(1);
    }

    const authStorePath = resolveAuthStorePath(process.env.GOOGLE_AUTH_STORE);
    const materialCacheDbPath = process.env.GOOGLE_MATERIAL_CACHE_DB
      || path.join(path.dirname(authStorePath), 'material-cache.sqlite');
    let secureStoreEnabled = true;

    try {
      saveSecureAuthStore(authStorePath, {
        clientId: client_id,
        clientSecret: client_secret,
        redirectUri,
        refreshToken: tokens.refresh_token,
      });
    } catch (secureStoreError) {
      secureStoreEnabled = false;
      console.warn('Secure store is unavailable in this environment. Falling back to plain .env token storage.');
      console.warn(`Secure store error: ${secureStoreError instanceof Error ? secureStoreError.message : secureStoreError}`);
    }

    // Save tokens to .env file (new method)
    const envLines = [
      '# Google Classroom MCP Server Environment Variables',
      `# Generated on ${new Date().toISOString()}`,
      `GOOGLE_CLIENT_ID="${client_id}"`,
      `GOOGLE_CLIENT_SECRET="${client_secret}"`,
      `GOOGLE_REDIRECT_URI="${redirectUri}"`,
    ];

    if (secureStoreEnabled) {
      envLines.push(`GOOGLE_AUTH_STORE="${authStorePath}"`);
    } else {
      envLines.push(`GOOGLE_REFRESH_TOKEN="${tokens.refresh_token}"`);
    }
    envLines.push(`GOOGLE_MATERIAL_CACHE_DB="${materialCacheDbPath}"`);

    const envContent = `${envLines.join('\n')}
`;

    fs.writeFileSync('.env', envContent);

    if (process.env.WRITE_LEGACY_TOKENS === 'true') {
      const tokensForLegacy = {
        access_token: tokens.access_token,
        refresh_token: tokens.refresh_token,
        scope: tokens.scope,
        token_type: tokens.token_type,
        expiry_date: tokens.expiry_date
      };
      fs.writeFileSync('tokens.json', JSON.stringify(tokensForLegacy, null, 2));
    }

    console.log('Authentication successful!');
    if (secureStoreEnabled) {
      console.log(`Refresh token saved securely to: ${authStorePath}`);
      console.log('Non-secret settings saved to .env');
    } else {
      console.log('Refresh token saved in .env (legacy fallback mode).');
    }
    if (process.env.WRITE_LEGACY_TOKENS === 'true') {
      console.log('Legacy tokens.json created because WRITE_LEGACY_TOKENS=true');
    }
    console.log('\nYou can now run the MCP server with:');
    console.log('   npm run build && npm start');
    console.log('\nOr test it with:');
    console.log('   npm test');
    console.log('\nAdd to Claude Desktop config:');
    console.log('   See README.md for configuration instructions');

  } catch (error) {
    console.error('Error getting tokens:', error);
    process.exit(1);
  }
}

setupAuthentication().catch(console.error);
