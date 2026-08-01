'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');

/**
 * Google access tokens, minted from a refresh token already on this machine.
 *
 * Two credential sources, in preference order:
 *   1. ~/.pm_drive_refresh  — written by scratchpad/drive_auth.py, Drive + Sheets
 *   2. ~/.clasprc.json      — the clasp `default` profile, which carries full Drive scope
 *
 * The clasp profile is what makes this work without a fresh consent round-trip. It is a
 * user credential, so everything runs as the signed-in user against their own Drive.
 */

const CLIENT_JSON = path.join(
  os.homedir(),
  'Downloads',
  'client_secret_890339390227-f5v12mj9s9crt2a8g6uspi4ur902o86h.apps.googleusercontent.com.json'
);

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

function credentials() {
  const driveRefresh = path.join(os.homedir(), '.pm_drive_refresh');
  if (fs.existsSync(driveRefresh) && fs.existsSync(CLIENT_JSON)) {
    const client = readJson(CLIENT_JSON).installed;
    return {
      source: '~/.pm_drive_refresh',
      clientId: client.client_id,
      clientSecret: client.client_secret,
      refreshToken: fs.readFileSync(driveRefresh, 'utf8').trim(),
    };
  }

  const clasprc = path.join(os.homedir(), '.clasprc.json');
  if (fs.existsSync(clasprc)) {
    const tokens = readJson(clasprc).tokens || {};
    // `default` is the profile verified to hold https://www.googleapis.com/auth/drive.
    const profile = tokens.default || tokens._tu55h4r_default;
    if (profile && profile.refresh_token) {
      return {
        source: '~/.clasprc.json (default)',
        clientId: profile.client_id,
        clientSecret: profile.client_secret,
        refreshToken: profile.refresh_token,
      };
    }
  }

  throw new Error(
    'No Google credential found. Expected ~/.pm_drive_refresh or a clasp login in ~/.clasprc.json.'
  );
}

let cached = null;

/** Access tokens last ~1h; cache within a run and refresh a minute before expiry. */
async function getAccessToken() {
  if (cached && Date.now() < cached.expiresAt - 60000) return cached.token;

  const cred = credentials();
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: cred.clientId,
      client_secret: cred.clientSecret,
      refresh_token: cred.refreshToken,
      grant_type: 'refresh_token',
    }),
  });

  if (!res.ok) {
    const detail = await res.text();
    throw new Error(
      `Token refresh failed (${res.status}) using ${cred.source}. ` +
        `If this says invalid_grant the credential was revoked and needs re-authorising.\n${detail}`
    );
  }

  const data = await res.json();
  cached = {
    token: data.access_token,
    expiresAt: Date.now() + (data.expires_in || 3600) * 1000,
    scope: data.scope || '',
  };

  if (!/\/auth\/drive(\s|$)/.test(cached.scope)) {
    throw new Error(
      `Credential ${cred.source} lacks full Drive scope, which is required to upload and ` +
        `convert. Scopes granted: ${cached.scope}`
    );
  }
  return cached.token;
}

module.exports = { getAccessToken, credentials };
