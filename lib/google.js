const store = require('./store');

const AUTH_URL  = 'https://accounts.google.com/o/oauth2/v2/auth';
const TOKEN_URL = 'https://oauth2.googleapis.com/token';
const DRIVE     = 'https://www.googleapis.com/drive/v3';
const SHEETS    = 'https://sheets.googleapis.com/v4/spreadsheets';

/**
 * drive.file only.
 *
 * It grants access to the specific files the user picks through the Google
 * Picker, and nothing else. Google classes it NON-SENSITIVE, so publishing
 * needs only basic app verification — no annual CASA security assessment,
 * which drive.readonly would force because that scope is Restricted.
 *
 * spreadsheets.readonly is deliberately absent. drive.file is itself a valid
 * Sheets API scope, so the Sheets API reads a picked file's values under it.
 * Adding spreadsheets.readonly would drag a SENSITIVE scope back in for no
 * capability we do not already have.
 */
const SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'openid', 'email', 'profile'
];

const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file';

// Without this a stalled Google request never settles, and the sync it
// belongs to never finishes - the task stays "running" and its next
// scheduled turn is skipped for good.
const TIMEOUT_MS = 20000;

function cfg() {
  // Trim: a trailing space or a stray quote copied out of the console is a
  // common cause of "invalid_client", and it is invisible in the file.
  const clean = v => String(v || '').trim().replace(/^["']|["']$/g, '');
  const clientId = clean(process.env.GOOGLE_CLIENT_ID);
  return {
    clientId,
    clientSecret: clean(process.env.GOOGLE_CLIENT_SECRET),
    apiKey: clean(process.env.GOOGLE_API_KEY),
    // The Cloud project number. The Picker REQUIRES it for the drive.file
    // scope — without it a picked file is never actually shared with the app,
    // and every later read fails with "Requested entity was not found".
    // It is the part of the client ID before the first hyphen, so there is
    // nothing extra to configure; the override exists only as an escape hatch.
    projectNumber: clean(process.env.GOOGLE_PROJECT_NUMBER) || clientId.split('-')[0] || '',
    redirectUri: clean(process.env.GOOGLE_REDIRECT_URI) || 'http://localhost:3000/oauth/callback'
  };
}

/**
 * Checks the credentials look like credentials before we send the user to
 * Google. Without this the only feedback is Google's own "Error 401:
 * invalid_client" page, which never says which value is wrong.
 */
function credentialProblems() {
  const c = cfg();
  const out = [];

  if (!c.clientId) {
    out.push('GOOGLE_CLIENT_ID is empty.');
  } else if (c.clientId.includes('your-client-id')) {
    out.push('GOOGLE_CLIENT_ID is still the placeholder from .env.example.');
  } else if (c.clientId.startsWith('AIza')) {
    out.push('GOOGLE_CLIENT_ID looks like an API key (starts with "AIza"), not an OAuth client ID. ' +
             'The client ID comes from Credentials \u203a OAuth 2.0 Client IDs.');
  } else if (!c.clientId.endsWith('.apps.googleusercontent.com')) {
    out.push('GOOGLE_CLIENT_ID does not end with ".apps.googleusercontent.com", ' +
             'so it is not a valid OAuth client ID.');
  }

  if (!c.clientSecret) {
    out.push('GOOGLE_CLIENT_SECRET is empty.');
  } else if (c.clientSecret.includes('your-client-secret')) {
    out.push('GOOGLE_CLIENT_SECRET is still the placeholder from .env.example.');
  } else if (c.clientSecret.startsWith('AIza')) {
    out.push('GOOGLE_CLIENT_SECRET looks like an API key, not an OAuth client secret.');
  }

  if (!c.apiKey) {
    out.push('GOOGLE_API_KEY is empty. Google\u2019s file picker cannot open without it — ' +
             'create one under Credentials \u203a API key.');
  } else if (c.apiKey.includes('your-api-key')) {
    out.push('GOOGLE_API_KEY is still the placeholder from .env.example.');
  } else if (!c.apiKey.startsWith('AIza')) {
    out.push('GOOGLE_API_KEY does not look like a Google API key (they start with "AIza").');
  }

  if (!/^\d+$/.test(c.projectNumber)) {
    out.push('Could not work out the Cloud project number from GOOGLE_CLIENT_ID. ' +
             'Set GOOGLE_PROJECT_NUMBER in .env (IAM & Admin \u203a Settings).');
  }

  // Any host is allowed, not just localhost: this runs on a real domain once
  // it is deployed. Only the shape is checked; Google itself rejects a URI
  // that does not match the one registered in the console.
  if (!/^https?:\/\/[^\s/]+(\/[^\s]*)?\/oauth\/callback$/.test(c.redirectUri)) {
    out.push('GOOGLE_REDIRECT_URI is "' + c.redirectUri +
             '". It should look like https://your-domain/oauth/callback and match ' +
             'the Authorised redirect URI in Google Cloud exactly.');
  } else if (/^http:\/\//i.test(c.redirectUri) && !/^http:\/\/localhost/i.test(c.redirectUri)) {
    out.push('GOOGLE_REDIRECT_URI uses http on a public host. Google only accepts ' +
             'https there, so use https://your-domain/oauth/callback.');
  }

  return out;
}

function authUrl(state) {
  const c = cfg();
  const scopes = SCOPES;
  const p = new URLSearchParams({
    client_id: c.clientId,
    redirect_uri: c.redirectUri,
    response_type: 'code',
    scope: scopes.join(' '),
    access_type: 'offline',      // we want a refresh token for the scheduler
    prompt: 'consent',           // force it even on re-auth
    include_granted_scopes: 'true'
  });
  // The workspace is carried through Google and handed back on the callback,
  // so tokens land against the org that started the sign-in.
  if (state) p.set('state', state);
  return AUTH_URL + '?' + p.toString();
}

async function exchangeCode(tenant, code) {
  const c = cfg();
  const res = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    signal: AbortSignal.timeout(TIMEOUT_MS),
    body: new URLSearchParams({
      code,
      client_id: c.clientId,
      client_secret: c.clientSecret,
      redirect_uri: c.redirectUri,
      grant_type: 'authorization_code'
    })
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.error || 'Token exchange failed');

  const session = {
    access_token: data.access_token,
    refresh_token: data.refresh_token || ((store.getOrg(tenant) || {}).google || {}).refresh_token || '',
    expires_at: Date.now() + (data.expires_in || 3600) * 1000,
    scope: data.scope || '',
    email: ''
  };

  try {
    const me = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: 'Bearer ' + session.access_token },
      signal: AbortSignal.timeout(TIMEOUT_MS)
    });
    if (me.ok) session.email = (await me.json()).email || '';
  } catch (e) { /* not fatal */ }

  store.setOrgGoogle(tenant, session);
  return session;
}

async function refresh(tenant) {
  const c = cfg();
  const g = (store.getOrg(tenant) || {}).google;
  if (!g || !g.refresh_token) throw new Error('No refresh token. Sign in with Google again.');

  const res = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    signal: AbortSignal.timeout(TIMEOUT_MS),
    body: new URLSearchParams({
      refresh_token: g.refresh_token,
      client_id: c.clientId,
      client_secret: c.clientSecret,
      grant_type: 'refresh_token'
    })
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.error || 'Token refresh failed');

  g.access_token = data.access_token;
  g.expires_at = Date.now() + (data.expires_in || 3600) * 1000;
  store.setOrgGoogle(tenant, g);
  return g;
}

const refreshing = {};   // tenant -> in-flight refresh

/**
 * Valid access token, refreshed if it is within 60s of expiry.
 * Concurrent callers share one refresh rather than each starting their own.
 */
async function token(tenant) {
  const g = (store.getOrg(tenant) || {}).google;
  if (!g || !g.access_token) throw new Error('Not signed in with Google.');

  if (Date.now() > (g.expires_at || 0) - 60000) {
    if (!refreshing[tenant]) {
      refreshing[tenant] = refresh(tenant).finally(() => { delete refreshing[tenant]; });
    }
    return (await refreshing[tenant]).access_token;
  }
  return g.access_token;
}

/**
 * Google grants scopes at CONSENT time, not per request. An older session
 * signed in before drive.readonly was required still holds the narrower
 * grant, and the symptom is silent: files.list returns an EMPTY LIST rather
 * than an error. This is what lets the UI say so plainly instead of showing
 * "no spreadsheets found".
 */
function scopeStatus(tenant) {
  const g = (store.getOrg(tenant) || {}).google;
  if (!g || !g.access_token) {
    return { signedIn: false, ok: false, needed: DRIVE_SCOPE, granted: [] };
  }

  const granted = String(g.scope || '').split(/\s+/).filter(Boolean);

  // A session signed in under an earlier build may hold a broader Drive
  // scope. Those still cover everything drive.file allows, so there is no
  // reason to make the user re-authorise.
  const SUPERSETS = [
    DRIVE_SCOPE,
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.readonly'
  ];

  // Fail OPEN when Google returned no scope string at all: the token may be
  // perfectly good, and the real authority is the API call itself, which now
  // reports a clear error of its own. Failing closed here produced a
  // "sign in again" banner over a working session.
  const ok = granted.length === 0 || granted.some(x => SUPERSETS.includes(x));

  return { signedIn: true, ok, needed: DRIVE_SCOPE, granted };
}

function friendly(e) {
  if (e && e.name === 'TimeoutError') return new Error('Google did not respond within 20 seconds.');
  return e;
}

async function api(tenant, url) {
  const t = await token(tenant);
  let res;
  try {
    res = await fetch(url, {
      headers: { Authorization: 'Bearer ' + t },
      signal: AbortSignal.timeout(TIMEOUT_MS)
    });
  } catch (e) { throw friendly(e); }
  const text = await res.text();
  let body;
  try { body = JSON.parse(text); } catch (e) { body = { raw: text.slice(0, 400) }; }

  if (!res.ok) {
    const msg = (body.error && (body.error.message || body.error)) || ('HTTP ' + res.status);
    const err = new Error(typeof msg === 'string' ? msg : JSON.stringify(msg));
    err.status = res.status;
    throw err;
  }
  return body;
}

// ── Drive ─────────────────────────────────────────────────

/**
 * Lists the spreadsheets this app has been granted, which under drive.file
 * means exactly the ones the user has picked before. An empty result is
 * normal on a fresh install — it means "nothing picked yet", not an error.
 * Grants accumulate, so this list grows with each pick and nothing has to be
 * disconnected to add another sheet.
 */
async function listSpreadsheets(tenant) {
  const st = scopeStatus(tenant);
  if (!st.ok) {
    const err = new Error(
      'This Google sign-in does not include the Drive permission the app needs.\n\n' +
      'Google grants permissions at sign-in, so an older session keeps the ' +
      'permissions it was given.\n\n' +
      'Go to Connections and click "Sign in with Google" again.'
    );
    err.code = 'SCOPE_MISMATCH';
    err.status = 409;
    throw err;
  }

  const p = new URLSearchParams({
    q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
    fields: 'files(id,name,modifiedTime,owners(emailAddress))',
    orderBy: 'modifiedTime desc',
    pageSize: '100'
  });
  const data = await api(tenant, DRIVE + '/files?' + p.toString());
  return data.files || [];
}

/** One cheap call used by the scheduler to decide whether to read the sheet. */
async function fileMeta(tenant, spreadsheetId) {
  return api(tenant, DRIVE + '/files/' + encodeURIComponent(spreadsheetId) +
             '?fields=id,name,modifiedTime');
}

// ── Sheets ────────────────────────────────────────────────

async function listTabs(tenant, spreadsheetId) {
  const data = await api(tenant, SHEETS + '/' + encodeURIComponent(spreadsheetId) +
                         '?fields=properties.title,sheets.properties');
  return {
    title: (data.properties || {}).title || '',
    tabs: (data.sheets || []).map(s => ({
      title: s.properties.title,
      rows: (s.properties.gridProperties || {}).rowCount || 0,
      cols: (s.properties.gridProperties || {}).columnCount || 0
    }))
  };
}

/** Rows as objects keyed by header name, plus the sheet row number. */
async function readTab(tenant, spreadsheetId, tabTitle, headerRow) {
  const range = encodeURIComponent("'" + String(tabTitle).replace(/'/g, "''") + "'");
  const data = await api(tenant,
    SHEETS + '/' + encodeURIComponent(spreadsheetId) + '/values/' + range +
    '?majorDimension=ROWS&valueRenderOption=UNFORMATTED_VALUE' +
    '&dateTimeRenderOption=FORMATTED_STRING'
  );

  const grid = data.values || [];
  const hIndex = Math.max(0, (parseInt(headerRow, 10) || 1) - 1);
  const headers = (grid[hIndex] || []).map(h => String(h == null ? '' : h).trim());

  const rows = [];
  for (let i = hIndex + 1; i < grid.length; i++) {
    const raw = grid[i] || [];
    if (!raw.some(c => String(c == null ? '' : c).trim() !== '')) continue;
    const obj = { __rowNumber: i + 1 };
    headers.forEach((h, c) => {
      if (!h) return;
      const v = raw[c];
      obj[h] = v == null ? '' : String(v).trim();
    });
    rows.push(obj);
  }

  return { headers: headers.filter(Boolean), rows };
}

function signOut(tenant) {
  store.setOrgGoogle(tenant, null);
  return true;
}

module.exports = {
  SCOPES, cfg, credentialProblems, authUrl, exchangeCode, refresh, token,
  scopeStatus,
  listSpreadsheets, fileMeta, listTabs, readTab, signOut
};
