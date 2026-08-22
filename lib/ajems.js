const store = require('./store');

/**
 * AJEMS JSON Builder API client.
 *
 * Two hard-won rules baked in here:
 *  1. Every endpoint MUST keep its trailing slash. Without it nginx answers
 *     301, and any client that follows the redirect turns POST into GET and
 *     silently drops the body. That was the cause of repeated "405 Not
 *     Allowed" errors during the Excel Bridge work.
 *  2. Redirects are never followed automatically. A 3xx is reported as a
 *     real error naming the Location, so a bad URL says so instead of
 *     failing later in a confusing way.
 */

const TIMEOUT_MS = 20000;

function withSlash(u) {
  if (!u) return '';
  let s = String(u).trim();
  let hash = '', query = '';
  const hi = s.indexOf('#'); if (hi > -1) { hash = s.slice(hi); s = s.slice(0, hi); }
  const qi = s.indexOf('?'); if (qi > -1) { query = s.slice(qi); s = s.slice(0, qi); }
  if (!s.endsWith('/')) s += '/';
  return s + query + hash;
}

/**
 * The workspace URL is forgiving. People paste the browser UI domain
 * (kathaa.ajems.in), the API host (kathaa.buildprohub-server.com), or just
 * the tenant name. Only the API host actually serves JSON — ajems.in returns
 * the HTML app — so instead of demanding the right one we work out the tenant
 * and try every known host until one authenticates.
 */
const API_DOMAINS = ['buildprohub-server.com', 'ajems.in'];

function extractTenant(raw) {
  let s = String(raw || '').trim().replace(/\s+/g, '');
  if (!s) return '';
  s = s.replace(/^https?:\/\//i, '');           // strip scheme
  s = s.split('/')[0];                          // strip path
  s = s.split(':')[0];                          // strip port
  if (!s.includes('.')) return s.toLowerCase(); // already a bare tenant
  return s.split('.')[0].toLowerCase();         // first label is the tenant
}

/** Every base URL worth trying, most likely first. */
function candidateWorkspaceUrls(raw) {
  const s = String(raw || '').trim();
  if (!s) return [];

  const out = [];
  const push = u => { if (u && !out.includes(u)) out.push(u); };

  // localhost / explicit host:port — honour it exactly, don't guess domains
  if (/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?/i.test(s)) {
    let base = s.replace(/\/+$/, '');
    if (!/\/json_builder$/i.test(base)) base += '/json_builder';
    push(base + '/');
    return out;
  }

  const tenant = extractTenant(s);
  if (!tenant) return out;

  // If they typed a full host, try that host first — they may know something.
  const typedHost = s.replace(/^https?:\/\//i, '').split('/')[0];
  if (typedHost.includes('.') && !/(^|\.)ajems\.in$/i.test(typedHost)) {
    push('https://' + typedHost + '/json_builder/');
  }

  API_DOMAINS.forEach(d => push(`https://${tenant}.${d}/json_builder/`));
  return out;
}

/** Kept for compatibility — the first candidate. */
function normalizeBase(raw) {
  return candidateWorkspaceUrls(raw)[0] || '';
}

function conn() {
  const a = store.get().ajems || {};
  return { baseUrl: a.baseUrl || '', secretKey: a.secretKey || '' };
}

function headers(extra) {
  const c = conn();
  return Object.assign({
    'X-Json-Builder-Secret-Key': c.secretKey,
    'Accept': 'application/json'
  }, extra || {});
}

async function request(method, url, body) {
  const target = withSlash(url);
  const opts = {
    method,
    headers: headers(body ? { 'Content-Type': 'application/json' } : null),
    redirect: 'manual',
    signal: AbortSignal.timeout(TIMEOUT_MS)
  };
  if (body !== undefined) opts.body = JSON.stringify(body);

  const res = await fetch(target, opts);

  if (res.status >= 300 && res.status < 400) {
    const loc = res.headers.get('location') || '(none)';
    throw new Error(
      `Redirect ${res.status} from ${target} -> ${loc}. ` +
      'The URL is not the final endpoint; AJEMS paths need a trailing slash.'
    );
  }

  const text = await res.text();
  let data = null;
  try { data = text ? JSON.parse(text) : null; } catch (e) { data = null; }

  if (!res.ok) {
    const detail = data
      ? (data.detail || data.error || data.message || JSON.stringify(data).slice(0, 300))
      : text.slice(0, 300);
    const err = new Error(`HTTP ${res.status}: ${detail || '(empty response)'}`);
    err.status = res.status;
    err.body = data || text;
    throw err;
  }

  if (data === null && text.trim().startsWith('<')) {
    throw new Error(
      'The server returned HTML, not JSON. Check the base URL — the API host ' +
      'is usually <tenant>.buildprohub-server.com, not the browser UI domain.'
    );
  }

  return data;
}

const get   = (url)       => request('GET', url);
const post  = (url, body) => request('POST', url, body);
const patch = (url, body) => request('PATCH', url, body);

/** Try one base URL. Returns {ok, data} or {ok:false, reason, fatal}. */
async function tryWorkspace(base, secretKey) {
  const url = withSlash(base + 'workspace_config');
  let res;

  try {
    res = await fetch(url, {
      method: 'GET',
      headers: { 'X-Json-Builder-Secret-Key': secretKey, 'Accept': 'application/json' },
      redirect: 'manual',
      signal: AbortSignal.timeout(TIMEOUT_MS)
    });
  } catch (e) {
    const why = e.name === 'TimeoutError'
      ? 'no response within 20s'
      : (e.cause && e.cause.code) || e.message;
    return { ok: false, reason: `could not reach it (${why})` };
  }

  if (res.status >= 300 && res.status < 400) {
    return { ok: false, reason: `redirected (HTTP ${res.status})` };
  }

  const text = await res.text().catch(() => '');
  const looksLikeHtml = text.trim().startsWith('<');

  let parsed = null;
  try { parsed = text ? JSON.parse(text) : null; } catch (e) { parsed = null; }

  // A wrong key fails on EVERY host, so stop rather than trying the rest.
  // But only when the refusal really came from the API: DRF answers with a
  // JSON body. A proxy, firewall or CDN in front of an unrelated host also
  // answers 401/403, in HTML or plain text, and that is not a key problem.
  if (res.status === 401 || res.status === 403) {
    if (parsed && typeof parsed === 'object') {
      return { ok: false, fatal: true, reason: `the secret key was rejected (HTTP ${res.status})` };
    }
    const snippet = text.trim().slice(0, 120).replace(/\s+/g, ' ');
    return {
      ok: false,
      reason: `HTTP ${res.status} from something that is not the AJEMS API` +
              (snippet ? ` — ${snippet}` : '')
    };
  }

  if (!res.ok) {
    return { ok: false, reason: `HTTP ${res.status}` };
  }
  if (looksLikeHtml) {
    return { ok: false, reason: 'returned HTML, not JSON — this is the browser UI, not the API' };
  }

  let data;
  try { data = JSON.parse(text); }
  catch (e) { return { ok: false, reason: 'response was not valid JSON' }; }

  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    return { ok: false, reason: 'response was not a workspace object' };
  }
  // An EMPTY apps list is a legitimate new workspace; a MISSING apps key is not.
  if (!('apps' in data) || !Array.isArray(data.apps)) {
    return { ok: false, reason: 'response has no "apps" array' };
  }
  if (!data.tenant || typeof data.tenant !== 'string') {
    return { ok: false, reason: 'response has no "tenant"' };
  }

  return { ok: true, data };
}

/**
 * One call that both tests the connection and returns everything:
 * tenant + apps + their forms + per-form URLs. Tries each candidate host.
 */
async function workspaceConfig(baseUrl, secretKey) {
  const candidates = candidateWorkspaceUrls(baseUrl);
  if (!candidates.length) throw new Error('Enter a workspace URL or just your tenant name.');

  const tried = [];

  for (const base of candidates) {
    const r = await tryWorkspace(base, secretKey);
    if (r.ok) return { base, data: r.data };

    tried.push(`  ${base}\n    → ${r.reason}`);
    if (r.fatal) {
      throw new Error(
        `${r.reason}.\n\nCheck the secret key in AJEMS under Settings → JSON Builder.`
      );
    }
  }

  throw new Error(
    'Could not reach an AJEMS API on any known host.\n\n' + tried.join('\n') +
    '\n\nThe API host is usually https://<tenant>.buildprohub-server.com — ' +
    'the ajems.in address is the browser UI and serves HTML, not JSON.'
  );
}

/** Pull the per-form URLs out of workspace_config, falling back to the documented shape. */
function formUrls(base, app, form) {
  const u = form.jsonBuilderUrls || form.json_builder_urls || {};
  // The API calls this form_id; the others are defensive fallbacks.
  const id = form.form_id || form.id || form._id;
  return {
    formId: id,
    detail:    withSlash(u.form     || u.detail   || (base + 'forms/' + id)),
    responses: withSlash(u.responses|| u.response || (base + 'forms/' + id + '/responses'))
  };
}

async function createApp(base, payload) {
  return post(base + 'apps/', payload);
}

async function createForm(base, appId, fields, meta) {
  return post(base + 'forms/', Object.assign({
    app: appId,
    app_id: appId,
    title: meta.title,
    description: meta.description || '',
    fields: fields,
    isMasterForm: false,
    isThirdPartyEnabled: true,
    isThirdPartyGetAllowed: true,
    isThirdPartyPostAllowed: true,
    isPublicPostRecaptchaEnabled: false,
    allowFormColStat: false,
    allowToClone: false,
    clonedBy: ''
  }, meta.extra || {}));
}

async function getForm(detailUrl) {
  return get(detailUrl);
}

async function listResponses(responsesUrl) {
  const data = await get(responsesUrl);
  if (Array.isArray(data)) return data;
  return data.responses || data.results || data.data || [];
}

async function createResponse(responsesUrl, payload) {
  return post(responsesUrl, payload);
}

async function updateResponse(responsesUrl, responseId, payload) {
  return patch(responsesUrl + responseId + '/', payload);
}

module.exports = {
  withSlash, normalizeBase, candidateWorkspaceUrls, extractTenant,
  workspaceConfig, tryWorkspace, formUrls,
  createApp, createForm, getForm,
  listResponses, createResponse, updateResponse,
  get, post, patch
};
