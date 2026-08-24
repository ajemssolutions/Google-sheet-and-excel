const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

/**
 * Storage for the connector.
 *
 * Everything is keyed by ORG. An org is one AJEMS workspace, identified by
 * its tenant name, and it owns its own AJEMS credentials, Google account,
 * tasks, uploads and dedupe links. Nothing is global: a visitor who has not
 * signed in sees an empty app.
 *
 * Sessions map a browser cookie to an org. The cookie holds a random id and
 * nothing else, so it is useless on its own.
 *
 * The shape is deliberately close to how it would look in a database - an
 * orgs table keyed by tenant, a sessions table pointing at it - so moving off
 * a JSON file later is a port rather than a redesign.
 */

const FILE_NAME = 'data.json';
const SESSION_DAYS = 30;

let FILE = null;
let cache = null;

const DEFAULTS = { orgs: {}, sessions: {} };

function blankOrg(tenant) {
  return {
    tenant,
    ajems: { baseUrl: '', secretKey: '', verified: false },
    tasks: [],
    uploads: [],
    links: {},
    createdAt: new Date().toISOString()
  };
}

let deferred = 0;
let dirty = false;

function init(dir) {
  FILE = path.join(dir || path.join(__dirname, '..'), FILE_NAME);
  load();
  return cache;
}

function load() {
  if (!FILE) FILE = path.join(__dirname, '..', FILE_NAME);
  try {
    const raw = JSON.parse(fs.readFileSync(FILE, 'utf8'));
    cache = Object.assign({}, DEFAULTS, raw);
    if (!cache.orgs) cache.orgs = {};
    if (!cache.sessions) cache.sessions = {};
    migrate(raw);
  } catch (e) {
    cache = JSON.parse(JSON.stringify(DEFAULTS));
  }
  pruneSessions();
  return cache;
}

/**
 * Older builds kept one global connection with tasks tagged by tenant.
 * Fold that into the org it belonged to so nothing is lost on upgrade.
 */
function migrate(raw) {
  if (!raw || !raw.ajems || !raw.ajems.tenant) return;
  const tenant = raw.ajems.tenant;
  if (cache.orgs[tenant]) return;

  const o = blankOrg(tenant);
  o.ajems = {
    baseUrl: raw.ajems.baseUrl || '',
    secretKey: raw.ajems.secretKey || '',
    verified: !!raw.ajems.verified
  };
  o.tasks = (raw.tasks || []).filter(t => !t.tenant || t.tenant === tenant);
  o.uploads = raw.uploads || [];
  o.links = raw.links || {};
  cache.orgs[tenant] = o;

  delete cache.ajems; delete cache.google; delete cache.tasks;
  delete cache.uploads; delete cache.links; delete cache.log;
  save();
  console.log('[info] Existing data moved into workspace "' + tenant + '".');
}

function save() {
  if (deferred) { dirty = true; return true; }
  try {
    fs.writeFileSync(FILE, JSON.stringify(cache, null, 2), 'utf8');
  } catch (e) {
    console.error('Could not write data.json:', e.message);
  }
  return true;
}

/** Runs fn with saving held back, then writes once if anything changed. */
function batch(fn) {
  deferred++;
  try {
    return fn();
  } finally {
    deferred--;
    if (!deferred && dirty) { dirty = false; save(); }
  }
}

function all() {
  if (!cache) load();
  return cache;
}

// ── orgs ──────────────────────────────────────────────────
function getOrg(tenant) {
  if (!tenant) return null;
  return all().orgs[tenant] || null;
}

function org(tenant) {
  const c = all();
  if (!c.orgs[tenant]) { c.orgs[tenant] = blankOrg(tenant); save(); }
  return c.orgs[tenant];
}

function orgTenants() {
  return Object.keys(all().orgs);
}

function setOrgAjems(tenant, ajems) {
  const o = org(tenant);
  o.ajems = { baseUrl: ajems.baseUrl, secretKey: ajems.secretKey, verified: true };
  save();
  return o;
}

// ── sessions ──────────────────────────────────────────────
function newSession(tenant) {
  const id = crypto.randomBytes(24).toString('hex');
  all().sessions[id] = { tenant, at: Date.now(), seen: Date.now(), google: null };
  save();
  return id;
}

/**
 * A Google credential lives in one of two places, and both are read and
 * written through this small handle so google.js does not care which:
 *
 *   - on a SESSION, which is what the browser uses. One person signing in
 *     with Google does not connect it for everyone else in the workspace.
 *   - on a TASK, captured when the task is saved, which is what the
 *     scheduler uses. A task has to keep syncing after its author closes
 *     the browser, so it carries its own copy.
 */
function sessionGoogle(id) {
  return {
    get: () => (all().sessions[id] || {}).google || null,
    set: g => {
      const sess = all().sessions[id];
      if (sess) { sess.google = g; save(); }
    }
  };
}

function taskGoogle(tenant, taskId) {
  return {
    get: () => (getTask(tenant, taskId) || {}).google || null,
    set: g => {
      const t = getTask(tenant, taskId);
      if (t) { t.google = g; save(); }
    }
  };
}

/** The tenant a session belongs to, or '' if unknown or expired. */
function sessionTenant(id) {
  if (!id) return '';
  const s = all().sessions[id];
  if (!s) return '';
  if (Date.now() - s.at > SESSION_DAYS * 864e5) {
    delete all().sessions[id];
    save();
    return '';
  }
  if (!all().orgs[s.tenant]) return '';
  s.seen = Date.now();
  return s.tenant;
}

function endSession(id) {
  if (id && all().sessions[id]) { delete all().sessions[id]; save(); }
  return true;
}

function pruneSessions() {
  const c = all();
  const cutoff = Date.now() - SESSION_DAYS * 864e5;
  let changed = false;
  Object.keys(c.sessions).forEach(id => {
    if (c.sessions[id].at < cutoff) { delete c.sessions[id]; changed = true; }
  });
  if (changed) save();
}

// ── tasks, scoped to one org ──────────────────────────────
function newTask() {
  return {
    id: 't_' + Date.now() + '_' + Math.floor(Math.random() * 9000 + 1000),
    name: 'Untitled task',
    enabled: true,

    source: 'sheet',       // 'sheet' (Google) or 'excel' (an uploaded file)
    uploadId: '',
    spreadsheetId: '',
    spreadsheetName: '',
    sheetTabs: [],
    sheetTab: '',
    headerRow: 1,

    appId: '',
    appTitle: '',
    selections: [],        // { tab, formId, formTitle, postUrl, getUrl, mapping }

    // Google credential captured when the task was saved, so scheduled runs
    // survive the author closing their browser. Null for uploaded files.
    google: null,

    schedule: 'hourly',
    updateExisting: true,
    identity: 'hash',
    keyColumns: [],

    lastModifiedTime: '',
    lastRun: null,
    lastResult: null,
    consecutiveFailures: 0,
    pausedReason: ''
  };
}

function tasks(tenant) {
  const o = getOrg(tenant);
  return o ? o.tasks : [];
}

function getTask(tenant, id) {
  return tasks(tenant).find(t => t.id === id) || null;
}

function upsertTask(tenant, task) {
  const o = org(tenant);
  const i = o.tasks.findIndex(t => t.id === task.id);
  if (i === -1) o.tasks.push(task); else o.tasks[i] = task;
  save();
  return task;
}

function removeTask(tenant, id) {
  const o = getOrg(tenant);
  if (!o) return true;
  o.tasks = o.tasks.filter(t => t.id !== id);
  delete o.links[id];
  save();
  return true;
}

// ── uploads, scoped to one org ────────────────────────────
function uploads(tenant) {
  const o = getOrg(tenant);
  return o ? o.uploads : [];
}

function addUpload(tenant, u) {
  org(tenant).uploads.unshift(u);
  save();
  return u;
}

function removeUpload(tenant, id) {
  const o = getOrg(tenant);
  if (!o) return true;
  o.uploads = o.uploads.filter(u => u.id !== id);
  save();
  return true;
}

// ── row -> AJEMS response links, scoped to one org ────────
function getLinks(tenant, taskId) {
  const o = org(tenant);
  if (!o.links[taskId]) o.links[taskId] = {};
  return o.links[taskId];
}

function clearLinks(tenant, taskId) {
  org(tenant).links[taskId] = {};
  save();
  return true;
}

/** Activity goes to the terminal. Nothing reads a stored copy. */
function log(level, task, message) {
  console.log(`[${level}] ${task ? task + ' - ' : ''}${message}`);
}

module.exports = {
  init, load, save, batch, all,
  getOrg, org, orgTenants, setOrgAjems,
  newSession, sessionTenant, endSession, sessionGoogle, taskGoogle,
  newTask, tasks, getTask, upsertTask, removeTask,
  uploads, addUpload, removeUpload,
  getLinks, clearLinks,
  log, SESSION_DAYS
};
