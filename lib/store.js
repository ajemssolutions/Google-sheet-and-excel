const fs = require('fs');
const path = require('path');

const FILE = path.join(__dirname, '..', 'data.json');

const DEFAULTS = {
  google: null,          // { access_token, refresh_token, expires_at, email }
  ajems: {               // AJEMS JSON Builder connection
    baseUrl: '',
    secretKey: '',
    tenant: '',
    verified: false
  },
  uploads: [],           // { id, name, tabs, at } - shared, not owned by a task
  tasks: [],             // see newTask()
  links: {},             // { [taskId]: { [rowKey]: ajemsResponseId } }
};

let cache = null;

function load() {
  try {
    cache = Object.assign({}, DEFAULTS, JSON.parse(fs.readFileSync(FILE, 'utf8')));
    delete cache.log;                 // written by builds that had an activity page
  } catch (e) {
    cache = JSON.parse(JSON.stringify(DEFAULTS));
  }
  return cache;
}

let deferred = 0;      // >0 while a batch is open
let dirty = false;

function save() {
  if (deferred) { dirty = true; return true; }
  try {
    fs.writeFileSync(FILE, JSON.stringify(cache, null, 2), 'utf8');
  } catch (e) {
    console.error('Could not write data.json:', e.message);
  }
  return true;
}

/**
 * Runs fn with saving held back, then writes once if anything changed.
 *
 * A sync finishes with a save, a task update and two log lines - four full
 * rewrites of a file that is around a megabyte once a large sheet has been
 * synced. This makes it one.
 */
function batch(fn) {
  deferred++;
  try {
    return fn();
  } finally {
    deferred--;
    if (!deferred && dirty) { dirty = false; save(); }
  }
}

function get() {
  if (!cache) load();
  return cache;
}

function set(patch) {
  Object.assign(get(), patch);
  save();
  return cache;
}

// ── uploaded files ────────────────────────────────────────
function addUpload(u) {
  get().uploads.unshift(u);
  save();
  return u;
}

function removeUpload(id) {
  const c = get();
  c.uploads = c.uploads.filter(u => u.id !== id);
  save();
  return true;
}

function newTask() {
  return {
    id: 't_' + Date.now() + '_' + Math.floor(Math.random() * 9000 + 1000),
    name: 'Untitled task',
    enabled: true,

    // source: 'sheet' (Google) or 'excel' (an uploaded file)
    source: 'sheet',
    uploadId: '',          // set when source is 'excel'
    spreadsheetId: '',     // set when source is 'sheet'
    spreadsheetName: '',   // display name either way
    sheetTabs: [],         // mirrors selections[].tab, for display
    sheetTab: '',          // kept so tasks saved by older builds still run
    headerRow: 1,          // always 1; kept so stored tasks stay valid

    // destination
    appId: '',
    appTitle: '',

    // one entry per tab: { tab, formId, formTitle, postUrl, getUrl, mapping }
    // mapping is [{ column, fieldKey, label, fieldType }]
    selections: [],

    // behaviour
    schedule: 'manual',      // manual | 5min | 15min | hourly | daily
    updateExisting: true,    // PATCH when a linked row changed
    identity: 'hash',        // 'hash' (whole row) | 'key' (keyColumns)
    keyColumns: [],

    // runtime
    lastModifiedTime: '',    // Drive modifiedTime seen at last sync
    lastRun: null,
    lastResult: null,
    consecutiveFailures: 0,
    pausedReason: ''
  };
}

function getTask(id) {
  return get().tasks.find(t => t.id === id) || null;
}

function upsertTask(task) {
  const list = get().tasks;
  const i = list.findIndex(t => t.id === task.id);
  if (i === -1) list.push(task); else list[i] = task;
  save();
  return task;
}

function removeTask(id) {
  const c = get();
  c.tasks = c.tasks.filter(t => t.id !== id);
  delete c.links[id];
  save();
  return true;
}

// ── row -> AJEMS response id links ────────────────────────
function getLinks(taskId) {
  const c = get();
  if (!c.links[taskId]) c.links[taskId] = {};
  return c.links[taskId];
}

function clearLinks(taskId) {
  get().links[taskId] = {};
  save();
  return true;
}

// ── activity log ──────────────────────────────────────────
/** Activity goes to the terminal. Nothing reads a stored copy any more. */
function log(level, task, message) {
  console.log(`[${level}] ${task ? task + ' - ' : ''}${message}`);
}

module.exports = {
  load, save, batch, get, set,
  addUpload, removeUpload,
  newTask, getTask, upsertTask, removeTask,
  getLinks, clearLinks,
  log
};
