const store = require('./store');
const google = require('./google');
const engine = require('./engine');

/**
 * Polls every workspace's tasks on their schedule.
 *
 * The scheduler is not tied to whoever has a browser open: each org keeps its
 * own AJEMS key and Google tokens, so a task keeps syncing whether or not
 * anyone is signed in.
 *
 * A tick does NOT read the sheet. It asks Drive for the file's modifiedTime,
 * one small call on a separate quota from the Sheets API, and only reads when
 * that timestamp has moved.
 */

const INTERVAL_MS = {
  '5min':  5 * 60 * 1000,
  '15min': 15 * 60 * 1000,
  hourly:  60 * 60 * 1000,
  daily:   24 * 60 * 60 * 1000
};

let timer = null;
const running = new Set();     // "tenant:taskId" while a sync is in flight
const nextDue = {};            // "tenant:taskId" -> epoch ms

const keyOf = (tenant, id) => tenant + ':' + id;

function dueIn(task) {
  const ms = INTERVAL_MS[task.schedule];
  if (!ms) return null;
  // Stagger so many tasks on one interval do not all fire together.
  const jitter = Math.floor(Math.random() * Math.min(ms * 0.1, 60000));
  return Date.now() + ms + jitter;
}

/** Tasks a timer should touch: Google sheets only, enabled, with a target. */
function scheduled(tenant) {
  return store.tasks(tenant).filter(t =>
    t.enabled &&
    INTERVAL_MS[t.schedule] &&
    t.spreadsheetId &&
    t.source !== 'excel' &&      // an uploaded file only changes on re-upload
    t.google &&                  // captured when the task was saved
    (t.selections || []).some(s => s.postUrl)
  );
}

async function checkTask(tenant, task) {
  const k = keyOf(tenant, task.id);
  if (running.has(k)) return;
  running.add(k);
  try {
    let changed = true;
    try {
      const meta = await google.fileMeta(store.taskGoogle(tenant, task.id), task.spreadsheetId);
      if (task.lastModifiedTime && meta.modifiedTime === task.lastModifiedTime) {
        changed = false;
      }
    } catch (e) {
      store.log('warn', task.name, 'Change check failed, syncing anyway: ' + e.message);
    }
    if (!changed) return;
    await engine.runTask(tenant, task.id, 'scheduled', { concurrency: 4 });
  } finally {
    running.delete(k);
  }
}

function tick() {
  store.orgTenants().forEach(tenant => {
    const org = store.getOrg(tenant);
    if (!org || !org.ajems || !org.ajems.verified) return;   // nothing to sync with

    scheduled(tenant).forEach(task => {
      const k = keyOf(tenant, task.id);
      if (nextDue[k] === undefined) {
        nextDue[k] = dueIn(task);      // first run one interval from now
        return;
      }
      if (Date.now() >= nextDue[k]) {
        nextDue[k] = dueIn(task);
        checkTask(tenant, task).catch(e => store.log('error', task.name, e.message));
      }
    });
  });
}

function start() {
  if (timer) return;
  timer = setInterval(tick, 20000);
  store.log('info', '', 'Scheduler started.');
}

function stop() {
  if (timer) clearInterval(timer);
  timer = null;
}

/** What the task list shows next to each task, for one org. */
function status(tenant) {
  const out = {};
  store.tasks(tenant).forEach(t => {
    const k = keyOf(tenant, t.id);
    out[t.id] = {
      scheduled: !!INTERVAL_MS[t.schedule] && t.enabled && t.source !== 'excel',
      nextDue: nextDue[k] || null,
      running: running.has(k)
    };
  });
  return out;
}

module.exports = { start, stop, tick, status, INTERVAL_MS };
