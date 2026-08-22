const store = require('./store');
const google = require('./google');
const engine = require('./engine');

/**
 * Polls tasks on their schedule, but does NOT read the sheet every tick.
 *
 * Each check is one tiny Drive call for modifiedTime. Only if the timestamp
 * moved since the last sync do we read the sheet and push. That matters at
 * scale for two reasons: it is a fraction of the payload, and it hits the
 * Drive API quota rather than the Sheets API quota (300 reads/min), leaving
 * the Sheets budget for real work.
 */

const INTERVAL_MS = {
  '5min':  5 * 60 * 1000,
  '15min': 15 * 60 * 1000,
  hourly:  60 * 60 * 1000,
  daily:   24 * 60 * 60 * 1000
};

let timer = null;
const running = new Set();
const nextDue = {};        // taskId -> epoch ms

function dueIn(task) {
  const ms = INTERVAL_MS[task.schedule];
  if (!ms) return null;
  // Stagger so 50 tasks on "hourly" do not all fire in the same second.
  const jitter = Math.floor(Math.random() * Math.min(ms * 0.1, 60000));
  return Date.now() + ms + jitter;
}

async function checkTask(task) {
  if (running.has(task.id)) return;
  running.add(task.id);
  try {
    let changed = true;
    try {
      const meta = await google.fileMeta(task.spreadsheetId);
      if (task.lastModifiedTime && meta.modifiedTime === task.lastModifiedTime) {
        changed = false;
      }
    } catch (e) {
      store.log('warn', task.name, 'Change check failed, syncing anyway: ' + e.message);
    }

    if (!changed) {
      store.log('info', task.name, 'No change since last sync — skipped.');
      return;
    }

    await engine.runTask(task.id, 'scheduled', { concurrency: 4 });
  } finally {
    running.delete(task.id);
  }
}

async function tick() {
  // Only the connected workspace's tasks can run: the secret key in hand
  // belongs to that tenant, so another organisation's tasks would fail.
  const tenant = (store.get().ajems || {}).tenant || '';
  if (!tenant) return;

  // Uploaded files are excluded on purpose: nothing changes on disk between
  // uploads, so a timer would re-read identical bytes forever.
  const tasks = store.get().tasks.filter(
    t => t.enabled && INTERVAL_MS[t.schedule] && t.spreadsheetId && t.source !== 'excel' &&
         (!t.tenant || t.tenant === tenant)
  );

  for (const task of tasks) {
    if (nextDue[task.id] === undefined) {
      nextDue[task.id] = dueIn(task);   // first due one interval from startup
      continue;
    }
    if (Date.now() >= nextDue[task.id]) {
      nextDue[task.id] = dueIn(task);
      checkTask(task).catch(e => store.log('error', task.name, e.message));
    }
  }
}

function start() {
  if (timer) return;
  timer = setInterval(tick, 20000);   // evaluate every 20s; tasks fire on their own interval
  store.log('info', '', 'Scheduler started.');
}

function stop() {
  if (timer) clearInterval(timer);
  timer = null;
}

/** What the UI shows next to each task. */
function status() {
  const out = {};
  store.get().tasks.forEach(t => {
    out[t.id] = {
      scheduled: !!INTERVAL_MS[t.schedule] && t.enabled,
      nextDue: nextDue[t.id] || null,
      running: running.has(t.id)
    };
  });
  return out;
}

module.exports = { start, stop, tick, status, INTERVAL_MS };
