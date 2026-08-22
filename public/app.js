const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

let STATUS = {};
let WORKSPACE = null;      // { tenant, apps: [...] }
let TASKS = [];
let UPLOADS = [];          // shared library of uploaded workbooks
let LAST_TASK_SIGNATURE = '';   // avoids redrawing the list when nothing changed
let SCHED = {};

// wizard state
let W = null;              // the task being edited
let WSTEP = 1;
let WCOLUMNS = {};        // tab -> [{ column, field_type, samples, ... }]
let WFIELDS = {};         // formId -> [{ key, label, field_type }]
let WBUSY = false;
let EDITING = false;      // editing an existing task: schedule only
let PICKED_APP = null;    // the AJEMS app chosen in step 2
let MAPTAB = null;        // which tab the mapping step is showing
let SRC = 'sheet';        // 'sheet' (Google) or 'excel' (an uploaded file)
let ROWCOUNTS = {};       // tab -> data row count, for the step 1 summary
let NAMING = null;        // index of the row currently naming a new form
let NAMEDRAFT = '';

// ══════════ helpers ══════════

async function api(path, opts) {
  let res;
  try {
    res = await fetch(path, Object.assign({
      headers: { 'Content-Type': 'application/json' }
    }, opts || {}));
  } catch (e) {
    // fetch() itself threw — this is the local Node server, not AJEMS.
    throw new Error(
      'The local server is not responding.\n\n' +
      'Check the terminal where you ran "npm start" — it may have stopped ' +
      'or printed an error. Restart it and reload this page.'
    );
  }
  const data = await res.json().catch(() => ({ error: 'The server sent a malformed response.' }));
  if (!res.ok) throw new Error(data.error || ('HTTP ' + res.status));
  return data;
}

let toastTimer;
function toast(text, kind) {
  const t = $('toast');
  t.textContent = text;
  t.className = 'toast show' + (kind ? ' ' + kind : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 4200);
}

function msg(id, text, kind) {
  const el = $(id);
  el.className = 'msg ' + (kind || '');
  el.textContent = text;
}

function goto(page) {
  document.querySelectorAll('.nav-item').forEach(b =>
    b.classList.toggle('active', b.dataset.page === page));
  document.querySelectorAll('.page').forEach(p =>
    p.classList.toggle('active', p.id === 'page-' + page));
  if (page === 'tasks') loadTasks();
  if (page.startsWith('guide-')) renderGuide(page);
}

// ══════════ boot ══════════

(async function init() {
  document.querySelectorAll('.nav-item').forEach(b =>
    b.addEventListener('click', () => goto(b.dataset.page)));

  bindConnections();
  bindWizard();

  $('btnNewTask').addEventListener('click', () => openWizard(null));
  document.querySelector('[data-goto-new]').addEventListener('click', () => openWizard(null));
  await refreshStatus();
  loadUploads();
  if (new URLSearchParams(location.search).get('signedin')) {
    toast('Signed in with Google.', 'ok');
    history.replaceState({}, '', '/');
  }
  loadTasks();
  setInterval(loadTasks, 15000);
})();

// ══════════ status / connections ══════════

async function refreshStatus() {
  STATUS = await api('/api/status');

  // The Google button is always shown. Hiding it made the section look broken,
  // with an OR divider and nothing above it. If the server has no credentials
  // the click says so plainly; the reason stays in the terminal.
  $('googleUnavailable').style.display = 'none';

  const gp = $('googlePill');
  if (STATUS.signedIn) {
    gp.textContent = STATUS.email || 'Signed in';
    gp.className = 'pill on';
    $('btnSignIn').lastChild.textContent = ' Switch Google account';
    $('btnSignOut').style.display = '';
  } else {
    gp.textContent = 'Not signed in';
    gp.className = 'pill off';
    $('btnSignOut').style.display = 'none';
  }

  const ap = $('ajemsPill');
  if (STATUS.ajems.verified) {
    ap.textContent = STATUS.ajems.tenant;
    ap.className = 'pill on';
  } else {
    ap.textContent = 'Not connected';
    ap.className = 'pill off';
  }
  if (STATUS.ajems.baseUrl) $('ajemsUrl').value = STATUS.ajems.baseUrl;

  $('tickConnect').className = 'tick' +
    (STATUS.signedIn && STATUS.ajems.verified ? ' on' : '');

  $('statusCard').innerHTML =
    (STATUS.signedIn ? '&#9679; Google: ' + esc(STATUS.email || 'signed in') : '&#9675; Google: not signed in') +
    '<br>' +
    (STATUS.ajems.verified ? '&#9679; AJEMS: ' + esc(STATUS.ajems.tenant) : '&#9675; AJEMS: not connected') +
    '<br>&#9679; Tasks: ' + STATUS.taskCount;

  renderNextStep();
}

/** Tells you the one thing to do next, so nothing has to be guessed. */
function renderNextStep() {
  const box = $('nextStep');

  if (!STATUS.ajems.verified) {
    box.className = 'nextstep';
    box.innerHTML = '<b>Next:</b> connect AJEMS in section 1 &mdash; your tenant name and secret key.';
    return;
  }
  if (!STATUS.signedIn && !UPLOADS.length) {
    box.className = 'nextstep';
    box.innerHTML = STATUS.googleConfigured
      ? '<b>Next:</b> continue with Google, or upload an Excel sheet, in section 2.'
      : '<b>Next:</b> upload an Excel sheet in section 2.';
    return;
  }
  if (!STATUS.signedIn) {
    box.className = 'nextstep done';
    box.innerHTML = '<span><b>Ready.</b> You have an uploaded file to work from - ' +
      'go to <b>Tasks &rsaquo; Add task</b>.</span>' +
      '<button class="btn btn-primary btn-sm" id="btnGoTasks">Go to Tasks</button>';
    $('btnGoTasks').addEventListener('click', () => { goto('tasks'); openWizard(null); });
    return;
  }
  // Scopes are granted at consent time, so an older session keeps its old
  // permissions. Without this the symptom is silent: an empty sheet list.
  if (STATUS.signedIn && !STATUS.scopeOk) {
    box.className = 'nextstep warn';
    box.innerHTML =
      '<span><b>Sign in again.</b> This Google session has no Drive permission' +
      ((STATUS.grantedScopes || []).length
        ? ' — it was granted: ' + esc((STATUS.grantedScopes || []).join(', '))
        : '') + '.</span>' +
      '<a class="btn btn-primary btn-sm" href="/auth/google">Sign in with Google</a>';
    return;
  }

  box.className = 'nextstep done';
  box.innerHTML =
    '<span><b>Both connected.</b> Sheets are chosen per task &mdash; ' +
    'go to <b>Tasks &rsaquo; Add task</b>.</span>' +
    '<button class="btn btn-primary btn-sm" id="btnGoTasks">Go to Tasks</button>';
  $('btnGoTasks').addEventListener('click', () => { goto('tasks'); openWizard(null); });
}

function bindConnections() {
  $('btnSignIn').addEventListener('click', () => {
    if (!STATUS.googleConfigured) {
      return msg('googleMsg', 'Google sign-in is not set up on this server yet. ' +
        'You can still upload an Excel or CSV file.', 'err');
    }
    location.href = '/auth/google';
  });
  $('btnUpload').addEventListener('click', () => $('fileInput').click());
  $('fileInput').addEventListener('change', uploadFile);

  $('btnSignOut').addEventListener('click', async () => {
    await api('/api/google/signout', { method: 'POST' });
    await refreshStatus();
    toast('Signed out of Google.');
  });

  $('btnAjemsConnect').addEventListener('click', async () => {
    msg('ajemsMsg', 'Testing…', 'info');
    try {
      const data = await api('/api/ajems/connect', {
        method: 'POST',
        body: JSON.stringify({
          baseUrl: $('ajemsUrl').value,
          secretKey: $('ajemsKey').value
        })
      });
      WORKSPACE = data;
      const forms = data.apps.reduce((n, a) => n + a.forms.length, 0);
      msg('ajemsMsg', `Connected to "${data.tenant}" — ${data.apps.length} application(s), ${forms} form(s).`, 'ok');
      await refreshStatus();
    } catch (e) {
      msg('ajemsMsg', e.message, 'err');
    }
  });
}

/** Reads the chosen file in the browser and posts it as base64. */
async function uploadFile(fileOrEvent, fromWizard) {
  // Called two ways: from the Connections page (an input event) and from
  // inside the wizard (the File itself), which must not re-open the wizard.
  const file = fileOrEvent instanceof File
    ? fileOrEvent
    : (fileOrEvent.target.files && fileOrEvent.target.files[0]);
  if (!file) return;
  if (!(fileOrEvent instanceof File)) fileOrEvent.target.value = '';

  const where = fromWizard ? 'wPreviewMsg' : 'googleMsg';

  if (file.size > 20 * 1024 * 1024) {
    return msg(where, 'That file is larger than 20 MB. Split it, or use a Google Sheet.', 'err');
  }

  msg(where, 'Uploading ' + file.name + '\u2026', 'info');
  try {
    const base64 = await new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result).split(',')[1]);
      r.onerror = () => reject(new Error('The file could not be read.'));
      r.readAsDataURL(file);
    });

    const up = await api('/api/uploads', {
      method: 'POST',
      body: JSON.stringify({ name: file.name, data: base64 })
    });
    msg(where, '');
    await loadUploads();
    await refreshStatus();

    if (fromWizard) {
      // Already in the wizard: just select the file that was uploaded.
      buildUploadPick();
      await chooseUpload(up.id);
    } else {
      // Uploading is only ever a step towards a task, so open the same flow a
      // Google Sheet would open, with this file already chosen.
      newTaskFromUpload(up.id);
    }
  } catch (err) {
    msg(where, err.message, 'err');
  }
}

/** Opens the task wizard with an uploaded file already selected. */
function newTaskFromUpload(uploadId) {
  goto('tasks');
  SRC = 'excel';               // openWizard reads this when it lays out step 1
  openWizard(null);
  setWizardSource('excel');
  chooseUpload(uploadId);
}

async function loadUploads() {
  try {
    const d = await api('/api/uploads');
    UPLOADS = d.uploads || [];
  } catch (e) { UPLOADS = []; }

  $('uploadList').innerHTML = UPLOADS.length
    ? '<div class="uplist"><div class="label">Uploaded files</div>' + UPLOADS.map(u => `
        <div class="uprow">
          <span class="uprow-ic">
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
              <path fill="#1D6F42" d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <path fill="#A9D5B8" d="M14 2v6h6z"/>
              <path fill="#fff" d="m9.3 12.2 1.5 2.3-1.6 2.5h1.5l.9-1.5.9 1.5h1.5l-1.6-2.5 1.5-2.3h-1.4l-.9 1.4-.9-1.4z"/>
            </svg>
          </span>
          <span class="uprow-main">
            <span class="uprow-name">${esc(u.name)}</span>
            <span class="uprow-sub">${u.tabs.length} sheet(s) &middot; ${esc(u.tabs.join(', '))}</span>
          </span>
          <button class="btn btn-primary btn-sm" data-newfrom="${esc(u.id)}">New task</button>
          <button class="btn btn-danger btn-sm" data-rmupload="${esc(u.id)}">Remove</button>
        </div>`).join('') + '</div>'
    : '';

  $('uploadList').querySelectorAll('[data-newfrom]').forEach(b =>
    b.addEventListener('click', () => newTaskFromUpload(b.dataset.newfrom)));

  $('uploadList').querySelectorAll('[data-rmupload]').forEach(b =>
    b.addEventListener('click', async () => {
      try {
        await api('/api/uploads/' + encodeURIComponent(b.dataset.rmupload), { method: 'DELETE' });
        await loadUploads();
        msg('googleMsg', '');
      } catch (e) { msg('googleMsg', e.message, 'err'); }
    }));
}

// ══════════ tasks list ══════════

async function loadTasks() {
  try {
    const data = await api('/api/tasks');
    TASKS = data.tasks;
    SCHED = data.status || {};
  } catch (e) { return; }

  $('taskEmpty').style.display = TASKS.length ? 'none' : 'block';

  // The poll runs every 15s. Rewriting the list each time wiped whatever a
  // sync had just reported into out_<id>, so only redraw when something the
  // list actually shows has changed.
  const signature = JSON.stringify(TASKS.map(t => [
    t.id, t.name, t.enabled, t.schedule, t.source, t.spreadsheetName, t.appTitle,
    t.pausedReason, t.lastRun, t.lastResult,
    (t.selections || []).map(s => [s.tab, s.formTitle]),
    (SCHED[t.id] || {}).running, (SCHED[t.id] || {}).nextDue
  ]));
  if (signature === LAST_TASK_SIGNATURE) return;
  LAST_TASK_SIGNATURE = signature;

  $('taskList').innerHTML = TASKS.map(t => {
    const st = SCHED[t.id] || {};
    const r = t.lastResult;
    const dot = t.pausedReason ? 'err' : (t.enabled ? 'on' : 'off');
    const when = t.lastRun ? new Date(t.lastRun).toLocaleString() : 'waiting for its first run';

    let stats = '<span>Last run: <b>' + esc(when) + '</b></span>';
    if (r) {
      stats += `<span>Created <b>${r.created}</b></span>` +
               `<span>Updated <b>${r.updated}</b></span>` +
               `<span>Unchanged <b>${r.unchanged}</b></span>` +
               (r.failed ? `<span style="color:var(--err)">Failed <b>${r.failed}</b></span>` : '');
    }
    if (st.running) stats += '<span class="tag">syncing now</span>';
    else if (st.nextDue && t.source !== 'excel') {
      stats += '<span>Next: <b>' + new Date(st.nextDue).toLocaleTimeString() + '</b></span>';
    }

    return `<div class="task">
      <div class="task-head">
        <div>
          <div class="task-name"><span class="dot ${dot}"></span>${esc(t.name)}
            <span class="tag grey">${t.source === 'excel' ? 'One-off import' : esc(scheduleLabel(t.schedule))}</span>
            ${t.pausedReason ? '<span class="tag err">paused</span>' : ''}
          </div>
          <div class="task-path">${t.source === 'excel' ? 'Uploaded file' : 'Google Sheet'} &middot; ${esc(t.spreadsheetName || 'no source')} &nbsp;&rarr;&nbsp; ${esc(t.appTitle || 'no app')}</div>
          <div class="task-path">${taskRoutes(t)}</div>
          <div class="task-stats">${stats}</div>
          ${t.pausedReason ? `<div class="msg err">${esc(t.pausedReason)}</div>` : ''}
        </div>
        <div class="task-actions">
          ${t.source === 'excel'
            ? `<button class="btn btn-primary btn-sm" data-act="newfile" data-id="${t.id}">Upload new file &amp; sync</button>
               <button class="btn btn-ghost btn-sm" data-act="run" data-id="${t.id}">Sync again</button>`
            : `<button class="btn btn-primary btn-sm" data-act="run" data-id="${t.id}">Sync now</button>`}
          <button class="btn btn-ghost btn-sm" data-act="edit" data-id="${t.id}">Edit</button>
          <button class="btn btn-danger btn-sm" data-act="del" data-id="${t.id}">Delete</button>
        </div>
      </div>
      <div id="out_${t.id}"></div>
    </div>`;
  }).join('');

  $('taskList').querySelectorAll('[data-act]').forEach(b =>
    b.addEventListener('click', () => taskAction(b.dataset.act, b.dataset.id, b)));
}

/** "North → Leads · South → South leads" */
function taskRoutes(t) {
  const sels = (t.selections && t.selections.length)
    ? t.selections
    : (t.sheetTab ? [{ tab: t.sheetTab, formTitle: t.formTitle }] : []);
  if (!sels.length) return '<span class="tag grey">no tabs</span>';
  return sels.map(s => {
    const per = ((t.lastResult && t.lastResult.perTab) || []).find(p => p.tab === s.tab);
    const count = per ? ` <span class="tag grey">${per.created + per.updated} sent</span>` : '';
    return esc(s.tab) + ' &rarr; ' + esc(s.formTitle || 'no form') + count;
  }).join(' &nbsp;·&nbsp; ');
}

function scheduleLabel(v) {
  return ({ '5min': 'Every 5 min', '15min': 'Every 15 min',
            hourly: 'Hourly', daily: 'Daily' })[v] || 'Hourly';
}

async function taskAction(act, id, btn) {
  if (act === 'edit') return openWizard(TASKS.find(t => t.id === id));

  // Replacing the file is the only way an uploaded source can change, so it
  // is one action: pick the file, send it, sync, report.
  if (act === 'newfile') {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xlsm,.xls,.csv';
    input.addEventListener('change', async () => {
      const file = input.files && input.files[0];
      if (!file) return;
      const out = $('out_' + id);

      if (file.size > 20 * 1024 * 1024) {
        out.innerHTML = '<div class="msg err">That file is larger than 20 MB.</div>';
        return;
      }

      btn.disabled = true;
      out.innerHTML = '<div class="msg info">Uploading ' + esc(file.name) + ' and syncing\u2026</div>';
      try {
        const base64 = await new Promise((resolve, reject) => {
          const r = new FileReader();
          r.onload = () => resolve(String(r.result).split(',')[1]);
          r.onerror = () => reject(new Error('The file could not be read.'));
          r.readAsDataURL(file);
        });

        const d = await api('/api/tasks/' + id + '/replace-file', {
          method: 'POST',
          body: JSON.stringify({ name: file.name, data: base64 })
        });

        const r = d.result || {};
        if (r.error) throw new Error(r.error);
        out.innerHTML = `<div class="msg ${r.failed ? 'err' : 'ok'}">` +
          esc(file.name) + ': ' + r.rows + ' row(s) read - ' +
          `${r.created} new, ${r.updated} updated, ${r.unchanged} already there` +
          (r.failed ? ', ' + r.failed + ' failed' : '') + '.</div>' +
          (r.errors && r.errors.length ? `<div class="msg err">${esc(r.errors[0])}</div>` : '');
        loadTasks();
        loadUploads();
      } catch (e) {
        out.innerHTML = `<div class="msg err">${esc(e.message)}</div>`;
      }
      btn.disabled = false;
    });
    input.click();
    return;
  }

  if (act === 'run') {
    const out = $('out_' + id);
    btn.disabled = true;
    out.innerHTML = '<div class="msg info">Syncing…</div>';
    try {
      const r = await api('/api/tasks/' + id + '/run', {
        method: 'POST', body: JSON.stringify({ trigger: 'manual', seedFirst: true })
      });
      if (r.error) throw new Error(r.error);
      out.innerHTML = `<div class="msg ${r.failed ? 'err' : 'ok'}">` +
        `${r.rows} row(s) read — ${r.created} created, ${r.updated} updated, ` +
        `${r.unchanged} unchanged, ${r.failed} failed.</div>` +
        (r.errors && r.errors.length ? `<div class="msg err">${esc(r.errors[0])}</div>` : '');
      loadTasks();
    } catch (e) {
      out.innerHTML = `<div class="msg err">${esc(e.message)}</div>`;
    }
    btn.disabled = false;
    return;
  }

  if (act === 'del') {
    if (!confirm('Delete this task? The AJEMS records it already created are not touched.')) return;
    await api('/api/tasks/' + id, { method: 'DELETE' });
    loadTasks(); refreshStatus();
    toast('Task deleted.');
  }
}

// ══════════ setup guides ══════════
//
// Two guides, one renderer. The shapes mirror the Tally Connector setup page
// so the three connectors read as one product: integration header, "before
// you begin", numbered steps with a screenshot beside each, and a closing
// call to action.
//
// To drop a real screenshot in later, add `img: '/shots/whatever.png'` to a
// step. Nothing else changes.

const ICON = {
  sheets: '<svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="#188038" stroke-width="1.7"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/></svg>',
  excel: '<svg viewBox="0 0 24 24" width="26" height="26"><path fill="#1D6F42" d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path fill="#A9D5B8" d="M14 2v6h6z"/><path fill="#fff" d="m9.3 12.2 1.5 2.3-1.6 2.5h1.5l.9-1.5.9 1.5h1.5l-1.6-2.5 1.5-2.3h-1.4l-.9 1.4-.9-1.4z"/></svg>',
  key: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="7.5" cy="15.5" r="3.5"/><path d="m10 13 8-8M17 4l3 3M14 7l3 3"/></svg>',
  building: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 8h1M14 8h1M9 12h1M14 12h1M9 16h6"/></svg>',
  google: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M12 3a15 15 0 0 0 0 18a15 15 0 0 0 0-18M3 12h18"/></svg>',
  file: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>',
  grid: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>',
  clock: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
  alert: '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16.5v.01"/></svg>',
  wifi: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M5 12.5a10 10 0 0 1 14 0M8.5 16a5 5 0 0 1 7 0M12 19.5v.01"/></svg>',
  image: '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="m4 18 5-4 4 3 3-2 4 3"/></svg>'
};

const GUIDES = {
  'guide-sheets': {
    icon: ICON.sheets,
    name: 'Google Sheets',
    blurb: 'Keep a Google Sheet and an AJEMS form in step. Rows you add or change are picked up on a schedule, without anyone exporting anything.',
    meta: [['Runs', 'On a schedule'], ['Needs', 'A Google account'], ['Access', 'Only the sheets you pick']],
    requirements: [
      [ICON.building, 'Your AJEMS organisation name'],
      [ICON.key, 'Your AJEMS secret key'],
      [ICON.google, 'A Google account with the sheet in it'],
      [ICON.grid, 'A sheet whose first row holds the column names']
    ],
    steps: [
      {
        label: 'Step 01', title: 'Connect your AJEMS workspace',
        text: 'The connector writes into one workspace at a time. Tasks are saved against it, so reconnecting later brings them all back.',
        list: ['Open <b>Connections</b>', 'Type your organisation name, for example <code>kathaa</code>', 'Paste the secret key from AJEMS under <b>Settings &rsaquo; JSON Builder</b>', 'Press <b>Test connection</b>'],
        note: 'A full URL works too. If you paste the browser address by mistake, the connector finds the API host itself.',
        shots: [['Connections', 'connections.png']]
      },
      {
        label: 'Step 02', title: 'Continue with Google',
        text: 'Signing in lets the connector open Google\u2019s own file picker. It never reads your Drive; only the sheets you choose become readable.',
        list: ['In section 2, press <b>Continue with Google</b>', 'Choose the account that owns the sheet', 'Read what is being asked for, then press <b>Allow</b>'],
        info: ['Only what you pick', 'The middle permission is the one that matters: the app can only touch the specific files you open with it. Each pick adds to the list rather than replacing it.'],
        reverse: true,
        shots: [['Google consent screen', 'google-consent.png']]
      },
      {
        label: 'Step 03', title: 'Start a task and pick your sheet',
        text: 'A task is one sheet tab going into one AJEMS form. Add as many as you need; they run independently.',
        list: ['Open <b>Tasks</b> and press <b>Add task</b>', 'Leave the source on <b>Google Sheet</b>', 'Press <b>Choose from Google Drive</b>', 'Pick the spreadsheet and press <b>Select</b>'],
        note: 'The picker shows My Drive, files shared with you, and ones you picked before, filtered to Google Sheets only.',
        shots: [['New task', 'task-source-google.png'], ['Choosing a spreadsheet', 'google-picker.png']]
      },
      {
        label: 'Step 04', title: 'Choose the tabs',
        text: 'Tick one or more tabs. Each is read the moment you tick it, so you can see what was detected before going on.',
        list: ['Tick the tabs you want to sync', 'Check the detected type against each column', 'Press <b>Next</b>'],
        info: ['Row 1 is the header', 'Column names come from the first row of each tab. The count beside each column shows how many rows actually have a value.'],
        reverse: true,
        shots: [['Tabs and detected columns', 'tabs-columns.png']]
      },
      {
        label: 'Step 05', title: 'Choose an app, then a form per tab',
        text: 'Pick the AJEMS application that will hold the forms, then give every tab its own form.',
        list: ['Choose an application, or press <b>+ Create new app</b> and name it', 'For each tab, pick a form or press <b>+ New</b> to create one', '<b>Create a form for every tab</b> does the outstanding ones at once'],
        note: 'A new form is built from that tab\u2019s own columns, so its fields line up with the sheet from the start.',
        shots: [['Choosing an application', 'app-grid.png'], ['One form per tab', 'forms-per-tab.png']]
      },
      {
        label: 'Step 06', title: 'Check the mapping',
        text: 'Columns are matched to AJEMS fields by name, ignoring case and punctuation. Everything is shown so you can change any of it.',
        list: ['Review each column and the field it will be sent to', 'Change anything that looks wrong', 'With several tabs, use the bar at the top to switch between them'],
        info: ['What the counts mean', 'The number beside each tab, for example 6/8, is how many of its columns are mapped. A column set to "Do not send" is simply skipped.'],
        reverse: true,
        shots: [['Field mapping', 'mapping.png']]
      },
      {
        label: 'Step 07', title: 'Set the schedule and save',
        text: 'Saving syncs the task straight away, then it keeps running on its own.',
        list: ['Choose how often it runs, from every 5 minutes to daily', 'Choose how a row is recognised: a key column, or the whole row', 'Press <b>Save task</b>'],
        info: ['Nothing is sent twice', 'Each check is one small call to Drive; the sheet is only read when it actually changed. Rows already sent are left alone, new rows are added, and a changed row updates in place when identity is a key column.'],
        shots: [['Schedule', 'schedule.png'], ['The task after a run', 'task-running.png']]
      }
    ],
    cta: ['Ready to connect a sheet?', 'Connect your workspace, sign in with Google, and the first task takes about a minute.', 'connect']
  },

  'guide-excel': {
    icon: ICON.excel,
    name: 'Excel and CSV',
    blurb: 'Bring a spreadsheet from your computer into an AJEMS form. Nothing is scheduled: the file only changes when you upload it again.',
    meta: [['Runs', 'Once, on demand'], ['Needs', 'No Google account'], ['Accepts', '.xlsx, .xlsm, .xls, .csv']],
    requirements: [
      [ICON.building, 'Your AJEMS organisation name'],
      [ICON.key, 'Your AJEMS secret key'],
      [ICON.file, 'An .xlsx, .xlsm, .xls or .csv file'],
      [ICON.grid, 'A sheet whose first row holds the column names']
    ],
    steps: [
      {
        label: 'Step 01', title: 'Connect your AJEMS workspace',
        text: 'The connector writes into one workspace at a time. Tasks are saved against it, so reconnecting later brings them all back.',
        list: ['Open <b>Connections</b>', 'Type your organisation name, for example <code>kathaa</code>', 'Paste the secret key from AJEMS under <b>Settings &rsaquo; JSON Builder</b>', 'Press <b>Test connection</b>'],
        note: 'This is the only step Excel shares with Google Sheets. Everything after it works without a Google account.',
        shots: [['Connections', 'connections.png']]
      },
      {
        label: 'Step 02', title: 'Upload your spreadsheet',
        text: 'You can upload from section 2 of Connections, or from the task itself. Either way the file is read here and its sheet names are listed straight away.',
        list: ['Open <b>Tasks</b> and press <b>Add task</b>', 'Choose <b>Excel or CSV file</b>', 'Press <b>Upload Excel sheet</b> and pick the file, up to 20 MB'],
        info: ['One file, many tasks', 'Uploaded files are listed on the Connections page and shared, so several tasks can read the same file. A file cannot be removed while a task still uses it.'],
        reverse: true,
        shots: [['Choosing the Excel source', 'task-source-excel.png']]
      },
      {
        label: 'Step 03', title: 'Choose the sheets',
        text: 'Tick the sheets you want to import. Each is read the moment you tick it, so you can see what was detected before going on.',
        list: ['Tick the sheets you want', 'Check the detected type against each column', 'Press <b>Next</b>'],
        info: ['Row 1 is the header', 'Column names come from the first row of each sheet. The count beside each column shows how many rows actually have a value.'],
        shots: [['Sheets and detected columns', 'tabs-columns.png']]
      },
      {
        label: 'Step 04', title: 'Choose an app, then a form per sheet',
        text: 'Pick the AJEMS application that will hold the forms, then give every sheet its own form.',
        list: ['Choose an application, or press <b>+ Create new app</b> and name it', 'For each sheet, pick a form or press <b>+ New</b> to create one', '<b>Create a form for every tab</b> does the outstanding ones at once'],
        note: 'A new form is built from that sheet\u2019s own columns, so its fields line up with the file from the start.',
        reverse: true,
        shots: [['Choosing an application', 'app-grid.png'], ['One form per sheet', 'forms-per-tab.png']]
      },
      {
        label: 'Step 05', title: 'Check the mapping and save',
        text: 'Columns are matched to AJEMS fields by name. Everything is shown so you can change any of it, then the import runs as soon as you save.',
        list: ['Review each column and the field it will be sent to', 'Change anything that looks wrong', 'Press <b>Save task</b>'],
        note: 'An uploaded file has no schedule, so the schedule step is fixed to a single import.',
        shots: [['Field mapping', 'mapping.png']]
      },
      {
        label: 'Step 06', title: 'Import a newer copy later',
        text: 'When the spreadsheet changes, give the task the new file. It replaces the old one and syncs in a single action.',
        list: ['Open <b>Tasks</b>', 'Press <b>Upload new file &amp; sync</b> on the task', 'Choose the newer copy of the same spreadsheet'],
        info: ['No duplicates', 'Rows already sent are recognised in the new file and left alone. Only genuinely new rows are added, and a changed row updates in place when identity is a key column.'],
        reverse: true,
        shots: [['A Google task and an uploaded one', 'tasks-both.png']]
      }
    ],
    cta: ['Ready to import a spreadsheet?', 'Connect your workspace and upload a file. No Google account is needed.', 'connect']
  }
};

function shotHtml(step) {
  return (step.shots || []).map(([caption, file]) => `<figure class="shot">
    <div class="shot-bar">
      <span class="sdot sdot-r"></span><span class="sdot sdot-y"></span><span class="sdot sdot-g"></span>
      <span class="shot-bar-title">${esc(caption)}</span>
    </div>
    <img class="shot-img" src="shots/${esc(file)}" alt="${esc(caption)}" loading="lazy">
  </figure>`).join('');
}

function stepHtml(step) {
  const n = String(step.label).replace(/\D/g, '');
  return `<article class="gstep">
    <div class="gstep-rail"><span class="gstep-node">${esc(n)}</span><span class="gstep-line"></span></div>
    <div class="gstep-card${step.reverse ? ' reverse' : ''}">
      <div class="gstep-text">
        <p class="eyebrow">${esc(step.label)}</p>
        <h3>${esc(step.title)}</h3>
        <p class="muted">${esc(step.text)}</p>
        <ol class="steps-ol">${step.list.map(li => '<li>' + li + '</li>').join('')}</ol>
        ${step.note ? `<p class="note">${ICON.alert}<span>${step.note}</span></p>` : ''}
        ${step.info ? `<div class="info-box">
            <p class="info-title">${ICON.wifi}${esc(step.info[0])}</p>
            <p>${step.info[1]}</p>
          </div>` : ''}
      </div>
      <div class="gstep-media">${shotHtml(step)}</div>
    </div>
  </article>`;
}

function renderGuide(key) {
  const g = GUIDES[key];
  if (!g) return;
  const box = $('page-' + key);
  if (box.dataset.rendered === '1') return;   // static content, render once

  box.innerHTML = `
    <div class="int-card">
      <span class="int-logo">${g.icon}</span>
      <div class="int-main">
        <div class="int-title">
          <h2>${esc(g.name)}</h2>
          <span class="badge-official">AJEMS connector</span>
        </div>
        <p class="int-desc">${esc(g.blurb)}</p>
        <div class="int-meta">
          ${g.meta.map(m => `<div><span>${esc(m[0])}</span><b>${esc(m[1])}</b></div>`).join('')}
        </div>
      </div>
    </div>

    <div class="card guide-intro">
      <h3 style="font-size:15px;font-weight:640;margin-bottom:4px">Before you begin</h3>
      <p class="muted" style="margin-bottom:14px">Four things to have ready. It takes a couple of minutes.</p>
      <div class="req-grid">
        ${g.requirements.map(r => `<div class="req">${r[0]}<span>${esc(r[1])}</span></div>`).join('')}
      </div>
    </div>

    ${g.steps.map(stepHtml).join('')}

    <div class="cta">
      <div class="cta-main">
        <h3>${esc(g.cta[0])}</h3>
        <p>${esc(g.cta[1])}</p>
      </div>
      <button class="btn btn-white" data-guide-go="${esc(g.cta[2])}">Open connections</button>
    </div>`;

  box.querySelectorAll('[data-guide-go]').forEach(b =>
    b.addEventListener('click', () => goto(b.dataset.guideGo)));
  box.dataset.rendered = '1';
}

// ══════════ wizard ══════════

const APP_COLOURS = ['Cyan', 'Blue', 'Indigo', 'Purple', 'Teal',
                     'Green', 'Orange', 'Pink', 'Red', 'Gray'];
const LAST = 5;

function openWizard(task) {
  // Google is only needed for a Google Sheet. A task built from an uploaded
  // file must not ask for it.
  if (!STATUS.ajems.verified) { goto('connect'); return toast('Connect to AJEMS first.', 'err'); }
  if (!STATUS.signedIn && !UPLOADS.length) {
    goto('connect');
    return toast('Continue with Google, or upload an Excel sheet first.', 'err');
  }
  if (STATUS.signedIn && !STATUS.scopeOk && !UPLOADS.length) {
    goto('connect');
    return toast('Sign in with Google again — this session cannot read your sheets.', 'err');
  }

  W = task ? JSON.parse(JSON.stringify(task)) : null;
  WSTEP = 1; WCOLUMNS = {}; WFIELDS = {}; ROWCOUNTS = {}; WBUSY = false;
  EDITING = !!task;
  PICKED_APP = null;
  MAPTAB = null;
  NAMING = null; NAMEDRAFT = '';
  SRC = (task && task.source) || 'sheet';

  $('wizTitle').textContent = task ? 'Schedule' : 'New task';
  $('wName').value = task ? task.name : '';
  $('wSchedule').value = (task && task.schedule && task.schedule !== 'manual') ? task.schedule : 'hourly';
  $('wIdentity').value = task ? task.identity : 'hash';
  $('wUpdateExisting').checked = task ? !!task.updateExisting : true;
  $('wAnalysis').innerHTML = '';
  $('wMapping').innerHTML = '';
  $('mapTabs').innerHTML = '';
  ['wPreviewMsg', 'wAppMsg', 'wDestMsg', 'wMapMsg', 'wSyncMsg'].forEach(i => msg(i, ''));

  if (EDITING) {
    // Sheet, forms and mapping are already settled — only the schedule is
    // worth changing. Column names come from the saved mapping, so the key
    // chips can be built without re-reading anything.
    (W.selections || []).forEach(sel => {
      WCOLUMNS[sel.tab] = (sel.mapping || []).map(m => ({ column: m.column, samples: [] }));
    });
    updateIdentityHint();
    showStep(LAST);
    $('wizard').style.display = 'flex';
    return;
  }

  setWizardSource(SRC);
  $('tabField').style.display = 'none';
  $('wTabs').innerHTML = '';
  $('newAppPanel').style.display = 'none';
  $('btnNewApp').textContent = '+ Create new app';

  updateIdentityHint();
  updateScheduleForSource();
  showStep(1);
  $('wizard').style.display = 'flex';
  loadWorkspace();
}

function bindWizard() {
  $('wizClose').addEventListener('click', closeWizard);
  $('wizCancel').addEventListener('click', closeWizard);
  $('wizBack').addEventListener('click', () => showStep(WSTEP - 1));
  $('wizNext').addEventListener('click', onNext);

  document.querySelectorAll('.step').forEach(b =>
    b.addEventListener('click', () => {
      const n = parseInt(b.dataset.step, 10);
      if (n < WSTEP) showStep(n);          // going back is always safe
    }));

  $('btnPick').addEventListener('click', openPicker);
  $('uploadPick').addEventListener('change', e => chooseUpload(e.target.value));

  document.querySelectorAll('.srccard').forEach(b =>
    b.addEventListener('click', () => setWizardSource(b.dataset.src)));
  $('btnWizSignIn').addEventListener('click', () => {
    if (!STATUS.googleConfigured) {
      return msg('wPreviewMsg', 'Google sign-in is not set up on this server yet. ' +
        'Use an Excel or CSV file instead.', 'err');
    }
    location.href = '/auth/google';
  });
  $('btnWizUpload').addEventListener('click', () => $('wizFileInput').click());
  $('wizFileInput').addEventListener('change', e => {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (f) uploadFile(f, true);
  });

  $('btnNewApp').addEventListener('click', () => {
    const open = $('newAppPanel').style.display !== 'none';
    $('newAppPanel').style.display = open ? 'none' : '';
    $('btnNewApp').textContent = open ? '+ Create new app' : 'Cancel';
    if (!open) { buildColourChips(); $('naTitle').focus(); }
  });
  $('btnCreateApp').addEventListener('click', createApp);
  $('btnCreateAllForms').addEventListener('click', createAllForms);
  $('wIdentity').addEventListener('change', updateIdentityHint);
}

function closeWizard() { $('wizard').style.display = 'none'; }

function setNextLabel() {
  const btn = $('wizNext');
  // Step 2 has no Next: choosing an app is what moves things on.
  btn.style.display = (WSTEP === 2) ? 'none' : '';
  if (WBUSY) { btn.textContent = 'Working…'; btn.disabled = true; return; }
  btn.disabled = false;
  btn.textContent = WSTEP === LAST ? (EDITING ? 'Save and sync now' : 'Save task') : 'Next';
}

function showStep(n) {
  if (n < 1 || n > LAST) return;
  WSTEP = n;
  document.querySelectorAll('.wstep').forEach(x =>
    x.classList.toggle('active', +x.dataset.step === n));
  document.querySelectorAll('.step').forEach(x => {
    const i = +x.dataset.step;
    x.classList.toggle('active', i === n);
    x.classList.toggle('done', i < n);
    x.style.display = EDITING ? 'none' : '';    // no step bar when editing
  });
  $('wizBack').style.visibility = (n === 1 || EDITING) ? 'hidden' : '';
  setNextLabel();

  if (n === 2) loadWorkspace();      // builds the app grid once it has data
  if (n === 3) buildTabFormList();
  if (n === 4) buildMappingStep();
  if (n === 5) updateScheduleForSource();
}

async function onNext() {
  if (WBUSY) return;

  // ── step 1: read the tabs for them, then move on ──
  if (WSTEP === 1) {
    if (!W || !sourceId()) return toast('Choose a spreadsheet or an uploaded file.', 'err');
    const tabs = selectedTabs();
    if (!tabs.length) return toast('Pick at least one tab.', 'err');

    W.name = $('wName').value.trim() || (W.spreadsheetName + ' \u2192 AJEMS');
    W.sheetTabs = tabs;
    W.sheetTab = tabs[0];              // keeps older readers working

    // Keep any selection whose tab is still ticked; add rows for new tabs.
    const kept = (W.selections || []).filter(s => tabs.includes(s.tab));
    W.selections = tabs.map(tab =>
      kept.find(s => s.tab === tab) ||
      { tab, formId: '', formTitle: '', postUrl: '', getUrl: '', mapping: [] });

    const unread = tabs.filter(t => !WCOLUMNS[t]);
    if (unread.length) {
      WBUSY = true; setNextLabel();
      try { await refreshTabPreview(); }
      finally { WBUSY = false; setNextLabel(); }
    }
    if (!tabs.every(t => (WCOLUMNS[t] || []).length)) {
      return msg('wPreviewMsg', 'Some tabs have no readable columns — untick them to continue.', 'err');
    }
    return showStep(2);
  }

  // step 2 has no Next button — chooseApp() advances

  // ── step 3: every tab needs a form ──
  if (WSTEP === 3) {
    const missing = W.selections.filter(s => !s.formId);
    if (missing.length) {
      return msg('wDestMsg', missing.length + ' tab(s) still need a form: ' +
        missing.map(s => s.tab).join(', '), 'err');
    }
    return showStep(4);
  }

  // ── step 4: mapping ──
  if (WSTEP === 4) {
    collectMapping();
    const empty = W.selections.filter(s => !(s.mapping || []).some(m => m.column && m.fieldKey));
    if (empty.length) {
      return msg('wMapMsg', 'Nothing is mapped for: ' + empty.map(s => s.tab).join(', ') +
        '. Map at least one column per tab.', 'err');
    }
    return showStep(LAST);
  }

  // ── step 5: save ──
  W.schedule = $('wSchedule').value;
  W.identity = $('wIdentity').value;
  W.updateExisting = $('wUpdateExisting').checked;
  W.keyColumns = [...document.querySelectorAll('#wKeyCols .chip.on')].map(c => c.dataset.col);
  W.enabled = true;

  if (W.identity === 'key' && !W.keyColumns.length) {
    return toast('Pick at least one key column, or switch to whole-row matching.', 'err');
  }

  WBUSY = true; setNextLabel();
  try {
    const exists = TASKS.some(t => t.id === W.id);
    if (exists) await api('/api/tasks/' + W.id, { method: 'PUT', body: JSON.stringify(W) });
    else        await api('/api/tasks', { method: 'POST', body: JSON.stringify(W) });
    closeWizard();
    await loadTasks();
    await refreshStatus();
    goto('tasks');
    toast('Saved — syncing now, then ' + scheduleLabel(W.schedule).toLowerCase() + '.', 'ok');
    setTimeout(loadTasks, 2500);
  } catch (e) {
    msg('wSyncMsg', e.message, 'err');
  } finally {
    WBUSY = false; setNextLabel();
  }
}

// ══════════ step 1 — sheet and tabs ══════════

function ensureTask() {
  if (!W) {
    W = {
      id: 't_' + Date.now() + '_' + Math.floor(Math.random() * 9000 + 1000),
      name: '', enabled: true, source: SRC, uploadId: '',
      spreadsheetId: '', spreadsheetName: '',
      sheetTabs: [], sheetTab: '', appId: '', appTitle: '',
      selections: [], schedule: 'hourly',
      updateExisting: true, identity: 'hash', keyColumns: []
    };
  }
  return W;
}

/**
 * Switches step 1 between the two sources and shows whatever that source
 * needs: the Drive button, a sign-in prompt, or the upload controls.
 */
function setWizardSource(src) {
  SRC = src;
  document.querySelectorAll('.srccard').forEach(b => b.classList.toggle('on', b.dataset.src === src));

  const googleReady = STATUS.signedIn && STATUS.scopeOk;
  const sheet = src === 'sheet';

  $('btnPick').style.display      = (sheet && googleReady) ? '' : 'none';
  $('btnWizUpload').style.display = sheet ? 'none' : '';
  $('googleNeeded').style.display = (sheet && !googleReady) ? '' : 'none';
  $('pickedSheet').style.display  = (sheet && !googleReady) ? 'none' : '';

  ensureTask();
  W.source = src;

  // Switching source clears whatever the other one had chosen.
  if (sheet) { W.uploadId = ''; } else { W.spreadsheetId = ''; }
  W.spreadsheetName = '';
  W.selections = [];
  WCOLUMNS = {}; ROWCOUNTS = {};
  $('wAnalysis').innerHTML = '';
  $('wTabs').innerHTML = '';
  $('tabField').style.display = 'none';
  msg('wPreviewMsg', '');

  $('pickedSheet').classList.remove('on');
  $('pickedName').textContent = sheet ? 'No spreadsheet chosen' : 'No file chosen';
  $('pickedSub').textContent = sheet
    ? 'Pick one from your Google Drive.'
    : (UPLOADS.length ? 'Upload a file, or use one you uploaded before.'
                      : 'Upload an .xlsx, .xlsm, .xls or .csv from this computer.');
  $('btnPick').textContent = 'Choose from Google Drive';
  $('btnWizUpload').textContent = 'Upload Excel sheet';

  buildUploadPick();
  updateScheduleForSource();
}

/** Lists uploaded files next to the Drive button, when there are any. */
function buildUploadPick() {
  const sel = $('uploadPick');
  if (!UPLOADS.length || SRC !== 'excel') { sel.style.display = 'none'; return; }
  sel.style.display = '';
  sel.innerHTML = '<option value="">Or use an uploaded file</option>' +
    UPLOADS.map(u => `<option value="${esc(u.id)}">${esc(u.name)}</option>`).join('');
  if (W && W.source === 'excel' && W.uploadId) sel.value = W.uploadId;
}

async function chooseUpload(uploadId) {
  if (!uploadId) return;
  const up = UPLOADS.find(u => u.id === uploadId);
  if (!up) return;

  ensureTask();
  SRC = 'excel';
  W.source = 'excel';
  W.uploadId = up.id;
  W.spreadsheetId = '';
  W.spreadsheetName = up.name;
  W.selections = [];
  if (!$('wName').value) $('wName').value = up.name.replace(/\.[^.]+$/, '') + ' \u2192 AJEMS';

  WCOLUMNS = {}; ROWCOUNTS = {};
  $('wAnalysis').innerHTML = '';
  $('pickedName').textContent = up.name;
  $('pickedSub').textContent = 'Uploaded file. Imports once when the task is saved.';
  $('pickedSheet').classList.add('on');
  $('uploadPick').value = up.id;
  $('btnWizUpload').textContent = 'Upload a different file';
  updateScheduleForSource();
  msg('wPreviewMsg', '');

  await loadTabs(up.id);
}

/**
 * Google's own file dialog, with the views a Drive user expects: My Drive,
 * shared files, and the ones picked before. Picking is what grants this app
 * access to the file under drive.file — there is no other way in, and nothing
 * else in the user's Drive becomes readable.
 */
function openPicker() {
  if (!STATUS.apiKey) {
    return toast('GOOGLE_API_KEY is missing from .env — the picker cannot open.', 'err');
  }
  if (!STATUS.projectNumber) {
    return toast('The Cloud project number is unknown — set GOOGLE_PROJECT_NUMBER in .env.', 'err');
  }
  if (typeof gapi === 'undefined') {
    return toast('Google\u2019s picker script has not loaded yet — try again in a moment.', 'err');
  }

  gapi.load('picker', async () => {
    let token;
    try {
      token = (await api('/api/picker-token')).token;
    } catch (e) {
      return toast('Could not get a Google access token: ' + e.message, 'err');
    }

    const SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
    const view = (label, ownedByMe) => {
      const v = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS)
        .setMimeTypes(SHEET_MIME)
        .setMode(google.picker.DocsViewMode.LIST)   // thumbnails need a broader scope
        .setIncludeFolders(true)
        .setSelectFolderEnabled(false)
        .setLabel(label);
      if (ownedByMe !== undefined) v.setOwnedByMe(ownedByMe);
      return v;
    };

    new google.picker.PickerBuilder()
      .addView(view('My Drive', true))
      .addView(view('Shared with me', false))
      .addView(view('Recent'))
      .enableFeature(google.picker.Feature.NAV_HIDDEN, false)
      .setOAuthToken(token)
      .setDeveloperKey(STATUS.apiKey)
      // REQUIRED for drive.file. Without it the pick does not grant the app
      // access to the file, and every later read fails.
      .setAppId(STATUS.projectNumber)
      .setTitle('Choose a spreadsheet to connect')
      .setCallback(async data => {
        if (data.action !== google.picker.Action.PICKED) return;
        const doc = data.docs && data.docs[0];
        if (!doc) return;

        if (doc.mimeType && doc.mimeType !== SHEET_MIME) {
          return toast('"' + doc.name + '" is not a Google Sheet. Open it in Sheets and ' +
                       'save it as a Google Sheet first.', 'err');
        }

        ensureTask();
        SRC = 'sheet';
        W.source = 'sheet';
        W.uploadId = '';
        $('uploadPick').value = '';
        W.spreadsheetId = doc.id;
        W.spreadsheetName = doc.name;
        W.selections = [];
        if (!$('wName').value) $('wName').value = doc.name + ' \u2192 AJEMS';

        WCOLUMNS = {}; ROWCOUNTS = {};
        $('wAnalysis').innerHTML = '';
        $('pickedName').textContent = doc.name;
        $('pickedSub').textContent = 'Connected. Other sheets stay private.';
        $('pickedSheet').classList.add('on');
        $('btnPick').textContent = 'Change';
        msg('wPreviewMsg', '');

        await loadTabs(doc.id);
      })
      .build()
      .setVisible(true);
  });
}

async function loadTabs(spreadsheetId, preselect) {
  const box = $('wTabs');
  $('tabField').style.display = '';
  box.innerHTML = '<span class="hint">Loading tabs…</span>';
  try {
    const d = await api('/api/sheets/' + encodeURIComponent(spreadsheetId) +
      '/tabs?source=' + encodeURIComponent(SRC));
    if (!d.tabs.length) {
      box.innerHTML = '<span class="hint">This spreadsheet has no tabs.</span>';
      return;
    }
    const chosen = preselect || [];
    box.innerHTML = d.tabs.map((t, i) => {
      const on = chosen.length ? chosen.includes(t.title) : i === 0;
      return `<span class="chip${on ? ' on' : ''}" data-tab="${esc(t.title)}">${esc(t.title)}</span>`;
    }).join('');
    box.querySelectorAll('.chip').forEach(ch =>
      ch.addEventListener('click', async () => {
        ch.classList.toggle('on');
        await refreshTabPreview();
      }));
    // Show the first tab's columns straight away rather than making them
    // press Next to find out what was detected.
    refreshTabPreview();
    if (d.title && W && SRC === 'sheet') W.spreadsheetName = d.title;
  } catch (e) {
    box.innerHTML = '';
    const notFound = /not found|404/i.test(e.message);
    msg('wPreviewMsg', notFound
      ? 'Google says this file is not available to the app. Pick it again through ' +
        '"Choose from Google Drive" — that is what grants access.'
      : e.message, 'err');
  }
}

/** The id the server needs for the current source. */
function sourceId() { return SRC === 'excel' ? W.uploadId : W.spreadsheetId; }

function selectedTabs() {
  return [...document.querySelectorAll('#wTabs .chip.on')].map(c => c.dataset.tab);
}

/**
 * Reads whichever tabs are ticked and renders what was detected, right there
 * in step 1. Previously this only happened on Next, so the analysis flashed
 * past and was only ever seen by going Back — which is what made it look like
 * selecting a tab did nothing.
 */
let previewSeq = 0;
async function refreshTabPreview() {
  const tabs = selectedTabs();
  const mine = ++previewSeq;

  if (!tabs.length) {
    $('wAnalysis').innerHTML = '';
    return msg('wPreviewMsg', '');
  }

  const unread = tabs.filter(t => !WCOLUMNS[t]);
  if (unread.length) {
    msg('wPreviewMsg', 'Reading ' + unread.length + ' tab(s)…', 'info');
    try {
      await readTabColumns(unread);
    } catch (e) {
      if (mine === previewSeq) msg('wPreviewMsg', e.message, 'err');
      return;
    }
  }
  if (mine !== previewSeq) return;    // a later click already took over
  renderAnalysis(tabs);
}

/** Fetches column analysis for tabs not read yet. */
async function readTabColumns(tabs) {
  for (const tab of tabs) {
    const d = await api('/api/sheets/' + encodeURIComponent(sourceId()) +
      '/preview?source=' + encodeURIComponent(SRC) +
      '&tabs=' + encodeURIComponent(tab));
    WCOLUMNS[tab] = d.analysis;
    ROWCOUNTS[tab] = d.rowCount;
  }
}

function renderAnalysis(tabs) {
  const good = tabs.filter(t => (WCOLUMNS[t] || []).length);
  const empty = tabs.filter(t => !(WCOLUMNS[t] || []).length);

  if (empty.length) {
    msg('wPreviewMsg', 'No columns found in: ' + empty.join(', ') +
      '. Check that row 1 of each tab holds the column names.', 'err');
  } else {
    const rows = good.reduce((n, t) => n + (ROWCOUNTS[t] || 0), 0);
    msg('wPreviewMsg', good.length + ' tab(s), ' + rows + ' data row(s) in total.', 'ok');
  }

  $('wAnalysis').innerHTML = good.map(tab => `
    <div class="crumb"><b>${esc(tab)}</b> — ${WCOLUMNS[tab].length} column(s), ${ROWCOUNTS[tab] || 0} row(s)</div>
    <table><tr><th>Column</th><th>Detected type</th><th>Filled</th><th>Sample</th></tr>` +
    WCOLUMNS[tab].map(c => `<tr>
      <td><b>${esc(c.column)}</b></td>
      <td><span class="tag">${esc(c.field_type)}</span></td>
      <td>${c.filled}/${c.total}</td>
      <td class="mono">${esc(c.samples.join(' · '))}</td>
    </tr>`).join('') + '</table>').join('');
}

// ══════════ step 2 — application ══════════

async function loadWorkspace(force) {
  if (!WORKSPACE || force) {
    try { WORKSPACE = await api('/api/ajems/workspace'); }
    catch (e) { return msg('wAppMsg', e.message, 'err'); }
  }
  if (WSTEP === 2) buildAppGrid();
}

function buildAppGrid() {
  const all = (WORKSPACE && WORKSPACE.apps) || [];
  const apps = all.filter(a => a.id);
  const broken = all.filter(a => !a.id);
  if (broken.length) {
    msg('wAppMsg', broken.length + ' application(s) came back from AJEMS without an id and ' +
      'cannot be used: ' + broken.map(a => a.title).join(', '), 'err');
  }

  $('appSub').textContent = apps.length + ' existing app' + (apps.length === 1 ? '' : 's') +
    ' — choose one, or create a new one.';

  $('appGrid').innerHTML = apps.map(a => `
    <button class="pcard${W && String(W.appId) === String(a.id) ? ' on' : ''}" data-app="${esc(a.id)}">
      <span class="pcard-ico">${esc(initials(a.title))}</span>
      <span>
        <span class="pcard-name">${esc(a.title)}</span>
        <span class="pcard-sub">${a.forms.length} form${a.forms.length === 1 ? '' : 's'}</span>
      </span>
    </button>`).join('') ||
    '<span class="hint">No applications in this workspace yet — create one.</span>';

  $('appGrid').querySelectorAll('[data-app]').forEach(b =>
    b.addEventListener('click', () => chooseApp(b.dataset.app)));
}

function initials(name) {
  return String(name || '?').replace(/[^A-Za-z0-9 ]/g, '').split(/\s+/)
    .filter(Boolean).slice(0, 2).map(w => w[0].toUpperCase()).join('') || '?';
}

function chooseApp(appId) {
  const app = WORKSPACE.apps.find(a => String(a.id) === String(appId));
  if (!app) return;
  ensureTask();
  if (String(W.appId) !== String(app.id)) {
    // A different app means the forms chosen before no longer exist here.
    W.selections.forEach(s => {
      s.formId = ''; s.formTitle = ''; s.postUrl = ''; s.getUrl = ''; s.mapping = [];
    });
    WFIELDS = {};
  }
  PICKED_APP = app;
  W.appId = app.id;
  W.appTitle = app.title;
  showStep(3);
}

function buildColourChips() {
  $('naColour').innerHTML = APP_COLOURS.map((c, i) =>
    `<span class="chip${i === 0 ? ' on' : ''}" data-colour="${c.toLowerCase()}">${c}</span>`).join('');
  $('naColour').querySelectorAll('.chip').forEach(ch =>
    ch.addEventListener('click', () => {
      $('naColour').querySelectorAll('.chip').forEach(x => x.classList.remove('on'));
      ch.classList.add('on');
    }));
}

async function createApp() {
  if (WBUSY) return;
  const title = $('naTitle').value.trim();
  if (!title) return msg('wAppMsg', 'Give the app a name.', 'err');
  const chip = $('naColour').querySelector('.chip.on');

  WBUSY = true; setNextLabel();
  msg('wAppMsg', 'Creating the app in AJEMS…', 'info');
  try {
    const created = await api('/api/ajems/create-app', {
      method: 'POST',
      body: JSON.stringify({
        title,
        description: $('naDesc').value.trim(),
        icon: $('naIcon').value,
        color: chip ? chip.dataset.colour : 'cyan'
      })
    });

    WORKSPACE = await api('/api/ajems/workspace');
    const newId = created && (created.app_id || created.id || created._id);
    const app = WORKSPACE.apps.find(a => String(a.id) === String(newId)) ||
                WORKSPACE.apps.find(a => a.title === title);
    if (!app || !app.id) {
      throw new Error('The app was created but AJEMS did not return a usable id. ' +
                      'Refresh and pick it from the list.');
    }

    $('newAppPanel').style.display = 'none';
    $('btnNewApp').textContent = '+ Create new app';
    $('naTitle').value = ''; $('naDesc').value = '';
    msg('wAppMsg', '');

    // A brand new app has no forms, so every tab will need one creating.
    chooseApp(app.id);
  } catch (e) {
    msg('wAppMsg', e.message, 'err');
  } finally {
    WBUSY = false; setNextLabel();
  }
}

// ══════════ step 3 — one form per tab ══════════

function buildTabFormList() {
  const app = PICKED_APP ||
    (WORKSPACE && WORKSPACE.apps.find(a => String(a.id) === String(W.appId)));
  if (!app) return showStep(2);
  PICKED_APP = app;

  const forms = app.forms || [];
  const done = W.selections.filter(s => s.formId).length;
  $('formSub').innerHTML = 'Into <b>' + esc(app.title) + '</b> — each tab goes into its own form. ' +
    done + ' of ' + W.selections.length + ' set.';

  $('tabFormList').innerHTML = W.selections.map((sel, i) => {
    const cols = (WCOLUMNS[sel.tab] || []).length;
    const opts = '<option value="">— choose a form —</option>' +
      forms.map(f => `<option value="${esc(f.id)}"${String(f.id) === String(sel.formId) ? ' selected' : ''}>` +
        `${esc(f.title)} (${f.fieldCount} fields)</option>`).join('');
    const naming = NAMING === i;
    const pick = naming
      ? `<input type="text" data-name="${i}" value="${esc(NAMEDRAFT || sel.tab)}" placeholder="Form name">
         <button class="btn btn-primary btn-sm" data-mkform="${i}">Create</button>
         <button class="btn btn-ghost btn-sm" data-cancelname="1">Cancel</button>`
      : `<select data-sel="${i}">${opts}</select>
         <button class="btn btn-ghost btn-sm" data-newform="${i}">+ New</button>`;
    return `<div class="tfrow${sel.formId ? ' done' : ''}">
      <div class="tf-tab">${esc(sel.tab)}<span>${cols} column(s)</span></div>
      <div class="arrow">&rarr;</div>
      <div class="tf-pick">${pick}</div>
    </div>`;
  }).join('');

  $('tabFormList').querySelectorAll('select[data-sel]').forEach(sel =>
    sel.addEventListener('change', () => chooseForm(parseInt(sel.dataset.sel, 10), sel.value)));

  // "+ New" reveals a name box for that row, pre-filled with the tab name.
  $('tabFormList').querySelectorAll('[data-newform]').forEach(b =>
    b.addEventListener('click', () => {
      NAMING = parseInt(b.dataset.newform, 10);
      NAMEDRAFT = '';
      buildTabFormList();
      const input = $('tabFormList').querySelector(`input[data-name="${NAMING}"]`);
      if (input) { input.focus(); input.select(); }
    }));
  $('tabFormList').querySelectorAll('[data-cancelname]').forEach(b =>
    b.addEventListener('click', () => { NAMING = null; NAMEDRAFT = ''; buildTabFormList(); }));
  $('tabFormList').querySelectorAll('input[data-name]').forEach(inp => {
    inp.addEventListener('input', () => { NAMEDRAFT = inp.value; });
    inp.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); createFormForTab(parseInt(inp.dataset.name, 10), false, inp.value); }
      if (e.key === 'Escape') { NAMING = null; NAMEDRAFT = ''; buildTabFormList(); }
    });
  });
  $('tabFormList').querySelectorAll('[data-mkform]').forEach(b =>
    b.addEventListener('click', () => {
      const i = parseInt(b.dataset.mkform, 10);
      const inp = $('tabFormList').querySelector(`input[data-name="${i}"]`);
      createFormForTab(i, false, inp ? inp.value : '');
    }));
}

async function chooseForm(index, formId) {
  const sel = W.selections[index];
  if (!sel) return;
  if (!formId) {
    sel.formId = ''; sel.formTitle = ''; sel.postUrl = ''; sel.getUrl = ''; sel.mapping = [];
    return buildTabFormList();
  }
  const form = (PICKED_APP.forms || []).find(f => String(f.id) === String(formId));
  if (!form) return;

  WBUSY = true; setNextLabel();
  msg('wDestMsg', 'Reading "' + form.title + '"…', 'info');
  try {
    const d = await api('/api/ajems/form?detailUrl=' + encodeURIComponent(form.detailUrl));
    const fields = d.fields.filter(f => f.key);
    if (!fields.length) throw new Error(`"${form.title}" has no fields to map to.`);

    WFIELDS[form.id] = fields;
    sel.formId = form.id;
    sel.formTitle = form.title;
    sel.postUrl = form.responsesUrl;
    sel.getUrl = form.responsesUrl;
    sel.mapping = autoMap(sel.tab, fields, sel.mapping);

    msg('wDestMsg', '');
    buildTabFormList();
  } catch (e) {
    msg('wDestMsg', e.message, 'err');
  } finally {
    WBUSY = false; setNextLabel();
  }
}

/** Create one form from a tab's own columns. */
async function createFormForTab(index, silent, name) {
  const sel = W.selections[index];
  if (!sel) return false;
  const cols = WCOLUMNS[sel.tab] || [];
  if (!cols.length) { msg('wDestMsg', 'No columns were read for "' + sel.tab + '".', 'err'); return false; }

  const title = String(name === undefined ? sel.tab : name).trim() || sel.tab;
  if (!silent) msg('wDestMsg', 'Creating "' + title + '"…', 'info');
  try {
    const d = await api('/api/ajems/create-form', {
      method: 'POST',
      body: JSON.stringify({
        source: SRC,
        spreadsheetId: W.spreadsheetId,
        uploadId: W.uploadId,
        tabs: [sel.tab],
        appId: PICKED_APP.id,
        title,
        columns: cols.map(c => c.column)
      })
    });

    WORKSPACE = await api('/api/ajems/workspace');
    const app = WORKSPACE.apps.find(a => String(a.id) === String(PICKED_APP.id));
    if (!app) throw new Error('The form was created but its application could not be found again.');
    PICKED_APP = app;

    const newId = d.form && (d.form.form_id || d.form.id || d.form._id);
    const form = (app.forms || []).find(f => String(f.id) === String(newId)) ||
                 (app.forms || []).find(f => f.title === title);
    if (!form || !form.id) {
      throw new Error('The form was created, but AJEMS did not return a usable id for it.');
    }

    const fields = d.fields.map(f => ({ key: f.key, label: f.label, field_type: f.field_type }));
    WFIELDS[form.id] = fields;
    sel.formId = form.id;
    sel.formTitle = form.title;
    sel.postUrl = form.responsesUrl;
    sel.getUrl = form.responsesUrl;
    sel.mapping = autoMap(sel.tab, fields, []);

    NAMING = null; NAMEDRAFT = '';
    if (!silent) { msg('wDestMsg', 'Created "' + form.title + '".', 'ok'); buildTabFormList(); }
    return true;
  } catch (e) {
    msg('wDestMsg', e.message, 'err');
    return false;
  }
}

async function createAllForms() {
  if (WBUSY) return;
  const todo = W.selections.map((s, i) => ({ s, i })).filter(x => !x.s.formId);
  if (!todo.length) return msg('wDestMsg', 'Every tab already has a form.', 'info');

  WBUSY = true; setNextLabel();
  msg('wDestMsg', 'Creating ' + todo.length + ' form(s)…', 'info');
  let ok = 0;
  try {
    for (const x of todo) {
      if (await createFormForTab(x.i, true)) ok++;
    }
  } finally {
    WBUSY = false; setNextLabel();
  }
  buildTabFormList();
  if (ok === todo.length) msg('wDestMsg', 'Created ' + ok + ' form(s), one per tab.', 'ok');
}

// ══════════ step 4 — mapping, always shown ══════════

function normLabel(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

/**
 * Match a tab's columns to a form's fields by label. A form created from the
 * tab matches perfectly; an existing form usually matches most of them. The
 * result is always shown in step 4 so any of it can be changed.
 */
function autoMap(tab, fields, saved) {
  const cols = WCOLUMNS[tab] || [];
  const used = new Set();
  const out = [];

  cols.forEach(col => {
    const prev = (saved || []).find(m => m.column === col.column && m.fieldKey);
    let field = prev ? fields.find(f => f.key === prev.fieldKey) : null;
    if (!field) {
      field = fields.find(f => !used.has(f.key) && normLabel(f.label) === normLabel(col.column));
    }
    if (field) {
      used.add(field.key);
      out.push({ column: col.column, fieldKey: field.key, label: field.label, fieldType: field.field_type });
    } else {
      out.push({ column: col.column, fieldKey: '', label: '', fieldType: '' });
    }
  });
  return out;
}

function buildMappingStep() {
  if (!W.selections.length) return showStep(1);

  // One tab: no tab bar. Several: a bar to switch between them.
  if (W.selections.length > 1) {
    if (!MAPTAB || !W.selections.some(s => s.tab === MAPTAB)) MAPTAB = W.selections[0].tab;
    $('mapTabs').innerHTML = W.selections.map(s => {
      const mapped = (s.mapping || []).filter(m => m.fieldKey).length;
      const total = (WCOLUMNS[s.tab] || []).length;
      return `<button class="tab${s.tab === MAPTAB ? ' active' : ''}" data-maptab="${esc(s.tab)}">` +
             `${esc(s.tab)} <span class="tag grey">${mapped}/${total}</span></button>`;
    }).join('');
    $('mapTabs').querySelectorAll('[data-maptab]').forEach(b =>
      b.addEventListener('click', () => {
        collectMapping();
        MAPTAB = b.dataset.maptab;
        buildMappingStep();
      }));
  } else {
    MAPTAB = W.selections[0].tab;
    $('mapTabs').innerHTML = '';
  }

  const sel = W.selections.find(s => s.tab === MAPTAB);
  const fields = WFIELDS[sel.formId] || [];
  const cols = WCOLUMNS[sel.tab] || [];

  if (!fields.length) {
    $('wMapping').innerHTML = '<div class="msg err">No fields loaded for "' +
      esc(sel.formTitle || sel.tab) + '". Go back and pick its form again.</div>';
    return;
  }

  $('wMapping').innerHTML =
    `<div class="crumb"><b>${esc(sel.tab)}</b> &rarr; ${esc(sel.formTitle)}</div>` +
    '<div class="map-row"><div class="lbl">Sheet column</div><div></div><div class="lbl">AJEMS field</div></div>' +
    cols.map(col => {
      const m = (sel.mapping || []).find(x => x.column === col.column) || {};
      const sample = (col.samples || []).slice(0, 2).join(' · ');
      return `<div class="map-row">
        <div><b>${esc(col.column)}</b>${sample ? `<div class="sample">${esc(sample)}</div>` : ''}</div>
        <div class="arrow">&rarr;</div>
        <div><select data-col="${esc(col.column)}">
          <option value="">— do not send —</option>
          ${fields.map(f => `<option value="${esc(f.key)}"${f.key === m.fieldKey ? ' selected' : ''}>` +
            `${esc(f.label)} (${esc(f.field_type || '?')})</option>`).join('')}
        </select></div>
      </div>`;
    }).join('');

  $('wMapping').querySelectorAll('select[data-col]').forEach(s =>
    s.addEventListener('change', () => { collectMapping(); refreshMapCounts(); }));

  const mapped = (sel.mapping || []).filter(m => m.fieldKey).length;
  msg('wMapMsg', mapped + ' of ' + cols.length + ' column(s) mapped for "' + sel.tab + '".',
      mapped ? 'ok' : 'err');
}

function refreshMapCounts() {
  W.selections.forEach(s => {
    const btn = $('mapTabs').querySelector(`[data-maptab="${CSS.escape(s.tab)}"] .tag`);
    if (btn) {
      btn.textContent = (s.mapping || []).filter(m => m.fieldKey).length +
                        '/' + (WCOLUMNS[s.tab] || []).length;
    }
  });
}

function collectMapping() {
  const sel = W.selections.find(s => s.tab === MAPTAB);
  if (!sel) return;
  const fields = WFIELDS[sel.formId] || [];
  document.querySelectorAll('#wMapping select[data-col]').forEach(s => {
    let entry = (sel.mapping || []).find(m => m.column === s.dataset.col);
    if (!entry) { entry = { column: s.dataset.col }; sel.mapping.push(entry); }
    const field = fields.find(f => f.key === s.value);
    entry.fieldKey = s.value;
    entry.label = field ? field.label : '';
    entry.fieldType = field ? field.field_type : '';
  });
}

// ══════════ step 5 — schedule ══════════

/** Every column across every tab, for the key-column chips. */
function allColumns() {
  const out = [];
  Object.keys(WCOLUMNS).forEach(tab => {
    (WCOLUMNS[tab] || []).forEach(c => { if (!out.includes(c.column)) out.push(c.column); });
  });
  return out;
}

/** Uploaded files import once; only Google Sheets have something to poll. */
function updateScheduleForSource() {
  const upload = SRC === 'excel';
  $('wSchedule').disabled = upload;
  const hint = $('wSchedule').parentElement.querySelector('.hint');
  if (hint) {
    hint.textContent = upload
      ? 'An uploaded file only changes when you upload it again, so this task imports once on save. Use Sync now to repeat it.'
      : 'Runs on its own. Each check is one small Drive call; the sheet is only read when it actually changed.';
  }
}

function updateIdentityHint() {
  const key = $('wIdentity').value === 'key';
  $('keyColsField').style.display = key ? '' : 'none';
  $('identityHint').textContent = key
    ? 'Edits to any other column are recognised as an update to the same AJEMS record.'
    : 'Good for append-only sheets. An edited row looks like a new row, so it is inserted again.';
  if (key) buildKeyChips();
}

function buildKeyChips() {
  const selected = (W && W.keyColumns) || [];
  $('wKeyCols').innerHTML = allColumns().map(c =>
    `<span class="chip${selected.includes(c) ? ' on' : ''}" data-col="${esc(c)}">${esc(c)}</span>`
  ).join('');
  $('wKeyCols').querySelectorAll('.chip').forEach(ch =>
    ch.addEventListener('click', () => ch.classList.toggle('on')));
}
