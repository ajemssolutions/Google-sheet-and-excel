require('./lib/env');

const path = require('path');
const express = require('express');

const store = require('./lib/store');
const google = require('./lib/google');
const excel = require('./lib/excel');
const ajems = require('./lib/ajems');
const detect = require('./lib/detect');
const engine = require('./lib/engine');
const scheduler = require('./lib/scheduler');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '30mb' }));   // uploads arrive as base64 JSON
app.use(express.static(path.join(__dirname, 'public')));

// Small wrapper so a thrown error becomes a clean JSON response.
const wrap = fn => (req, res) => {
  Promise.resolve(fn(req, res)).catch(err => {
    console.error(err);
    res.status(err.status && err.status < 600 ? err.status : 500)
       .json({ error: err.message || String(err) });
  });
};

// ══════════════════════════════════════════════════════════
// Status
// ══════════════════════════════════════════════════════════

app.get('/api/status', wrap(async (req, res) => {
  const s = store.get();
  const c = google.cfg();
  const scope = google.scopeStatus();
  const credProblems = google.credentialProblems();
  res.json({
    // credProblems is deliberately NOT sent: it names environment variables,
    // which is developer detail, not something a user should read. The list
    // is printed to the terminal at startup instead.
    scopeOk: scope.ok,
    neededScope: scope.needed,
    grantedScopes: scope.granted,
    googleConfigured: credProblems.length === 0,
    redirectUri: c.redirectUri,
    signedIn: !!(s.google && s.google.access_token),
    email: s.google ? s.google.email : '',
    scopes: s.google ? s.google.scope : '',
    clientId: c.clientId,
    apiKey: c.apiKey,
    projectNumber: c.projectNumber,
    ajems: {
      baseUrl: s.ajems.baseUrl,
      tenant: s.ajems.tenant,
      verified: s.ajems.verified,
      hasKey: !!s.ajems.secretKey
    },
    taskCount: (s.tasks || []).filter(t => !t.tenant || t.tenant === (s.ajems || {}).tenant).length,
    uploadCount: (s.uploads || []).length
  });
}));



// ══════════════════════════════════════════════════════════
// Google OAuth
// ══════════════════════════════════════════════════════════

app.get('/auth/google', (req, res) => {
  const problems = google.credentialProblems();
  if (problems.length) {
    const c = google.cfg();
    return res.status(500).send(`<!doctype html>
<meta charset="utf-8">
<title>Google credentials problem</title>
<style>
 body{font-family:"Segoe UI",system-ui,sans-serif;max-width:720px;margin:60px auto;padding:0 24px;color:#0F172A;line-height:1.6}
 h2{font-size:20px;margin-bottom:6px} code{background:#F1F5F9;padding:2px 6px;border-radius:4px;font-size:13px}
 li{margin-bottom:8px} .box{background:#FFFBEB;border:1px solid #FDE68A;border-radius:10px;padding:14px 18px;margin:18px 0}
 a{color:#2563EB}
</style>
<h2>Google sign-in is not configured correctly</h2>
<p>Fixing this before sending you to Google, because Google's own page only says
<code>Error 401: invalid_client</code> without saying which value is wrong.</p>
<div class="box"><ul>${problems.map(p => '<li>' + p + '</li>').join('')}</ul></div>
<p>Currently loaded from <code>.env</code>:</p>
<ul>
  <li>Client ID: <code>${c.clientId ? c.clientId.slice(0, 24) + '…' : '(empty)'}</code></li>
  <li>Client secret: <code>${c.clientSecret ? '(set, ' + c.clientSecret.length + ' chars)' : '(empty)'}</code></li>
  <li>Redirect URI: <code>${c.redirectUri}</code></li>
</ul>
<p>Edit <code>.env</code>, then <b>restart the server</b> — it only reads the file at startup.</p>
<p><a href="/">Back to the app</a></p>`);
  }
  res.redirect(google.authUrl());
});

app.get('/oauth/callback', wrap(async (req, res) => {
  if (req.query.error) {
    return res.send(`<h2>Google sign-in was cancelled</h2><p>${req.query.error}</p>
      <p><a href="/">Back to the app</a></p>`);
  }
  if (!req.query.code) return res.status(400).send('No authorization code returned.');

  const session = await google.exchangeCode(req.query.code);
  store.log('ok', '', 'Signed in with Google as ' + (session.email || 'unknown'));
  res.redirect('/?signedin=1');
}));

/**
 * The Google Picker runs in the browser and needs a live OAuth token.
 * Localhost prototype only — a hosted version would mint a short-lived token
 * scoped to the Picker rather than handing over the session token.
 */
app.get('/api/picker-token', wrap(async (req, res) => {
  res.json({ token: await google.token() });
}));

app.post('/api/google/signout', wrap(async (req, res) => {
  google.signOut();
  res.json({ ok: true });
}));

// ══════════════════════════════════════════════════════════
// Sheets
// ══════════════════════════════════════════════════════════

app.get('/api/sheets', wrap(async (req, res) => {
  res.json({ files: await google.listSpreadsheets() });
}));

app.get('/api/sheets/:id/tabs', wrap(async (req, res) => {
  if (req.query.source === 'excel') return res.json(excel.listTabs(req.params.id));
  res.json(await google.listTabs(req.params.id));
}));

// ══════════════════════════════════════════════════════════
// Uploaded Excel and CSV files
// ══════════════════════════════════════════════════════════

/**
 * The file arrives as base64 in the JSON body rather than multipart, which
 * keeps the dependency list short. Uploads are shared: several tasks can use
 * the same file, so removing a task does not remove it.
 */
app.post('/api/uploads', wrap(async (req, res) => {
  const { name, data } = req.body || {};
  if (!name || !data) throw new Error('Send both a file name and its contents.');

  const id = excel.save(name, data);
  const info = excel.listTabs(id);
  const upload = {
    id, name,
    tabs: info.tabs.map(t => t.title),
    at: new Date().toISOString()
  };
  store.addUpload(upload);
  store.log('ok', '', 'Uploaded ' + name + ' (' + upload.tabs.length + ' sheet(s)).');
  res.json(upload);
}));

app.get('/api/uploads', wrap(async (req, res) => {
  res.json({ uploads: store.get().uploads || [] });
}));

app.delete('/api/uploads/:id', wrap(async (req, res) => {
  const inUse = store.get().tasks.filter(t => t.uploadId === req.params.id);
  if (inUse.length) {
    throw new Error('That file is used by ' + inUse.length + ' task(s): ' +
      inUse.map(t => t.name).join(', ') + '. Delete them first.');
  }
  excel.remove(req.params.id);
  store.removeUpload(req.params.id);
  res.json({ ok: true });
}));

/**
 * Reads one or more tabs and reports the columns found across all of them.
 * The header is always row 1 — every sheet we have seen works that way, and
 * asking the user was one more decision for no benefit.
 */
async function readTabs(source, id, tabs) {
  const headers = [];
  const rows = [];
  for (const tab of tabs) {
    const r = source === 'excel' ? excel.readTab(id, tab, 1) : await google.readTab(id, tab, 1);
    r.headers.forEach(h => { if (!headers.includes(h)) headers.push(h); });
    r.rows.forEach(row => { row.__tab = tab; rows.push(row); });
  }
  return { headers, rows };
}

function tabsFromQuery(q) {
  const raw = q.tabs || q.tab || '';
  return String(raw).split('\u001f').map(t => t.trim()).filter(Boolean);
}

app.get('/api/sheets/:id/preview', wrap(async (req, res) => {
  const tabs = tabsFromQuery(req.query);
  if (!tabs.length) throw new Error('No tab selected.');
  const { headers, rows } = await readTabs(req.query.source, req.params.id, tabs);
  res.json({
    tabs,
    headers,
    rowCount: rows.length,
    analysis: detect.analyse(headers, rows),
    sample: rows.slice(0, 5)
  });
}));


/**
 * Normalise workspace_config into what the UI needs.
 *
 * The API names these `app_id` and `form_id` — NOT `id`. Reading the wrong
 * key yields an option with an empty value, and the UI then says "choose an
 * application first" while an application is visibly selected. The fallbacks
 * are only there in case the shape ever changes.
 */
function shapeApps(base, apps) {
  return (apps || []).map(a => ({
    id: a.app_id || a.id || a._id || '',
    title: a.title || a.name || '(untitled app)',
    forms: (a.forms || []).map(f => {
      const u = ajems.formUrls(base, a, f);
      return {
        id: u.formId || '',
        title: f.title || f.name || '(untitled form)',
        detailUrl: u.detail,
        responsesUrl: u.responses,
        fieldCount: (f.fields || []).length
      };
    })
  }));
}

// ══════════════════════════════════════════════════════════
// AJEMS connection
// ══════════════════════════════════════════════════════════

app.post('/api/ajems/connect', wrap(async (req, res) => {
  const baseUrl = String(req.body.baseUrl || '').trim();
  const secretKey = String(req.body.secretKey || '').trim();
  if (!baseUrl || !secretKey) throw new Error('Both the workspace URL and secret key are required.');

  const { base, data } = await ajems.workspaceConfig(baseUrl, secretKey);

  store.set({
    ajems: { baseUrl: base, secretKey, tenant: data.tenant, verified: true }
  });
  store.log('ok', '', `Connected to AJEMS workspace "${data.tenant}" — ${data.apps.length} app(s).`);

  res.json({ tenant: data.tenant, base, apps: shapeApps(base, data.apps) });
}));

app.get('/api/ajems/workspace', wrap(async (req, res) => {
  const a = store.get().ajems;
  if (!a.verified) throw new Error('Not connected to AJEMS yet.');
  const { base, data } = await ajems.workspaceConfig(a.baseUrl, a.secretKey);
  res.json({ tenant: data.tenant, apps: shapeApps(base, data.apps) });
}));

/** Fields of an existing form, used to auto-map by label. */
app.get('/api/ajems/form', wrap(async (req, res) => {
  const detailUrl = req.query.detailUrl;
  if (!detailUrl) throw new Error('detailUrl is required.');
  const form = await ajems.getForm(detailUrl);
  const fields = form.fields || form.form_fields || [];
  res.json({
    title: form.title || '',
    fields: fields.map(f => ({
      key: f.key, label: f.label, field_type: f.field_type
    })),
    labels: form.labels || null
  });
}));

app.post('/api/ajems/create-app', wrap(async (req, res) => {
  const a = store.get().ajems;
  const title = String(req.body.title || '').trim();
  if (!title) throw new Error('The app needs a name.');
  const created = await ajems.createApp(a.baseUrl, {
    title,
    description: req.body.description || '',
    icon: req.body.icon || 'pi pi-align-justify',
    color: req.body.color || 'cyan'
  });
  store.log('ok', '', 'Created AJEMS app: ' + title);
  res.json(created);
}));

app.post('/api/ajems/create-form', wrap(async (req, res) => {
  const a = store.get().ajems;
  const { spreadsheetId, uploadId, source, appId, title, columns } = req.body;
  const tabs = Array.isArray(req.body.tabs) ? req.body.tabs
             : (req.body.tab ? [req.body.tab] : []);
  if (!tabs.length) throw new Error('No tab selected.');

  const { headers, rows } = await readTabs(source, source === 'excel' ? uploadId : spreadsheetId, tabs);
  const wanted = Array.isArray(columns) && columns.length
    ? headers.filter(h => columns.includes(h))
    : headers;

  const fields = detect.buildFields(wanted, rows);
  const created = await ajems.createForm(a.baseUrl, appId, fields, { title });

  store.log('ok', '', `Created AJEMS form "${title}" with ${fields.length} field(s).`);
  res.json({ form: created, fields });
}));

// ══════════════════════════════════════════════════════════
// Tasks
// ══════════════════════════════════════════════════════════

/**
 * Run a task just after the response goes out, so saving a task syncs it
 * immediately instead of leaving it "waiting for its first run" until the
 * schedule comes round. Deliberately not awaited: a large sheet must not
 * hold up the Save button.
 */
function readyToSync(t) {
  if (!t.enabled) return false;
  const hasSource = t.source === 'excel' ? !!t.uploadId : !!t.spreadsheetId;
  const hasTarget = (t.selections || []).some(s => s.postUrl);
  return hasSource && hasTarget;
}

function syncSoon(taskId, trigger) {
  setTimeout(() => {
    engine.runTask(taskId, trigger, { seedFirst: true, concurrency: 4 })
      .catch(e => store.log('error', '', 'Sync after save failed: ' + e.message));
  }, 200);
}

/**
 * The connector holds one AJEMS workspace at a time, and a secret key only
 * works for its own tenant. Tasks are therefore tagged with the tenant that
 * created them and only that tenant's tasks are listed or run - reconnecting
 * to an organisation brings its tasks back.
 */
function activeTenant() {
  return (store.get().ajems || {}).tenant || '';
}

function tasksForTenant() {
  const tenant = activeTenant();
  if (!tenant) return [];
  const all = store.get().tasks;

  // Tasks saved before organisations existed are adopted by the first
  // workspace that opens them, so nothing is stranded.
  let adopted = false;
  all.forEach(t => { if (!t.tenant) { t.tenant = tenant; adopted = true; } });
  if (adopted) store.save();

  return all.filter(t => t.tenant === tenant);
}

app.get('/api/tasks', wrap(async (req, res) => {
  res.json({ tasks: tasksForTenant(), tenant: activeTenant(), status: scheduler.status() });
}));

app.post('/api/tasks', wrap(async (req, res) => {
  if (!activeTenant()) throw new Error('Connect an AJEMS workspace before saving a task.');
  const t = store.newTask();
  Object.assign(t, req.body || {}, { tenant: activeTenant() });
  store.upsertTask(t);
  res.json(t);
  if (readyToSync(t)) syncSoon(t.id, 'first run');
}));

app.put('/api/tasks/:id', wrap(async (req, res) => {
  const t = store.getTask(req.params.id);
  if (!t) throw new Error('Task not found.');
  const before = JSON.stringify(t.mapping);
  Object.assign(t, req.body || {}, { id: t.id, tenant: t.tenant || activeTenant() });
  // A changed mapping invalidates the content hashes we stored.
  if (JSON.stringify(t.mapping) !== before) store.clearLinks(t.id);
  if (t.enabled) { t.pausedReason = ''; t.consecutiveFailures = 0; }
  store.upsertTask(t);
  res.json(t);
  if (readyToSync(t)) syncSoon(t.id, 'saved');
}));

app.delete('/api/tasks/:id', wrap(async (req, res) => {
  store.removeTask(req.params.id);
  res.json({ ok: true });
}));

app.post('/api/tasks/:id/run', wrap(async (req, res) => {
  const result = await engine.runTask(req.params.id, req.body.trigger || 'manual', {
    seedFirst: !!req.body.seedFirst,
    concurrency: parseInt(req.body.concurrency, 10) || 4
  });
  res.json(result);
}));

/** Dry run: what would be sent for the first few rows of each tab. */
app.post('/api/tasks/:id/preview-payload', wrap(async (req, res) => {
  const task = store.getTask(req.params.id);
  if (!task) throw new Error('Task not found.');
  const links = store.getLinks(task.id);
  const out = [];

  for (const sel of engine.selectionsOf(task)) {
    if (!sel.tab) continue;
    const { rows } = await engine.readSourceTab(task, sel.tab);
    rows.slice(0, 3).forEach(row => {
      row.__tab = sel.tab;
      const key = engine.rowKey(task, row, sel);
      const existing = key ? links[key] : null;
      const hash = engine.contentHash(task, row, sel);
      out.push({
        tab: sel.tab,
        form: sel.formTitle,
        sheetRow: row.__rowNumber,
        payload: engine.buildPayload(task, row, sel),
        action: !existing ? 'create'
              : existing.hash === hash ? 'unchanged'
              : (task.updateExisting ? 'update' : 'skip (updates off)'),
        linkedTo: existing ? existing.responseId : null
      });
    });
  }

  res.json({ preview: out });
}));

/**
 * Replaces the file behind an Excel task and syncs it.
 *
 * The dedupe key is built from the tab name and the row's own values, never
 * from the file, so the task's history survives the swap: rows already sent
 * are recognised in the new file and only genuinely new or changed ones go
 * to AJEMS.
 */
app.post('/api/tasks/:id/replace-file', wrap(async (req, res) => {
  const task = store.getTask(req.params.id);
  if (!task) throw new Error('Task not found.');
  if (task.source !== 'excel') throw new Error('This task reads a Google Sheet, not a file.');

  const { name, data } = req.body || {};
  if (!name || !data) throw new Error('Send both a file name and its contents.');

  const newId = excel.save(name, data);

  // The new file must still contain every tab this task maps, or the sync
  // would silently do nothing for the missing ones.
  const tabs = excel.listTabs(newId).tabs.map(t => t.title);
  const wanted = engine.selectionsOf(task).map(sel => sel.tab);
  const missing = wanted.filter(t => !tabs.includes(t));
  if (missing.length) {
    excel.remove(newId);
    throw new Error('That file has no sheet named ' + missing.map(t => '"' + t + '"').join(', ') +
      '. It has: ' + tabs.join(', ') + '.');
  }

  const oldId = task.uploadId;
  task.uploadId = newId;
  task.spreadsheetName = name;
  task.pausedReason = '';
  task.consecutiveFailures = 0;
  store.upsertTask(task);

  store.addUpload({ id: newId, name, tabs, at: new Date().toISOString() });

  // Drop the previous file once nothing points at it.
  if (oldId && !store.get().tasks.some(t => t.uploadId === oldId)) {
    excel.remove(oldId);
    store.removeUpload(oldId);
  }

  store.log('ok', task.name, 'New file uploaded: ' + name);
  const result = await engine.runTask(task.id, 'new file', { seedFirst: true, concurrency: 4 });
  res.json({ upload: { id: newId, name, tabs }, result });
}));

app.post('/api/tasks/:id/clear-links', wrap(async (req, res) => {
  store.clearLinks(req.params.id);
  store.log('info', (store.getTask(req.params.id) || {}).name || '', 'Link table cleared — next sync re-sends every row.');
  res.json({ ok: true });
}));

// ══════════════════════════════════════════════════════════
// Log
// ══════════════════════════════════════════════════════════


// ══════════════════════════════════════════════════════════

store.load();
excel.init(path.join(__dirname, 'uploads'));
scheduler.start();

// A prototype that dies silently is worse than one that logs and carries on.
// Without these, one unhandled rejection kills the process and the browser
// just shows "Failed to fetch" with no clue why.
process.on('unhandledRejection', err => {
  console.error('\n[unhandled rejection]', (err && err.message) || err);
});
process.on('uncaughtException', err => {
  console.error('\n[uncaught exception]', (err && err.stack) || err);
});

app.listen(PORT, () => {
  const c = google.cfg();
  console.log('');
  console.log('  AJEMS Google Sheets Connector — prototype');
  console.log('  ────────────────────────────────────────');
  console.log('  Open:         http://localhost:' + PORT);
  console.log('  Redirect URI: ' + c.redirectUri);
  const problems = google.credentialProblems();
  console.log('  Google creds: ' + (problems.length ? 'NOT READY' : 'loaded'));
  if (problems.length) {
    problems.forEach(x => console.log('     - ' + x));
    console.log('     Google Sheets are unavailable until these are fixed.');
    console.log('     Uploading an Excel or CSV file works without them.');
  }
  console.log('');
  console.log('  Leave this window open. Closing it stops scheduled syncs.');
  console.log('');
}).on('error', err => {
  if (err.code === 'EADDRINUSE') {
    console.error(`\n  Port ${PORT} is already in use.`);
    console.error('  Another copy is probably still running — close it, or set PORT in .env.\n');
    process.exit(1);
  }
  throw err;
});
