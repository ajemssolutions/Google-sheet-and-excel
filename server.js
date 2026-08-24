require('./lib/env');

const path = require('path');
const crypto = require('crypto');
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

// ══════════════════════════════════════════════════════════
// Sessions
//
// Everything below belongs to ONE workspace, identified by the session
// cookie. A visitor with no session sees an empty app: no workspace, no
// Google account, no tasks. Signing in means proving you hold that
// workspace's secret key, which is what /api/ajems/connect does.
// ══════════════════════════════════════════════════════════

const COOKIE = 'ajems_sid';

function readCookie(req, name) {
  const raw = req.headers.cookie || '';
  for (const part of raw.split(';')) {
    const [k, ...v] = part.trim().split('=');
    if (k === name) return decodeURIComponent(v.join('='));
  }
  return '';
}

function setSessionCookie(res, id) {
  const secure = process.env.NODE_ENV === 'production' ||
                 /^https:/i.test(process.env.GOOGLE_REDIRECT_URI || '');
  res.setHeader('Set-Cookie',
    `${COOKIE}=${encodeURIComponent(id)}; Path=/; HttpOnly; SameSite=Lax; ` +
    `Max-Age=${store.SESSION_DAYS * 86400}` + (secure ? '; Secure' : ''));
}

function clearSessionCookie(res) {
  res.setHeader('Set-Cookie', `${COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`);
}

/** The workspace this request belongs to, or '' when signed out. */
function tenantOf(req) {
  return store.sessionTenant(readCookie(req, COOKIE));
}

/** Wraps a route and rejects it when there is no signed-in workspace. */
const needOrg = fn => (req, res) => {
  const tenant = tenantOf(req);
  if (!tenant) {
    return res.status(401).json({ error: 'Connect your AJEMS workspace first.' });
  }
  Promise.resolve(fn(req, res, tenant)).catch(err => {
    console.error(err);
    res.status(err.status && err.status < 600 ? err.status : 500)
       .json({ error: err.message || String(err) });
  });
};

/** Wraps a route that does not need a workspace. */
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
  const tenant = tenantOf(req);
  const c = google.cfg();
  const credProblems = google.credentialProblems();
  const org = tenant ? store.getOrg(tenant) : null;
  const auth = store.sessionGoogle(readCookie(req, COOKIE));
  const scope = google.scopeStatus(auth);
  const g = auth.get();

  res.json({
    // credProblems is deliberately NOT sent: it names environment variables,
    // which is developer detail. The list prints to the terminal at startup.
    signedIn: !!(g && g.access_token),
    email: g ? g.email : '',
    scopeOk: scope.ok,
    grantedScopes: scope.granted,
    googleConfigured: credProblems.length === 0,
    apiKey: c.apiKey,
    projectNumber: c.projectNumber,
    ajems: {
      baseUrl: org ? org.ajems.baseUrl : '',
      tenant: tenant,
      verified: !!(org && org.ajems.verified),
      hasKey: !!(org && org.ajems.secretKey)
    },
    taskCount: tenant ? store.tasks(tenant).length : 0,
    uploadCount: tenant ? store.uploads(tenant).length : 0
  });
}));

// ══════════════════════════════════════════════════════════
// Connecting a workspace - this is the sign-in
// ══════════════════════════════════════════════════════════

/**
 * Normalise workspace_config into what the UI needs.
 *
 * The API names these `app_id` and `form_id` - NOT `id`. Reading the wrong
 * key yields an option with an empty value, and the UI then says "choose an
 * application first" while an application is visibly selected.
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

app.post('/api/ajems/connect', wrap(async (req, res) => {
  const baseUrl = String(req.body.baseUrl || '').trim();
  const secretKey = String(req.body.secretKey || '').trim();
  if (!baseUrl || !secretKey) throw new Error('Both the workspace name and secret key are required.');

  // AJEMS itself decides whether the key is good. Only a key that works gets
  // a session, so the key is the credential.
  const { base, data } = await ajems.workspaceConfig(baseUrl, secretKey);
  const tenant = data.tenant;

  store.setOrgAjems(tenant, { baseUrl: base, secretKey });
  setSessionCookie(res, store.newSession(tenant));
  store.log('ok', '', `Signed in to "${tenant}" - ${data.apps.length} application(s).`);

  res.json({ tenant, base, apps: shapeApps(base, data.apps) });
}));

app.post('/api/signout', wrap(async (req, res) => {
  store.endSession(readCookie(req, COOKIE));
  clearSessionCookie(res);
  res.json({ ok: true });
}));

// ══════════════════════════════════════════════════════════
// Google OAuth
// ══════════════════════════════════════════════════════════

// Short-lived map of in-flight sign-ins, so the state cannot be forged.
const pendingAuth = new Map();
setInterval(() => {
  const cutoff = Date.now() - 10 * 60 * 1000;
  for (const [k, v] of pendingAuth) if (v.at < cutoff) pendingAuth.delete(k);
}, 60 * 1000).unref();

function page(title, body) {
  return `<!doctype html><meta charset="utf-8">
<style>body{font-family:Inter,system-ui,sans-serif;max-width:640px;margin:64px auto;padding:0 24px;color:#101828;line-height:1.6}
a{color:#2563EB}h2{font-size:20px;margin-bottom:6px}</style>
<h2>${title}</h2><p>${body}</p><p><a href="/">Back to the connector</a></p>`;
}

app.get('/auth/google', (req, res) => {
  const tenant = tenantOf(req);
  if (!tenant) return res.redirect('/?needorg=1');

  if (google.credentialProblems().length) {
    return res.status(500).send(page('Google sign-in is not configured',
      'The server is missing its Google credentials. Whoever runs it can see ' +
      'the details in the terminal output at startup.'));
  }

  // The state carries the workspace through Google and back, and is checked
  // on return so a callback cannot land on a different workspace.
  const nonce = crypto.randomBytes(12).toString('hex');
  pendingAuth.set(nonce, { sid: readCookie(req, COOKIE), tenant, at: Date.now() });
  res.redirect(google.authUrl(nonce));
});

app.get('/oauth/callback', wrap(async (req, res) => {
  if (req.query.error) {
    return res.send(page('Google sign-in was cancelled', String(req.query.error)));
  }
  if (!req.query.code) return res.status(400).send(page('No authorization code returned.', ''));

  const state = String(req.query.state || '');
  const pending = pendingAuth.get(state);
  pendingAuth.delete(state);
  if (!pending) {
    return res.status(400).send(page('That sign-in link has expired',
      'Go back to the connector and press Continue with Google again.'));
  }

  // Tokens land on the browser session that started the sign-in, not on the
  // workspace, so one person connecting Google does not connect it for
  // everyone who shares the secret key.
  const auth = store.sessionGoogle(pending.sid);
  const session = await google.exchangeCode(auth, req.query.code);
  store.log('ok', '', 'Google connected on one session of "' + pending.tenant +
    '" as ' + (session.email || 'unknown'));
  res.redirect('/?signedin=1');
}));

/**
 * The Google Picker runs in the browser and needs a live OAuth token.
 * Scoped to the caller's workspace, so one cannot borrow another's.
 */
app.get('/api/picker-token', needOrg(async (req, res, tenant) => {
  res.json({ token: await google.token(store.sessionGoogle(readCookie(req, COOKIE))) });
}));

app.post('/api/google/signout', needOrg(async (req, res, tenant) => {
  google.signOut(store.sessionGoogle(readCookie(req, COOKIE)));
  res.json({ ok: true });
}));

// ══════════════════════════════════════════════════════════
// Sheets and uploads
// ══════════════════════════════════════════════════════════

app.get('/api/sheets', needOrg(async (req, res, tenant) => {
  res.json({ files: await google.listSpreadsheets(store.sessionGoogle(readCookie(req, COOKIE))) });
}));

app.get('/api/sheets/:id/tabs', needOrg(async (req, res, tenant) => {
  if (req.query.source === 'excel') return res.json(excel.listTabs(req.params.id));
  res.json(await google.listTabs(store.sessionGoogle(readCookie(req, COOKIE)), req.params.id));
}));

/**
 * Reads one or more tabs and reports the columns found across them. The
 * header is always row 1 - every sheet we have seen works that way, and
 * asking the user was one more decision for no benefit.
 */
async function readTabs(auth, source, id, tabs) {
  const headers = [];
  const rows = [];
  for (const tab of tabs) {
    const r = source === 'excel'
      ? excel.readTab(id, tab, 1)
      : await google.readTab(auth, id, tab, 1);
    r.headers.forEach(h => { if (!headers.includes(h)) headers.push(h); });
    r.rows.forEach(row => { row.__tab = tab; rows.push(row); });
  }
  return { headers, rows };
}

function tabsFromQuery(q) {
  const raw = q.tabs || q.tab || '';
  return String(raw).split('\u001f').map(t => t.trim()).filter(Boolean);
}

app.get('/api/sheets/:id/preview', needOrg(async (req, res, tenant) => {
  const tabs = tabsFromQuery(req.query);
  if (!tabs.length) throw new Error('No tab selected.');
  const { headers, rows } = await readTabs(store.sessionGoogle(readCookie(req, COOKIE)),
                                           req.query.source, req.params.id, tabs);
  res.json({
    tabs,
    headers,
    rowCount: rows.length,
    analysis: detect.analyse(headers, rows),
    sample: rows.slice(0, 5)
  });
}));

/**
 * The file arrives as base64 in the JSON body rather than multipart, which
 * keeps the dependency list short. Uploads belong to the workspace: several
 * of its tasks can share one file, and no other workspace can see it.
 */
app.post('/api/uploads', needOrg(async (req, res, tenant) => {
  const { name, data } = req.body || {};
  if (!name || !data) throw new Error('Send both a file name and its contents.');

  const id = excel.save(name, data);
  const info = excel.listTabs(id);
  const upload = { id, name, tabs: info.tabs.map(t => t.title), at: new Date().toISOString() };
  store.addUpload(tenant, upload);
  store.log('ok', '', 'Uploaded ' + name + ' (' + upload.tabs.length + ' sheet(s)) for "' + tenant + '".');
  res.json(upload);
}));

app.get('/api/uploads', needOrg(async (req, res, tenant) => {
  res.json({ uploads: store.uploads(tenant) });
}));

app.delete('/api/uploads/:id', needOrg(async (req, res, tenant) => {
  if (!store.uploads(tenant).some(u => u.id === req.params.id)) {
    throw new Error('That file does not belong to this workspace.');
  }
  const inUse = store.tasks(tenant).filter(t => t.uploadId === req.params.id);
  if (inUse.length) {
    throw new Error('That file is used by ' + inUse.length + ' task(s): ' +
      inUse.map(t => t.name).join(', ') + '. Delete them first.');
  }
  excel.remove(req.params.id);
  store.removeUpload(tenant, req.params.id);
  res.json({ ok: true });
}));

// ══════════════════════════════════════════════════════════
// AJEMS applications and forms
// ══════════════════════════════════════════════════════════

app.get('/api/ajems/workspace', needOrg(async (req, res, tenant) => {
  const org = store.getOrg(tenant);
  const { base, data } = await ajems.workspaceConfig(org.ajems.baseUrl, org.ajems.secretKey);
  res.json({ tenant: data.tenant, apps: shapeApps(base, data.apps) });
}));

app.get('/api/ajems/form', needOrg(async (req, res, tenant) => {
  const detailUrl = req.query.detailUrl;
  if (!detailUrl) throw new Error('detailUrl is required.');
  const form = await ajems.getForm(tenant, detailUrl);
  const fields = form.fields || form.form_fields || [];
  res.json({
    title: form.title || '',
    fields: fields.map(f => ({ key: f.key, label: f.label, field_type: f.field_type })),
    labels: form.labels || null
  });
}));

app.post('/api/ajems/create-app', needOrg(async (req, res, tenant) => {
  const org = store.getOrg(tenant);
  const title = String(req.body.title || '').trim();
  if (!title) throw new Error('The app needs a name.');
  const created = await ajems.createApp(tenant, org.ajems.baseUrl, {
    title,
    description: req.body.description || '',
    icon: req.body.icon || 'pi pi-align-justify',
    color: req.body.color || 'cyan'
  });
  store.log('ok', '', 'Created AJEMS app: ' + title);
  res.json(created);
}));

app.post('/api/ajems/create-form', needOrg(async (req, res, tenant) => {
  const org = store.getOrg(tenant);
  const { spreadsheetId, uploadId, source, appId, title, columns } = req.body;
  const tabs = Array.isArray(req.body.tabs) ? req.body.tabs
             : (req.body.tab ? [req.body.tab] : []);
  if (!tabs.length) throw new Error('No tab selected.');

  const { headers, rows } = await readTabs(store.sessionGoogle(readCookie(req, COOKIE)),
    source, source === 'excel' ? uploadId : spreadsheetId, tabs);

  const wanted = Array.isArray(columns) && columns.length
    ? headers.filter(h => columns.includes(h))
    : headers;

  const fields = detect.buildFields(wanted, rows);
  const created = await ajems.createForm(tenant, org.ajems.baseUrl, appId, fields, { title });

  store.log('ok', '', `Created AJEMS form "${title}" with ${fields.length} field(s).`);
  res.json({ form: created, fields });
}));

// ══════════════════════════════════════════════════════════
// Tasks
// ══════════════════════════════════════════════════════════

/**
 * A Google task keeps its own copy of the credential that created it.
 *
 * Google is otherwise per browser session, which is what stops one person's
 * account showing up for everybody who shares the secret key. But a
 * scheduled run happens with no browser attached, so without this a task
 * would stop syncing the moment its author signed out.
 */
function captureGoogle(req, task) {
  if (task.source === 'excel') { task.google = null; return; }
  const g = store.sessionGoogle(readCookie(req, COOKIE)).get();
  if (g && g.refresh_token) {
    task.google = g;
  } else if (g) {
    // No refresh token means Google will not renew it, so scheduled runs
    // would fail later. Better to know now.
    task.google = g;
    store.log('warn', task.name,
      'Google did not return a refresh token, so scheduled syncs may stop. ' +
      'Sign in to Google again to fix it.');
  }
}

function readyToSync(t) {
  if (!t.enabled) return false;
  const hasSource = t.source === 'excel' ? !!t.uploadId : !!t.spreadsheetId;
  const hasTarget = (t.selections || []).some(s => s.postUrl);
  return hasSource && hasTarget;
}

/**
 * Run a task just after the response goes out, so saving syncs immediately
 * instead of waiting for the schedule. Deliberately not awaited: a large
 * sheet must not hold up the Save button.
 */
function syncSoon(tenant, taskId, trigger) {
  setTimeout(() => {
    engine.runTask(tenant, taskId, trigger, { seedFirst: true, concurrency: 4 })
      .catch(e => store.log('error', '', 'Sync after save failed: ' + e.message));
  }, 200);
}

app.get('/api/tasks', needOrg(async (req, res, tenant) => {
  res.json({ tasks: store.tasks(tenant), tenant, status: scheduler.status(tenant) });
}));

app.post('/api/tasks', needOrg(async (req, res, tenant) => {
  const t = store.newTask();
  Object.assign(t, req.body || {}, { id: t.id });
  captureGoogle(req, t);
  store.upsertTask(tenant, t);
  res.json(t);
  if (readyToSync(t)) syncSoon(tenant, t.id, 'first run');
}));

app.put('/api/tasks/:id', needOrg(async (req, res, tenant) => {
  const t = store.getTask(tenant, req.params.id);
  if (!t) throw new Error('Task not found.');
  const before = JSON.stringify(t.selections);
  Object.assign(t, req.body || {}, { id: t.id });
  // A changed mapping invalidates the content hashes already stored.
  if (JSON.stringify(t.selections) !== before) store.clearLinks(tenant, t.id);
  if (t.enabled) { t.pausedReason = ''; t.consecutiveFailures = 0; }
  captureGoogle(req, t);
  store.upsertTask(tenant, t);
  res.json(t);
  if (readyToSync(t)) syncSoon(tenant, t.id, 'saved');
}));

app.delete('/api/tasks/:id', needOrg(async (req, res, tenant) => {
  const t = store.getTask(tenant, req.params.id);
  if (t && t.source === 'excel' && t.uploadId) {
    const stillUsed = store.tasks(tenant).some(x => x.id !== t.id && x.uploadId === t.uploadId);
    if (!stillUsed) { excel.remove(t.uploadId); store.removeUpload(tenant, t.uploadId); }
  }
  store.removeTask(tenant, req.params.id);
  res.json({ ok: true });
}));

app.post('/api/tasks/:id/run', needOrg(async (req, res, tenant) => {
  // Pressing Sync now refreshes the task's stored credential from whoever
  // pressed it, so a task keeps working after its original author leaves.
  const t = store.getTask(tenant, req.params.id);
  if (t && t.source !== 'excel') {
    const g = store.sessionGoogle(readCookie(req, COOKIE)).get();
    if (g) { t.google = g; store.upsertTask(tenant, t); }
  }
  const result = await engine.runTask(tenant, req.params.id, req.body.trigger || 'manual', {
    seedFirst: !!req.body.seedFirst,
    concurrency: parseInt(req.body.concurrency, 10) || 4
  });
  res.json(result);
}));

/** Dry run: what would be sent for the first few rows of each tab. */
app.post('/api/tasks/:id/preview-payload', needOrg(async (req, res, tenant) => {
  const task = store.getTask(tenant, req.params.id);
  if (!task) throw new Error('Task not found.');
  const links = store.getLinks(tenant, task.id);
  const out = [];

  for (const sel of engine.selectionsOf(task)) {
    if (!sel.tab) continue;
    const { rows } = await engine.readSourceTab(tenant, task, sel.tab);
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
 * are recognised in the new file and only new or changed ones go to AJEMS.
 */
app.post('/api/tasks/:id/replace-file', needOrg(async (req, res, tenant) => {
  const task = store.getTask(tenant, req.params.id);
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
  store.upsertTask(tenant, task);
  store.addUpload(tenant, { id: newId, name, tabs, at: new Date().toISOString() });

  // Drop the previous file once nothing points at it.
  if (oldId && !store.tasks(tenant).some(t => t.uploadId === oldId)) {
    excel.remove(oldId);
    store.removeUpload(tenant, oldId);
  }

  store.log('ok', task.name, 'New file uploaded: ' + name);
  const result = await engine.runTask(tenant, task.id, 'new file', { seedFirst: true, concurrency: 4 });
  res.json({ upload: { id: newId, name, tabs }, result });
}));

app.post('/api/tasks/:id/clear-links', needOrg(async (req, res, tenant) => {
  store.clearLinks(tenant, req.params.id);
  res.json({ ok: true });
}));

// ══════════════════════════════════════════════════════════

store.init(__dirname);
excel.init(path.join(__dirname, 'uploads'));
scheduler.start();

// A prototype that dies silently is worse than one that logs and carries on.
process.on('unhandledRejection', err => {
  console.error('\n[unhandled rejection]', (err && err.message) || err);
});
process.on('uncaughtException', err => {
  console.error('\n[uncaught exception]', (err && err.stack) || err);
});

app.listen(PORT, () => {
  const c = google.cfg();
  const orgs = store.orgTenants();
  console.log('');
  console.log('  AJEMS Sheets Connector');
  console.log('  ----------------------');
  console.log('  Open:         http://localhost:' + PORT);
  console.log('  Redirect URI: ' + c.redirectUri);

  const problems = google.credentialProblems();
  console.log('  Google creds: ' + (problems.length ? 'NOT READY' : 'loaded'));
  if (problems.length) {
    problems.forEach(x => console.log('     - ' + x));
    console.log('     Google Sheets are unavailable until these are fixed.');
    console.log('     Uploading an Excel or CSV file works without them.');
  }
  console.log('  Workspaces:   ' + (orgs.length ? orgs.join(', ') : 'none yet'));
  console.log('');
  console.log('  Leave this window open. Closing it stops scheduled syncs.');
  console.log('');
}).on('error', err => {
  if (err.code === 'EADDRINUSE') {
    console.error(`\n  Port ${PORT} is already in use.`);
    console.error('  Another copy is probably still running, or set PORT in .env.\n');
    process.exit(1);
  }
  throw err;
});
