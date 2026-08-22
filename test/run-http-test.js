/**
 * Tests the HTTP routes the browser actually calls.
 *
 * The unit tests all passed while the UI was broken, because they exercised
 * lib/ modules directly and the mock used `id` where the real API uses
 * `app_id`. Anything the UI depends on gets checked here, through the server.
 */
const assert = require('assert');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const ROOT = path.join(__dirname, '..');
try { fs.unlinkSync(path.join(ROOT, 'data.json')); } catch (e) {}

let pass = 0, fail = 0;
function check(name, fn) {
  try { fn(); console.log('  PASS  ' + name); pass++; }
  catch (e) { console.log('  FAIL  ' + name + '\n        ' + e.message); fail++; }
}

const BASE = 'http://localhost:3011';

async function get(p) {
  const r = await fetch(BASE + p);
  return { status: r.status, body: await r.json().catch(() => null) };
}
async function del(p) {
  const r = await fetch(BASE + p, { method: 'DELETE' });
  return { status: r.status, body: await r.json().catch(() => null) };
}
async function post(p, body) {
  const r = await fetch(BASE + p, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body || {})
  });
  return { status: r.status, body: await r.json().catch(() => null) };
}

(async function main() {
  const mock = spawn('node', [path.join(ROOT, 'test', 'mock-ajems.js')], { stdio: 'ignore' });
  const srv = spawn('node', [path.join(ROOT, 'server.js')], {
    stdio: 'ignore',
    env: Object.assign({}, process.env, {
      PORT: '3011',
      GOOGLE_API_KEY: 'AIzaTestKeyFromEnv',
      GOOGLE_CLIENT_ID: '914021105321-test.apps.googleusercontent.com'
    })
  });

  const stop = () => { try { srv.kill(); mock.kill(); } catch (e) {} };
  process.on('exit', stop);

  await new Promise(r => setTimeout(r, 2200));

  console.log('\n1. Status');
  let r = await get('/api/status');
  check('status responds', () => assert.equal(r.status, 200));
  check('reports not signed in', () => assert.equal(r.body.signedIn, false));
  check('scopeOk is present', () => assert.equal(typeof r.body.scopeOk, 'boolean'));

  console.log('\n2. Connect to AJEMS');
  r = await post('/api/ajems/connect', {
    baseUrl: 'http://localhost:4100', secretKey: 'test-secret-key'
  });
  check('connects', () => assert.equal(r.status, 200));
  check('returns the tenant', () => assert.equal(r.body.tenant, 'kathaa'));

  console.log('\n3. THE BUG: every app and form must carry a usable id');
  const apps = r.body.apps;
  check('one app returned', () => assert.equal(apps.length, 1));
  check('app.id is not empty', () => {
    assert.ok(apps[0].id, 'app.id was "' + apps[0].id + '" — the UI would say ' +
      '"choose an application first" with one selected');
  });
  check('app.id matches app_id from the API', () => assert.equal(apps[0].id, 'app1'));
  check('app.title survives', () => assert.equal(apps[0].title, 'Sales'));

  const form = apps[0].forms[0];
  check('form.id is not empty', () => assert.ok(form.id, 'form.id was "' + form.id + '"'));
  check('form.id matches form_id from the API', () => assert.equal(form.id, 'form1'));
  check('form.responsesUrl ends with a trailing slash', () =>
    assert.ok(form.responsesUrl.endsWith('/'), form.responsesUrl));
  check('form.detailUrl is usable', () => assert.ok(/\/forms\/form1\/$/.test(form.detailUrl)));
  check('fieldCount is right', () => assert.equal(form.fieldCount, 5));

  console.log('\n4. Re-reading the workspace gives the same shape');
  r = await get('/api/ajems/workspace');
  check('workspace responds', () => assert.equal(r.status, 200));
  check('app id still present on re-read', () => assert.ok(r.body.apps[0].id));
  check('form id still present on re-read', () => assert.ok(r.body.apps[0].forms[0].id));

  console.log('\n5. Reading a form\'s fields');
  r = await get('/api/ajems/form?detailUrl=' + encodeURIComponent(form.detailUrl));
  check('fields returned', () => assert.equal(r.body.fields.length, 5));
  check('every field has a key', () => assert.ok(r.body.fields.every(f => f.key)));
  check('every field has a label', () => assert.ok(r.body.fields.every(f => f.label)));

  console.log('\n6. Creating an app, then a form inside it');
  r = await post('/api/ajems/create-app', { title: 'Test Ajems' });
  check('app created', () => assert.ok(r.status === 200 || r.status === 201));
  const newAppId = r.body.app_id || r.body.id;
  check('created app returns an id', () => assert.ok(newAppId, JSON.stringify(r.body)));

  r = await get('/api/ajems/workspace');
  const testApp = r.body.apps.find(a => a.title === 'Test Ajems');
  check('new app appears in the workspace', () => assert.ok(testApp));
  check('new app has a usable id', () => assert.ok(testApp && testApp.id));

  console.log('\n7. Google routes refuse cleanly when not signed in');
  r = await get('/api/sheets');
  check('sheets returns an error, not an empty list', () => {
    assert.ok(r.status >= 400, 'status was ' + r.status);
    assert.ok(/sign|scope|permission/i.test(r.body.error), r.body.error);
  });
  r = await get('/api/picker-token');
  check('picker token refuses when signed out', () => assert.ok(r.status >= 400));

  console.log('\n7b. API key is developer config, not user input');
  r = await get('/api/status');
  check('api key comes from the environment', () => assert.equal(r.body.apiKey, 'AIzaTestKeyFromEnv'));
  r = await post('/api/settings', { pickerApiKey: 'anything' });
  check('there is no settings route to set it', () => assert.equal(r.status, 404));

  // Without setAppId the Picker does not grant drive.file access to the
  // picked file, and every later read fails with "not found". The project
  // number is what feeds setAppId, so the UI must always receive it.
  r = await get('/api/status');
  check('project number reaches the browser for setAppId', () => {
    assert.ok(/^\d+$/.test(r.body.projectNumber), 'got "' + r.body.projectNumber + '"');
  });
  check('only drive.file is requested', () => {
    const g = require('../lib/google');
    const url = g.authUrl();
    assert.ok(url.includes('drive.file'), 'drive.file missing');
    assert.ok(!url.includes('drive.readonly'), 'drive.readonly must not be requested');
    assert.ok(!url.includes('spreadsheets'), 'spreadsheets scope must not be requested');
  });

  console.log('\n8. Tasks round-trip');
  r = await post('/api/tasks', { name: 'HTTP test task' });
  const tid = r.body.id;
  check('task created with an id', () => assert.ok(tid));
  r = await get('/api/tasks');
  check('task is listed', () => assert.ok(r.body.tasks.some(t => t.id === tid)));
  check('scheduler reports status for it', () => assert.ok(r.body.status[tid]));

  r = await post('/api/tasks/' + tid + '/run', {});
  check('running an unconfigured task explains why', () =>
    assert.ok(/sheet|form|mapped/i.test(r.body.error || ''), JSON.stringify(r.body)));

  await fetch(BASE + '/api/tasks/' + tid, { method: 'DELETE' });

  console.log('\n8b. Excel upload route');
  const XLSX2 = require('xlsx');
  const os2 = require('os');
  const wb2 = XLSX2.utils.book_new();
  XLSX2.utils.book_append_sheet(wb2, XLSX2.utils.aoa_to_sheet([
    ['Name', 'Email'], ['A', 'a@x.com']
  ]), 'Data');
  const fx = path.join(os2.tmpdir(), 'http-fixture.xlsx');
  XLSX2.writeFile(wb2, fx);

  r = await post('/api/uploads', { name: 'http-fixture.xlsx', data: fs.readFileSync(fx).toString('base64') });
  check('upload accepted without Google auth', () => assert.equal(r.status, 200));
  const upId = r.body.id;
  check('upload reports its sheet names', () => assert.deepEqual(r.body.tabs, ['Data']));

  r = await get('/api/uploads');
  check('upload appears in the shared library', () =>
    assert.ok(r.body.uploads.some(u => u.id === upId)));

  r = await get('/api/sheets/' + encodeURIComponent(upId) + '/tabs?source=excel');
  check('tabs route reads the upload', () => assert.equal(r.body.tabs.length, 1));

  r = await get('/api/sheets/' + encodeURIComponent(upId) + '/preview?source=excel&tabs=Data');
  check('preview detects columns in the upload', () => {
    assert.equal(r.body.analysis.length, 2);
    assert.equal(r.body.analysis[1].field_type, 'Email');
  });

  r = await post('/api/uploads', { name: 'notes.txt', data: 'AAAA' });
  check('a non-spreadsheet upload is refused', () => assert.ok(r.status >= 400));

  // An upload is shared, so it must not vanish while a task still uses it.
  r = await post('/api/tasks', { name: 'uses upload', source: 'excel', uploadId: upId });
  const holder = r.body.id;
  r = await del('/api/uploads/' + encodeURIComponent(upId));
  check('an upload in use cannot be removed', () => {
    assert.ok(r.status >= 400);
    assert.ok(/task/i.test(r.body.error || ''), r.body.error);
  });
  await del('/api/tasks/' + holder);
  r = await del('/api/uploads/' + encodeURIComponent(upId));
  check('an unused upload can be removed', () => assert.equal(r.status, 200));

  console.log('\n8c. Tasks belong to an organisation');
  r = await post('/api/tasks', { name: 'org A task' });
  const orgATask = r.body.id;
  r = await get('/api/tasks');
  check('org A sees its task', () => assert.ok(r.body.tasks.some(t => t.id === orgATask)));
  check('the tenant is reported', () => assert.equal(r.body.tenant, 'kathaa'));

  await post('/api/ajems/connect', { baseUrl: 'http://localhost:4100', secretKey: 'other-secret-key' });
  r = await get('/api/tasks');
  check('another organisation cannot see it', () =>
    assert.ok(!r.body.tasks.some(t => t.id === orgATask)));
  check('the other tenant is reported', () => assert.equal(r.body.tenant, 'otherorg'));

  await post('/api/ajems/connect', { baseUrl: 'http://localhost:4100', secretKey: 'test-secret-key' });
  r = await get('/api/tasks');
  check('reconnecting brings the task back', () =>
    assert.ok(r.body.tasks.some(t => t.id === orgATask)));
  await del('/api/tasks/' + orgATask);

  console.log('\n8d. Replacing the file behind an Excel task');
  const XLSX3 = require('xlsx');
  const os3 = require('os');
  const mkbook = (rows, sheet) => {
    const wb = XLSX3.utils.book_new();
    XLSX3.utils.book_append_sheet(wb, XLSX3.utils.aoa_to_sheet(rows), sheet || 'Leads');
    const f = path.join(os3.tmpdir(), 'rf_' + Math.random().toString(36).slice(2) + '.xlsx');
    XLSX3.writeFile(wb, f);
    return fs.readFileSync(f).toString('base64');
  };
  const HEAD = ['Name', 'Email', 'Amount'];
  const fileV1 = mkbook([HEAD, ['Asha', 'asha@x.com', '1000'], ['Vik', 'vik@x.com', '2000']]);
  const fileV2 = mkbook([HEAD, ['Asha', 'asha@x.com', '1000'], ['Vik', 'vik@x.com', '2000'], ['Meera', 'meera@x.com', '3000']]);
  const fileV3 = mkbook([HEAD, ['Asha', 'asha@x.com', '9999'], ['Vik', 'vik@x.com', '2000'], ['Meera', 'meera@x.com', '3000']]);
  const fileWrong = mkbook([HEAD, ['X', 'x@x.com', '1']], 'OtherTab');

  r = await post('/api/uploads', { name: 'rf-v1.xlsx', data: fileV1 });
  const rfUpload = r.body.id;
  r = await get('/api/ajems/workspace');
  const rfApp = r.body.apps[0], rfForm = rfApp.forms[0];
  r = await get('/api/ajems/form?detailUrl=' + encodeURIComponent(rfForm.detailUrl));
  const rfMap = [
    { column: 'Name',   fieldKey: r.body.fields[0].key, fieldType: 'Text' },
    { column: 'Email',  fieldKey: r.body.fields[2].key, fieldType: 'Email' },
    { column: 'Amount', fieldKey: r.body.fields[4].key, fieldType: 'Number' }
  ];
  r = await post('/api/tasks', {
    name: 'replace-file task', source: 'excel', uploadId: rfUpload, spreadsheetName: 'rf-v1.xlsx',
    appId: rfApp.id, appTitle: rfApp.title, enabled: true,
    identity: 'key', keyColumns: ['Email'], updateExisting: true,
    selections: [{ tab: 'Leads', formId: rfForm.id, formTitle: rfForm.title,
                   postUrl: rfForm.responsesUrl, getUrl: rfForm.responsesUrl, mapping: rfMap }]
  });
  const rfTask = r.body.id;
  await new Promise(x => setTimeout(x, 1200));      // the sync fired on save

  r = await post('/api/tasks/' + rfTask + '/replace-file', { name: 'rf-v1-again.xlsx', data: fileV1 });
  check('the same file again creates nothing', () => assert.equal(r.body.result.created, 0));
  check('and recognises the rows it already sent', () => assert.equal(r.body.result.unchanged, 2));

  r = await post('/api/tasks/' + rfTask + '/replace-file', { name: 'rf-v2.xlsx', data: fileV2 });
  check('a file with one new row sends only that row', () => {
    assert.equal(r.body.result.created, 1);
    assert.equal(r.body.result.unchanged, 2);
  });

  r = await post('/api/tasks/' + rfTask + '/replace-file', { name: 'rf-v3.xlsx', data: fileV3 });
  check('a changed row updates rather than duplicating', () => {
    assert.equal(r.body.result.created, 0);
    assert.equal(r.body.result.updated, 1);
  });

  r = await post('/api/tasks/' + rfTask + '/replace-file', { name: 'rf-wrong.xlsx', data: fileWrong });
  check('a file missing the mapped tab is refused by name', () => {
    assert.ok(r.status >= 400);
    assert.ok(/no sheet named "Leads"/.test(r.body.error || ''), r.body.error);
  });

  r = await post('/api/tasks/' + rfTask + '/run', {});
  check('the task still works after the refusal', () => assert.ok(!r.body.error));
  await del('/api/tasks/' + rfTask);

  console.log('\n9. Bad input is rejected, server stays up');
  r = await post('/api/ajems/connect', { baseUrl: '', secretKey: '' });
  check('empty connect rejected', () => assert.equal(r.status, 500));
  r = await get('/api/status');
  check('server still alive afterwards', () => assert.equal(r.status, 200));

  console.log(`\n────────────────────\n  ${pass} passed, ${fail} failed\n`);
  stop();
  process.exit(fail ? 1 : 0);
})();
