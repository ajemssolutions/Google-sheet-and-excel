/**
 * End-to-end test of everything that does not need real Google OAuth.
 * Google is stubbed at the module level; AJEMS is the mock server.
 */
process.env.SKIP_ENV = '1';

const assert = require('assert');
const path = require('path');
const fs = require('fs');

// fresh data file
const DATA = path.join(__dirname, '..', 'data.json');
try { fs.unlinkSync(DATA); } catch (e) {}

const store = require('../lib/store');
const TENANT = 'kathaa';
const ajems = require('../lib/ajems');
const detect = require('../lib/detect');

// ── stub Google ───────────────────────────────────────────
const google = require('../lib/google');
let SHEET = {
  modifiedTime: '2026-08-10T10:00:00.000Z',
  headers: ['First name', 'Last name', 'Email', 'Mobile', 'Deal value'],
  rows: [
    { __rowNumber: 2, 'First name': 'Asha',   'Last name': 'Rao',   Email: 'asha@x.com',   Mobile: '9820011223', 'Deal value': '12,000' },
    { __rowNumber: 3, 'First name': 'Vikram', 'Last name': 'Shah',  Email: 'vikram@x.com', Mobile: '9820011224', 'Deal value': '8500' },
    { __rowNumber: 4, 'First name': 'Meera',  'Last name': 'Patel', Email: 'meera@x.com',  Mobile: '9820011225', 'Deal value': '31000' }
  ]
};
google.fileMeta = async () => ({ id: 'sheet1', name: 'Leads', modifiedTime: SHEET.modifiedTime });
// TABS lets a test hand back different rows per tab.
let TABS = null;
google.readTab  = async (auth, id, tab) => {
  if (TABS && TABS[tab]) return { headers: SHEET.headers, rows: TABS[tab] };
  return { headers: SHEET.headers, rows: SHEET.rows };
};

const engine = require('../lib/engine');

let pass = 0, fail = 0;
function check(name, fn) {
  try { fn(); console.log('  PASS  ' + name); pass++; }
  catch (e) { console.log('  FAIL  ' + name + '\n        ' + e.message); fail++; }
}

(async function main() {
  require('./mock-ajems');
  await new Promise(r => setTimeout(r, 300));

  console.log('\n1. Field detection');
  const analysis = detect.analyse(SHEET.headers, SHEET.rows);
  const byCol = Object.fromEntries(analysis.map(a => [a.column, a.field_type]));
  check('Email column detected as Email',      () => assert.equal(byCol['Email'], 'Email'));
  check('Mobile column detected as Phone',     () => assert.equal(byCol['Mobile'], 'Phone'));
  check('Deal value detected as Number',       () => assert.equal(byCol['Deal value'], 'Number'));
  check('First name detected as Text',         () => assert.equal(byCol['First name'], 'Text'));

  console.log('\n2. URL normalisation (the trailing-slash trap)');
  check('adds a missing trailing slash', () =>
    assert.equal(ajems.withSlash('https://h/forms/abc'), 'https://h/forms/abc/'));
  check('leaves an existing one alone', () =>
    assert.equal(ajems.withSlash('https://h/forms/abc/'), 'https://h/forms/abc/'));
  check('keeps the query string after the slash', () =>
    assert.equal(ajems.withSlash('https://h/f?a=1'), 'https://h/f/?a=1'));
  check('bare host becomes a json_builder base', () =>
    assert.equal(ajems.normalizeBase('kathaa.buildprohub-server.com'),
                 'https://kathaa.buildprohub-server.com/json_builder/'));

  console.log('\n3. AJEMS connection');
  let ws;
  try {
    ws = await ajems.workspaceConfig('http://localhost:4100', 'test-secret-key');
    check('workspace_config returns the tenant', () => assert.equal(ws.data.tenant, 'kathaa'));
    check('one app with one form',               () => assert.equal(ws.data.apps.length, 1));
  } catch (e) { check('workspace_config', () => { throw e; }); return; }

  let rejected = false;
  try { await ajems.workspaceConfig('http://localhost:4100', 'wrong-key'); }
  catch (e) { rejected = /401|rejected/i.test(e.message); }
  check('a wrong secret key is rejected', () => assert.ok(rejected));

  store.setOrgAjems(TENANT, { baseUrl: ws.base, secretKey: 'test-secret-key' });

  console.log('\n4. Task setup');
  const app = ws.data.apps[0];
  const form = app.forms[0];
  const urls = ajems.formUrls(ws.base, app, form);
  SHEET.tab = 'Sheet1';

  check('app id read from app_id', () => assert.equal(app.app_id, 'app1'));
  check('form id read from form_id', () => assert.equal(urls.formId, 'form1'));

  const task = store.newTask();
  task.name = 'Leads test';
  task.spreadsheetId = 'sheet1';
  task.spreadsheetName = 'Leads';
  task.sheetTab = 'Sheet1';
  task.appId = app.id;
  task.identity = 'key';
  task.keyColumns = ['Email'];
  task.updateExisting = true;
  const MAPPING = [
    { column: 'First name', fieldKey: 'text_1000000000001',   fieldType: 'Text' },
    { column: 'Last name',  fieldKey: 'text_1000000000002',   fieldType: 'Text' },
    { column: 'Email',      fieldKey: 'email_1000000000003',  fieldType: 'Email' },
    { column: 'Mobile',     fieldKey: 'phone_1000000000004',  fieldType: 'Phone' },
    { column: 'Deal value', fieldKey: 'number_1000000000005', fieldType: 'Number' }
  ];
  task.selections = [{
    tab: 'Sheet1', formId: form.form_id, formTitle: form.title,
    postUrl: urls.responses, getUrl: urls.responses, mapping: MAPPING
  }];
  store.upsertTask(TENANT, task);

  const payload = engine.buildPayload(task, SHEET.rows[0], task.selections[0]);
  check('payload is keyed by AJEMS field key', () =>
    assert.equal(payload['email_1000000000003'], 'asha@x.com'));
  check('"12,000" is sent as the number 12000', () =>
    assert.strictEqual(payload['number_1000000000005'], 12000));

  console.log('\n5. First sync — everything is new');
  let r = await engine.runTask(TENANT, task.id, 'test');
  check('3 records created', () => assert.equal(r.created, 3));
  check('nothing updated',   () => assert.equal(r.updated, 0));
  check('no failures',       () => assert.equal(r.failed, 0));

  console.log('\n6. Re-run with no changes — nothing should be sent twice');
  r = await engine.runTask(TENANT, task.id, 'test');
  check('0 created',      () => assert.equal(r.created, 0));
  check('0 updated',      () => assert.equal(r.updated, 0));
  check('3 unchanged',    () => assert.equal(r.unchanged, 3));

  const mock = require('./mock-ajems');
  check('AJEMS still holds exactly 3 records', () =>
    assert.equal(mock.state.responses[form.form_id].length, 3));

  console.log('\n7. Edit a row in the sheet — it should UPDATE, not duplicate');
  SHEET.rows[1]['Deal value'] = '99999';
  SHEET.modifiedTime = '2026-08-10T11:00:00.000Z';
  r = await engine.runTask(TENANT, task.id, 'test');
  check('1 updated',   () => assert.equal(r.updated, 1));
  check('0 created',   () => assert.equal(r.created, 0));
  check('2 unchanged', () => assert.equal(r.unchanged, 2));
  check('still 3 records in AJEMS', () =>
    assert.equal(mock.state.responses[form.form_id].length, 3));
  check('the stored value really changed', () => {
    const rec = mock.state.responses[form.form_id].find(x => x.response['email_1000000000003'] === 'vikram@x.com');
    assert.strictEqual(rec.response['number_1000000000005'], 99999);
  });

  console.log('\n8. Append a new row — only that row goes up');
  SHEET.rows.push({ __rowNumber: 5, 'First name': 'Sunil', 'Last name': 'Kadam',
                    Email: 'sunil@x.com', Mobile: '9820011226', 'Deal value': '4200' });
  r = await engine.runTask(TENANT, task.id, 'test');
  check('1 created',   () => assert.equal(r.created, 1));
  check('3 unchanged', () => assert.equal(r.unchanged, 3));
  check('4 records in AJEMS', () =>
    assert.equal(mock.state.responses[form.form_id].length, 4));

  console.log('\n9. Rows reordered — must NOT be seen as new');
  SHEET.rows.reverse();
  SHEET.rows.forEach((row, i) => row.__rowNumber = i + 2);   // row numbers all shift
  r = await engine.runTask(TENANT, task.id, 'test');
  check('0 created despite every row number changing', () => assert.equal(r.created, 0));
  check('4 unchanged', () => assert.equal(r.unchanged, 4));

  console.log('\n10. Whole-row identity mode');
  const t2 = store.newTask();
  Object.assign(t2, task, { id: 't2', identity: 'hash', keyColumns: [],
    selections: [Object.assign({}, task.selections[0])] });
  store.upsertTask(TENANT, t2);
  await engine.runTask(TENANT, 't2', 'test');
  const before = mock.state.responses[form.form_id].length;
  await engine.runTask(TENANT, 't2', 'test');
  check('hash mode also skips unchanged rows on re-run', () =>
    assert.equal(mock.state.responses[form.form_id].length, before));

  console.log('\n11. Seeding from AJEMS (fresh machine, empty link table)');
  store.clearLinks(TENANT, task.id);
  const seeded = await engine.seedLinks(TENANT, store.getTask(TENANT, task.id), store.getTask(TENANT, task.id).selections[0]);
  check('links rebuilt from existing AJEMS records', () => assert.ok(seeded.seeded >= 4));
  r = await engine.runTask(TENANT, task.id, 'test');
  check('after seeding, nothing is re-sent', () => assert.equal(r.created, 0));

  console.log('\n12. Multiple tabs feeding one form');
  const before12 = mock.state.responses[form.form_id].length;
  TABS = {
    'North': [{ __rowNumber: 2, 'First name': 'Nita',  'Last name': 'Joshi', Email: 'nita@x.com',  Mobile: '9820000001', 'Deal value': '100' }],
    'South': [{ __rowNumber: 2, 'First name': 'Ravi',  'Last name': 'Kumar', Email: 'ravi@x.com',  Mobile: '9820000002', 'Deal value': '200' }]
  };
  const multi = store.newTask();
  Object.assign(multi, task, {
    id: 'multi', identity: 'hash', keyColumns: [],
    selections: [
      { tab: 'North', formId: form.form_id, formTitle: 'Leads', postUrl: urls.responses, getUrl: urls.responses, mapping: MAPPING },
      { tab: 'South', formId: form.form_id, formTitle: 'Leads', postUrl: urls.responses, getUrl: urls.responses, mapping: MAPPING }
    ]
  });
  store.upsertTask(TENANT, multi);
  let rm = await engine.runTask(TENANT, 'multi', 'test');
  check('rows from both tabs are read', () => assert.equal(rm.rows, 2));
  check('both rows created', () => assert.equal(rm.created, 2));
  rm = await engine.runTask(TENANT, 'multi', 'test');
  check('re-run sends nothing again', () => assert.equal(rm.created, 0));

  // Identical values in two different tabs are two different records.
  TABS = {
    'North': [{ __rowNumber: 2, 'First name': 'Same', 'Last name': 'Row', Email: 's@x.com', Mobile: '9', 'Deal value': '1' }],
    'South': [{ __rowNumber: 2, 'First name': 'Same', 'Last name': 'Row', Email: 's@x.com', Mobile: '9', 'Deal value': '1' }]
  };
  const dup = store.newTask();
  Object.assign(dup, task, {
    id: 'dup', identity: 'hash', keyColumns: [],
    selections: [
      { tab: 'North', formId: form.form_id, formTitle: 'Leads', postUrl: urls.responses, getUrl: urls.responses, mapping: MAPPING },
      { tab: 'South', formId: form.form_id, formTitle: 'Leads', postUrl: urls.responses, getUrl: urls.responses, mapping: MAPPING }
    ]
  });
  store.upsertTask(TENANT, dup);
  rm = await engine.runTask(TENANT, 'dup', 'test');
  check('same row in two tabs counts as two records', () => assert.equal(rm.created, 2));

  const err = engine.rowKey(dup, { 'First name': 'Same', __tab: 'North' }, dup.selections[0]) !==
              engine.rowKey(dup, { 'First name': 'Same', __tab: 'South' }, dup.selections[1]);
  check('row key includes the tab under hash identity', () => assert.ok(err));

  // Each tab can target a DIFFERENT form — the point of the per-tab model.
  console.log('');
  const twoForms = await ajems.post(TENANT, ws.base + 'forms/', {
    app: app.app_id, title: 'South leads', fields: form.fields
  });
  const ws2 = await ajems.workspaceConfig('http://localhost:4100', 'test-secret-key');
  const app2 = ws2.data.apps[0];
  const southForm = app2.forms.find(f => f.title === 'South leads');
  const southUrls = ajems.formUrls(ws2.base, app2, southForm);

  TABS = {
    'North': [{ __rowNumber: 2, 'First name': 'N1', 'Last name': 'x', Email: 'n1@x.com', Mobile: '1', 'Deal value': '1' }],
    'South': [{ __rowNumber: 2, 'First name': 'S1', 'Last name': 'y', Email: 's1@x.com', Mobile: '2', 'Deal value': '2' }]
  };
  const split = store.newTask();
  Object.assign(split, task, {
    id: 'split', identity: 'hash', keyColumns: [],
    selections: [
      { tab: 'North', formId: form.form_id, formTitle: 'Leads', postUrl: urls.responses, getUrl: urls.responses, mapping: MAPPING },
      { tab: 'South', formId: southForm.form_id, formTitle: 'South leads', postUrl: southUrls.responses, getUrl: southUrls.responses, mapping: MAPPING }
    ]
  });
  store.upsertTask(TENANT, split);
  const beforeSouth = (mock.state.responses[southForm.form_id] || []).length;
  rm = await engine.runTask(TENANT, 'split', 'test');
  check('two tabs, two different forms — both created', () => assert.equal(rm.created, 2));
  check('the second form received exactly its own row', () =>
    assert.equal((mock.state.responses[southForm.form_id] || []).length, beforeSouth + 1));
  check('per-tab breakdown is reported', () => {
    assert.equal(rm.perTab.length, 2);
    assert.equal(rm.perTab[1].form, 'South leads');
  });

  TABS = null;

  console.log('\n13. Error handling');
  const bad = store.newTask();
  Object.assign(bad, task, { id: 'bad', selections: [Object.assign({}, task.selections[0],
    { postUrl: 'http://localhost:4100/json_builder/forms/nope/responses/' })] });
  store.upsertTask(TENANT, bad);
  r = await engine.runTask(TENANT, 'bad', 'test');
  check('failures are counted, not thrown', () => assert.ok(r.failed > 0));
  check('the error names the tab and row', () => assert.ok(/row \d+/.test(r.errors[0]) &&
        /Sheet1|Sheet row/.test(r.errors[0]), r.errors[0]));

  console.log('\n14. Excel upload as a source');
  const XLSX = require('xlsx');
  const excel = require('../lib/excel');
  const os = require('os');
  const tmp = path.join(os.tmpdir(), 'ajems-uploads-test');
  excel.init(tmp);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['First name', 'Last name', 'Email', 'Mobile', 'Deal value'],
    ['Asha', 'Rao', 'asha@x.com', '9820011223', '12,000'],
    ['', '', '', '', ''],
    ['Vikram', 'Shah', 'vikram@x.com', '9820011224', '8500']
  ]), 'North');
  const xlsxPath = path.join(tmp, 'fixture.xlsx');
  fs.mkdirSync(tmp, { recursive: true });
  XLSX.writeFile(wb, xlsxPath);

  const uploadId = excel.save('fixture.xlsx', fs.readFileSync(xlsxPath).toString('base64'));
  check('upload is stored', () => assert.ok(excel.exists(uploadId)));
  check('a non-spreadsheet is rejected', () => {
    assert.throws(() => excel.save('notes.txt', 'AAA'));
  });

  const info = excel.listTabs(uploadId);
  check('tabs read from the workbook', () => assert.equal(info.tabs[0].title, 'North'));

  const read = excel.readTab(uploadId, 'North', 1);
  check('headers match the sheet reader shape', () =>
    assert.deepEqual(read.headers, ['First name', 'Last name', 'Email', 'Mobile', 'Deal value']));
  check('blank rows are skipped', () => assert.equal(read.rows.length, 2));

  const xlTask = store.newTask();
  Object.assign(xlTask, {
    id: 'xl', name: 'Excel import', source: 'excel', uploadId,
    enabled: true, identity: 'hash', updateExisting: true,
    selections: [{
      tab: 'North', formId: form.form_id, formTitle: 'Leads',
      postUrl: urls.responses, getUrl: urls.responses, mapping: MAPPING
    }]
  });
  store.upsertTask(TENANT, xlTask);

  const xr = await engine.runTask(TENANT, 'xl', 'test');
  check('excel rows reach AJEMS', () => assert.equal(xr.created, 2));
  check('no Google call was needed', () => assert.equal(xr.failed, 0));

  const xr2 = await engine.runTask(TENANT, 'xl', 'test');
  check('re-running an excel task sends nothing again', () => assert.equal(xr2.created, 0));

  check('"12,000" is sent as the number 12000', () => {
    const p = engine.buildPayload(xlTask, read.rows[0], xlTask.selections[0]);
    assert.strictEqual(p['number_1000000000005'], 12000);
  });

  excel.remove(uploadId);
  const gone = await engine.runTask(TENANT, 'xl', 'test');
  check('a missing file is reported, not thrown', () =>
    assert.ok(/no longer on disk/i.test(gone.error || '')));

  console.log(`\n────────────────────\n  ${pass} passed, ${fail} failed\n`);
  process.exit(fail ? 1 : 0);
})();
