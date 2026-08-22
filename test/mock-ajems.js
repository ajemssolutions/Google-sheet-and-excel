/**
 * Stand-in for the AJEMS JSON Builder API, matching the documented shapes.
 * Reproduces the real nginx behaviour: a path WITHOUT a trailing slash
 * answers 301, so the client's slash handling is genuinely exercised.
 */
const http = require('http');

const SECRET = 'test-secret-key';
const SECRET2 = 'other-secret-key';      // a second organisation
const TENANT2 = 'otherorg';
const state = {
  tenant: 'kathaa',
  apps: [
    // NOTE: the real API uses app_id / form_id, not id. The mock used `id`
    // once and hid a real bug — keep these names matching the API exactly.
    { app_id: 'app1', title: 'Sales', forms: [
      { form_id: 'form1', title: 'Leads', fields: [
        { key: 'text_1000000000001',  label: 'First name', field_type: 'Text' },
        { key: 'text_1000000000002',  label: 'Last name',  field_type: 'Text' },
        { key: 'email_1000000000003', label: 'Email',      field_type: 'Email' },
        { key: 'phone_1000000000004', label: 'Mobile',     field_type: 'Phone' },
        { key: 'number_1000000000005',label: 'Deal value', field_type: 'Number' }
      ]}
    ]}
  ],
  responses: {},   // formId -> [ {_id, response} ]
  nextId: 1
};

function send(res, code, body) {
  const text = JSON.stringify(body);
  res.writeHead(code, { 'Content-Type': 'application/json' });
  res.end(text);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const p = url.pathname;

  // nginx: no trailing slash -> 301
  if (!p.endsWith('/')) {
    res.writeHead(301, { Location: p + '/' });
    return res.end();
  }

  const key = req.headers['x-json-builder-secret-key'];
  if (key !== SECRET && key !== SECRET2) {
    return send(res, 401, { detail: 'Invalid secret key' });
  }
  const tenant = key === SECRET2 ? TENANT2 : state.tenant;

  let raw = '';
  req.on('data', c => raw += c);
  req.on('end', () => {
    const body = raw ? JSON.parse(raw) : null;

    // workspace_config
    if (p === '/json_builder/workspace_config/') {
      return send(res, 200, {
        tenant: tenant,
        apps: state.apps.map(a => ({
          app_id: a.app_id, title: a.title,
          forms: a.forms.map(f => ({
            form_id: f.form_id, title: f.title, fields: f.fields,
            jsonBuilderUrls: {
              form: `http://localhost:4100/json_builder/forms/${f.form_id}/`,
              responses: `http://localhost:4100/json_builder/forms/${f.form_id}/responses/`
            }
          }))
        }))
      });
    }

    // create form
    if (p === '/json_builder/forms/' && req.method === 'POST') {
      const app = state.apps.find(a => a.app_id === (body.app || body.app_id));
      if (!app) return send(res, 400, { detail: 'Unknown app' });
      const form = {
        form_id: 'form' + (++state.nextId),
        title: body.title,
        fields: body.fields || []
      };
      app.forms.push(form);
      return send(res, 201, form);
    }

    // create app
    if (p === '/json_builder/apps/' && req.method === 'POST') {
      const app = { app_id: 'app' + (++state.nextId), title: body.title, forms: [] };
      state.apps.push(app);
      return send(res, 201, app);
    }

    // form detail
    let m = p.match(/^\/json_builder\/forms\/([^/]+)\/$/);
    if (m && req.method === 'GET') {
      for (const a of state.apps) {
        const f = a.forms.find(x => x.form_id === m[1]);
        if (f) {
          const labels = {};
          f.fields.forEach(x => labels[x.key] = x.label);
          return send(res, 200, { form_id: f.form_id, title: f.title, fields: f.fields, labels });
        }
      }
      return send(res, 404, { detail: 'Form not found' });
    }

    // responses list / create
    m = p.match(/^\/json_builder\/forms\/([^/]+)\/responses\/$/);
    if (m) {
      const fid = m[1];
      const known = state.apps.some(a => a.forms.some(f => f.form_id === fid));
      if (!known) {
        return send(res, 404, { detail: 'Form not found or third-party POST access not enabled' });
      }
      if (!state.responses[fid]) state.responses[fid] = [];

      if (req.method === 'GET') {
        return send(res, 200, {
          count: state.responses[fid].length,
          responses: state.responses[fid]
        });
      }
      if (req.method === 'POST') {
        const rec = {
          _id: 'r' + (++state.nextId),
          form_id: fid,
          tenant_id: state.tenant,
          created_at: new Date().toISOString(),
          response: body
        };
        state.responses[fid].push(rec);
        return send(res, 201, rec);
      }
    }

    // response update
    m = p.match(/^\/json_builder\/forms\/([^/]+)\/responses\/([^/]+)\/$/);
    if (m && (req.method === 'PATCH' || req.method === 'PUT')) {
      const list = state.responses[m[1]] || [];
      const rec = list.find(r => r._id === m[2]);
      if (!rec) return send(res, 404, { detail: 'Response not found' });
      Object.assign(rec.response, body);
      rec.updated_at = new Date().toISOString();
      return send(res, 200, rec);
    }

    send(res, 404, { detail: 'No route: ' + req.method + ' ' + p });
  });
});

server.listen(4100, () => console.log('mock AJEMS on http://localhost:4100'));

module.exports = { state, SECRET };
