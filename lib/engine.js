const crypto = require('crypto');
const store = require('./store');
const google = require('./google');
const excel = require('./excel');
const ajems = require('./ajems');

/**
 * Source-agnostic sync engine. It takes rows plus a mapping and knows nothing
 * about Google Sheets — the same engine drives the Excel connector.
 *
 * A task holds one or more SELECTIONS. Each selection is one sheet tab going
 * into one AJEMS form, with its own field mapping:
 *
 *   { tab, formId, formTitle, postUrl, getUrl, mapping: [{column, fieldKey}] }
 *
 * Row identity is the hard part. AJEMS has no upsert, and a sheet row number
 * is NOT stable: insert one row at the top and every row below shifts. So the
 * key comes from the row's own content:
 *
 *   identity 'key'  -> hash of the user's chosen key columns (SKU, invoice no)
 *   identity 'hash' -> hash of every mapped value in the row
 *
 * The tab is always part of the key, because each tab targets its own form —
 * the same values in two tabs are two different records in two different
 * places.
 */

/** True for uploaded files, which have no timestamp to poll. */
function isUpload(task) {
  return task.source === 'excel' && !!task.uploadId;
}

/** Reads one tab from whichever source the task uses. */
async function readSourceTab(task, tab) {
  if (isUpload(task)) return excel.readTab(task.uploadId, tab, 1);
  return google.readTab(task.spreadsheetId, tab, 1);
}

/** Tabs this task reads. Older tasks stored a single sheetTab. */
function tabsOf(task) {
  if (Array.isArray(task.selections) && task.selections.length) {
    return task.selections.map(s => s.tab);
  }
  if (Array.isArray(task.sheetTabs) && task.sheetTabs.length) return task.sheetTabs;
  return task.sheetTab ? [task.sheetTab] : [];
}

/** Selections, rebuilt from the older single-form shape when necessary. */
function selectionsOf(task) {
  if (Array.isArray(task.selections) && task.selections.length) return task.selections;
  return tabsOf(task).map(tab => ({
    tab,
    formId: task.formId || '',
    formTitle: task.formTitle || '',
    postUrl: task.postUrl || '',
    getUrl: task.getUrl || task.postUrl || '',
    mapping: task.mapping || []
  }));
}

function norm(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim().toLowerCase();
  // "1,000.00" and "1000" should not be treated as different values
  if (/^-?[\d,]+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '').replace(/\.0+$/, '');
  return s.replace(/\s+/g, ' ');
}

function sha(tab, parts) {
  return crypto.createHash('sha1')
    .update(norm(tab) + '\u0000' + parts.join('\u0000'))
    .digest('hex');
}

function rowKey(task, row, sel) {
  const mapping = (sel && sel.mapping) || task.mapping || [];
  const tab = row.__tab || (sel && sel.tab) || '';

  if (task.identity === 'key' && task.keyColumns && task.keyColumns.length) {
    const parts = task.keyColumns.map(c => norm(row[c]));
    if (parts.every(p => p === '')) return null;      // no key -> never dedupe
    return 'k:' + sha(tab, parts);
  }

  const parts = mapping.filter(m => m.column).map(m => norm(row[m.column]));
  if (parts.every(p => p === '')) return null;
  return 'h:' + sha(tab, parts);
}

/** Hash of just the payload values, used to spot "same key, changed content". */
function contentHash(task, row, sel) {
  const mapping = (sel && sel.mapping) || task.mapping || [];
  const parts = mapping.filter(m => m.column).map(m => norm(row[m.column]));
  return crypto.createHash('sha1').update(parts.join('\u0000')).digest('hex');
}

function buildPayload(task, row, sel) {
  const mapping = (sel && sel.mapping) || task.mapping || [];
  const payload = {};
  mapping.forEach(m => {
    if (!m.column || !m.fieldKey) return;
    let v = row[m.column];
    if (v === undefined || v === null) v = '';
    if (m.fieldType === 'Number') {
      const n = Number(String(v).replace(/[^0-9.\-]/g, ''));
      v = (String(v).trim() === '' || isNaN(n)) ? null : n;
    }
    payload[m.fieldKey] = v;
  });
  return payload;
}

function rowLabel(row) {
  return (row.__tab ? '"' + row.__tab + '" row ' : 'Sheet row ') + row.__rowNumber;
}

/** Small concurrency pool — the API takes one response per POST. */
async function pool(items, limit, worker) {
  const results = new Array(items.length);
  let next = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (true) {
      const i = next++;
      if (i >= items.length) return;
      results[i] = await worker(items[i], i);
    }
  });
  await Promise.all(runners);
  return results;
}

/**
 * Seed the link table from what is already in AJEMS, so a restart (or a
 * second machine) still recognises rows that were pushed before.
 */
async function seedLinks(task, sel) {
  const links = store.getLinks(task.id);
  let existing;
  try {
    existing = await ajems.listResponses(sel.getUrl || sel.postUrl);
  } catch (e) {
    return { seeded: 0, warning: e.message };
  }

  let seeded = 0;
  for (const rec of existing) {
    const resp = rec.response || rec;
    const id = rec._id || rec.id;
    if (!id || !resp) continue;

    // Rebuild the same key from the stored record by mapping field keys back
    // to columns, so the hash matches what a sheet row would produce.
    const pseudoRow = { __tab: sel.tab };
    (sel.mapping || []).forEach(m => {
      if (m.column && m.fieldKey) pseudoRow[m.column] = resp[m.fieldKey];
    });
    const k = rowKey(task, pseudoRow, sel);
    if (k && !links[k]) {
      links[k] = { responseId: id, hash: contentHash(task, pseudoRow, sel) };
      seeded++;
    }
  }
  // runTask saves once when the run finishes; writing the whole file per
  // selection here would repeat that for nothing.
  return { seeded, warning: null };
}

/**
 * Run one task: every selection in turn, each tab into its own form.
 */
async function runTask(taskId, trigger, opts) {
  opts = opts || {};
  const task = store.getTask(taskId);
  if (!task) return { error: 'Task not found.' };
  if (!task.spreadsheetId && !isUpload(task)) return { error: 'No spreadsheet selected.' };
  if (isUpload(task) && !excel.exists(task.uploadId)) {
    return { error: 'The uploaded file is no longer on disk. Upload it again.' };
  }

  const selections = selectionsOf(task).filter(s => s.tab && s.postUrl);
  if (!selections.length) return { error: 'No tab is connected to an AJEMS form yet.' };
  if (!selections.some(s => (s.mapping || []).some(m => m.column && m.fieldKey))) {
    return { error: 'Nothing is mapped yet.' };
  }

  const result = {
    created: 0, updated: 0, unchanged: 0, failed: 0,
    rows: 0, errors: [], perTab: [], startedAt: new Date().toISOString(), trigger
  };

  try {
    // An uploaded file has no modifiedTime to record; only Google Sheets do.
    const meta = isUpload(task) ? {} : await google.fileMeta(task.spreadsheetId);
    const links = store.getLinks(task.id);

    for (const sel of selections) {
      const tabResult = {
        tab: sel.tab, form: sel.formTitle,
        created: 0, updated: 0, unchanged: 0, failed: 0, rows: 0
      };

      if (!(sel.mapping || []).some(m => m.column && m.fieldKey)) {
        tabResult.skipped = 'nothing mapped';
        result.perTab.push(tabResult);
        continue;
      }

      // Seeding reads the form's existing responses so rows already in AJEMS
      // are recognised, including any added outside this app. Kept on every
      // requested run: skipping it once the table has entries would be
      // cheaper but would stop noticing records created elsewhere.
      if (opts.seedFirst) {
        const s = await seedLinks(task, sel);
        if (s.warning && result.errors.length < 8) {
          result.errors.push('Could not read existing records for "' + sel.tab + '": ' + s.warning);
        }
      }

      let rows;
      try {
        const r = await readSourceTab(task, sel.tab);
        rows = r.rows;
        rows.forEach(row => { row.__tab = sel.tab; });
      } catch (e) {
        result.failed++;
        result.errors.push('Could not read "' + sel.tab + '": ' + e.message);
        tabResult.error = e.message;
        result.perTab.push(tabResult);
        continue;
      }

      tabResult.rows = rows.length;
      result.rows += rows.length;

      await pool(rows, opts.concurrency || 4, async (row) => {
        const key = rowKey(task, row, sel);
        const payload = buildPayload(task, row, sel);
        const hash = contentHash(task, row, sel);
        const existing = key ? links[key] : null;

        try {
          if (existing) {
            if (existing.hash === hash || !task.updateExisting) {
              tabResult.unchanged++; result.unchanged++;
              return;
            }
            await ajems.updateResponse(sel.postUrl, existing.responseId, payload);
            existing.hash = hash;
            tabResult.updated++; result.updated++;
            return;
          }

          const created = await ajems.createResponse(sel.postUrl, payload);
          const id = created && (created._id || created.id);
          if (key && id) links[key] = { responseId: id, hash };
          tabResult.created++; result.created++;
        } catch (e) {
          tabResult.failed++; result.failed++;
          if (result.errors.length < 8) result.errors.push(rowLabel(row) + ': ' + e.message);
        }
      });

      result.perTab.push(tabResult);
    }

    // One write for the whole tail: links, task state and both log lines.
    store.batch(() => {
      store.save();

      task.lastModifiedTime = meta.modifiedTime || '';
      task.lastRun = new Date().toISOString();
      task.lastResult = {
        created: result.created, updated: result.updated,
        unchanged: result.unchanged, failed: result.failed,
        rows: result.rows, perTab: result.perTab
      };
      task.consecutiveFailures = result.failed && !result.created && !result.updated
        ? (task.consecutiveFailures || 0) + 1
        : 0;
      store.upsertTask(task);

      store.log(result.failed ? 'warn' : 'ok', task.name,
        `${trigger}: ${result.created} created, ${result.updated} updated, ` +
        `${result.unchanged} unchanged, ${result.failed} failed ` +
        `(${result.rows} rows across ${selections.length} tab(s))`);
      if (result.errors.length) store.log('error', task.name, result.errors[0]);
    });

  } catch (e) {
    result.errors.push(e.message);
    task.consecutiveFailures = (task.consecutiveFailures || 0) + 1;
    if (task.consecutiveFailures >= 3) {
      task.enabled = false;
      task.pausedReason = e.message;
      store.log('error', task.name, 'Paused after 3 consecutive failures: ' + e.message);
    } else {
      store.log('error', task.name, e.message);
    }
    store.upsertTask(task);
  }

  result.finishedAt = new Date().toISOString();
  return result;
}

module.exports = {
  runTask, seedLinks, buildPayload, rowKey, contentHash,
  pool, tabsOf, selectionsOf, isUpload, readSourceTab
};
