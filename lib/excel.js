const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

/**
 * Reads uploaded Excel and CSV files.
 *
 * Deliberately mirrors lib/google.js: listTabs() and readTab() return the same
 * shapes, so the sync engine does not care which source a task uses.
 *
 * Uploaded files live on disk next to data.json. There is no timer for them -
 * a file only changes when the user uploads it again, so polling would be
 * reading the same bytes forever.
 */

let DIR = null;

function init(dir) {
  DIR = dir;
  try { fs.mkdirSync(DIR, { recursive: true }); } catch (e) {}
  return DIR;
}

function filePath(uploadId) {
  if (!DIR) throw new Error('Upload folder not ready.');
  if (!/^u_[A-Za-z0-9_]+\.(xlsx|xlsm|xls|csv)$/.test(uploadId)) {
    throw new Error('Bad upload id.');
  }
  return path.join(DIR, uploadId);
}

/** Saves a base64 upload and returns its id. */
function save(originalName, base64) {
  const ext = (String(originalName).match(/\.(xlsx|xlsm|xls|csv)$/i) || [])[1];
  if (!ext) throw new Error('Only .xlsx, .xlsm, .xls and .csv files can be uploaded.');

  const buf = Buffer.from(base64, 'base64');
  if (!buf.length) throw new Error('That file appears to be empty.');

  const id = 'u_' + Date.now() + '_' + Math.floor(Math.random() * 9000 + 1000) + '.' + ext.toLowerCase();
  fs.writeFileSync(path.join(DIR, id), buf);

  // Fail here rather than three screens later if it is not a real workbook.
  try {
    XLSX.readFile(path.join(DIR, id));
  } catch (e) {
    try { fs.unlinkSync(path.join(DIR, id)); } catch (e2) {}
    throw new Error('That file could not be read as a spreadsheet.');
  }
  return id;
}

function remove(uploadId) {
  try { fs.unlinkSync(filePath(uploadId)); } catch (e) {}
  return true;
}

function exists(uploadId) {
  try { return fs.existsSync(filePath(uploadId)); } catch (e) { return false; }
}

function book(uploadId) {
  const p = filePath(uploadId);
  if (!fs.existsSync(p)) {
    throw new Error('The uploaded file is no longer on disk. Upload it again.');
  }
  return XLSX.readFile(p, { cellDates: true, raw: false });
}

/** Same shape as google.listTabs(). */
function listTabs(uploadId) {
  const wb = book(uploadId);
  return {
    title: uploadId,
    tabs: wb.SheetNames.map(name => {
      const ws = wb.Sheets[name];
      const ref = ws && ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
      return {
        title: name,
        rows: ref ? ref.e.r + 1 : 0,
        cols: ref ? ref.e.c + 1 : 0
      };
    })
  };
}

function fmt(v) {
  if (v === undefined || v === null) return '';
  if (v instanceof Date) {
    const p = n => String(n).padStart(2, '0');
    return v.getFullYear() + '-' + p(v.getMonth() + 1) + '-' + p(v.getDate());
  }
  return String(v).trim();
}

/** Same shape as google.readTab(): { headers, rows } with __rowNumber. */
function readTab(uploadId, tabTitle, headerRow) {
  const wb = book(uploadId);
  const name = (tabTitle && wb.SheetNames.includes(tabTitle)) ? tabTitle : wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) throw new Error('Sheet "' + tabTitle + '" is not in this file.');

  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
  const hIndex = Math.max(0, (parseInt(headerRow, 10) || 1) - 1);
  const headers = (grid[hIndex] || []).map(h => fmt(h));

  const rows = [];
  for (let i = hIndex + 1; i < grid.length; i++) {
    const raw = grid[i] || [];
    if (!raw.some(c => fmt(c) !== '')) continue;
    const obj = { __rowNumber: i + 1 };
    headers.forEach((h, c) => { if (h) obj[h] = fmt(raw[c]); });
    rows.push(obj);
  }

  return { headers: headers.filter(Boolean), rows };
}

module.exports = { init, save, remove, exists, listTabs, readTab, filePath };
