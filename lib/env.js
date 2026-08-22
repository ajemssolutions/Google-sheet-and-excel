// Minimal .env reader so the project needs no dotenv dependency.
const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, '..', '.env');

try {
  const text = fs.readFileSync(file, 'utf8');
  text.split(/\r?\n/).forEach(line => {
    const t = line.trim();
    if (!t || t.startsWith('#')) return;
    const eq = t.indexOf('=');
    if (eq === -1) return;
    const key = t.slice(0, eq).trim();
    let val = t.slice(eq + 1).trim();
    if ((val.startsWith('"') && val.endsWith('"')) ||
        (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1);
    }
    if (!(key in process.env)) process.env[key] = val;
  });
} catch (e) {
  // No .env file — the app shows a setup message in the UI instead.
}
