/**
 * Storage abstraction:
 *   - Vercel production  → @vercel/kv (Redis)
 *   - Local development  → JSON files in /data
 *   - Vercel without KV  → in-memory (warns on console; data resets on cold start)
 */
const IS_VERCEL = !!process.env.VERCEL;
const HAS_KV    = !!(process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN);

const DEFAULTS  = { shops: [], stock: {}, logs: [] };

// ── Local / file-system storage ───────────────────────────────────────────────
let _fsReady = false;
function ensureFs() {
  if (_fsReady || IS_VERCEL) return;
  const fs   = require('fs');
  const path = require('path');
  const dir  = path.join(__dirname, '..', 'data');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  for (const [key, def] of Object.entries(DEFAULTS)) {
    const file = path.join(dir, `${key}.json`);
    if (!fs.existsSync(file)) { fs.writeFileSync(file, JSON.stringify(def)); continue; }
    try {
      const parsed = JSON.parse(fs.readFileSync(file, 'utf-8'));
      const bad = Array.isArray(def) ? !Array.isArray(parsed) : Array.isArray(parsed);
      if (bad) fs.writeFileSync(file, JSON.stringify(def));
    } catch { fs.writeFileSync(file, JSON.stringify(def)); }
  }
  _fsReady = true;
}

function fsGet(key) {
  ensureFs();
  const { readFileSync } = require('fs');
  const { join }         = require('path');
  try { return JSON.parse(readFileSync(join(__dirname, '..', 'data', `${key}.json`), 'utf-8')); }
  catch { return JSON.parse(JSON.stringify(DEFAULTS[key])); }
}
function fsSet(key, value) {
  ensureFs();
  require('fs').writeFileSync(
    require('path').join(__dirname, '..', 'data', `${key}.json`),
    JSON.stringify(value, null, 2)
  );
}

// ── In-memory fallback ────────────────────────────────────────────────────────
const mem = { shops: [], stock: {}, logs: [] };

// ── KV (Vercel / Upstash via @vercel/kv) ─────────────────────────────────────
let _kv = null;
function getKv() {
  if (!_kv && HAS_KV) {
    const mod = require('@vercel/kv');
    // v2 exports `kv` as named; v1 used default export — handle both
    _kv = mod.kv || mod.default || mod;
  }
  return _kv;
}

// ── Public API ────────────────────────────────────────────────────────────────
async function get(key) {
  if (!IS_VERCEL) return fsGet(key);
  const kv = getKv();
  if (kv) {
    const val = await kv.get(key);
    return val ?? JSON.parse(JSON.stringify(DEFAULTS[key]));
  }
  if (!_warnedNoKv) { console.warn('[storage] KV not configured — using in-memory (data will not persist)'); _warnedNoKv = true; }
  return JSON.parse(JSON.stringify(mem[key] ?? DEFAULTS[key]));
}

async function set(key, value) {
  if (!IS_VERCEL) { fsSet(key, value); return; }
  const kv = getKv();
  if (kv) { await kv.set(key, value); return; }
  mem[key] = value;
}

let _warnedNoKv = false;
module.exports = { get, set, IS_VERCEL, HAS_KV };
