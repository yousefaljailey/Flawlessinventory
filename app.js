const express    = require('express');
const nodemailer = require('nodemailer');
const ExcelJS    = require('exceljs');
const crypto     = require('crypto');
const path       = require('path');
const storage    = require('./lib/storage');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Products ──────────────────────────────────────────────────────────────────
const CDN = 'https://cdn.shopify.com/s/files/1/0647/2535/2606/files';
const PRODUCTS = [
  { name: 'Monte-Carlo 30ml',      category: '30ml',         sellPrice: 49, costPrice: 15, image: `${CDN}/Monte-Carlo-bottle.jpg?v=1719491982` },
  { name: 'Monte-Carlo 100ml',     category: '100ml',        sellPrice: 89, costPrice: 28, image: `${CDN}/Monte-Carlo-bottle.jpg?v=1719491982` },
  { name: 'Summer in Paris 30ml',  category: '30ml',         sellPrice: 49, costPrice: 15, image: `${CDN}/summer-in-paris-bottle.jpg?v=1719492048` },
  { name: 'Summer in Paris 100ml', category: '100ml',        sellPrice: 89, costPrice: 28, image: `${CDN}/summer-in-paris-bottle.jpg?v=1719492048` },
  { name: 'Bedouin 30ml',          category: '30ml',         sellPrice: 49, costPrice: 15, image: `${CDN}/Bedouin-perfume.jpg?v=1719491941` },
  { name: 'Bedouin 100ml',         category: '100ml',        sellPrice: 89, costPrice: 28, image: `${CDN}/Bedouin-perfume.jpg?v=1719491941` },
  { name: 'Rub Al Khali 30ml',     category: '30ml',         sellPrice: 49, costPrice: 15, image: `${CDN}/Rub-al-khari-bottle.jpg?v=1719492019` },
  { name: 'Rub Al Khali 100ml',    category: '100ml',        sellPrice: 89, costPrice: 28, image: `${CDN}/Rub-al-khari-bottle.jpg?v=1719492019` },
  { name: 'Discovery Set',         category: 'Discovery Set',sellPrice: 39, costPrice: 12, image: `${CDN}/Untitled_design.png?v=1746527249` },
];
const PRODUCT_NAMES = PRODUCTS.map(p => p.name);

const DEFAULT_SHOPS = [
  "Ardi's Barbershop",
  "Family's Barber",
  "Hair Essence",
  "Elevated Studio",
];

const REMOVED_SHOPS = [
  "Mall of the Emirates",
  "Dubai Mall",
  "City Centre Mirdif",
];

const WAREHOUSE_ITEMS = [
  "Monte-Carlo 100ml box (empty)",
  "Bedouin 100ml box (empty)",
  "Rub Al Khali 100ml box (empty)",
  "Summer in Paris 100ml box (empty)",
  "Monte-Carlo 30ml box (empty)",
  "Bedouin 30ml box (empty)",
  "Rub Al Khali 30ml box (empty)",
  "Summer in Paris 30ml box (empty)",
  "Discovery Set box (empty)",
  "Discovery set bottles printed",
  "100ml bottles (empty)",
  "Summer in Paris 30ml bottle (empty)",
  "Bedouin 30ml bottle (empty)",
  "Monte-Carlo 30ml bottle (empty)",
  "Rub Al Khali (pliesky 100ml)",
  "Monte-Carlo (pliesky 100ml)",
  "Bedouin (pliesky 100ml)",
  "Summer in Paris (pliesky 100ml)",
  "Pumpicky",
  "Aqua",
  "Fixater",
  "Alcohol",
  "Catalogues the (big one)",
  "Lawless paper bag medium",
  "Flawless paper bag mini",
  "Thank You for your purchase CARD",
  "Thank You empty CARD",
  "tester bottles",
  "Bedouin tester package mini",
  "Rub Al Khali tester package mini",
  "Monte-Carlo tester package mini",
  "Summer in Paris tester package mini",
  "Flawless satin pouches",
  "Chocolate shipping box",
];

const LOW_STOCK_THRESHOLD = 3;
const WA_NUMBER           = '33758033774';
const REPORT_EMAIL        = 'flawlessperfumesmanagement@gmail.com';

// ── Device helpers ────────────────────────────────────────────────────────────
function parseUA(ua='') {
  let browser='Unknown', os='Unknown', device='Desktop';
  if (/Edg\/(\d+)/.test(ua))                  browser=`Edge ${RegExp.$1}`;
  else if (/Chrome\/(\d+)/.test(ua))          browser=`Chrome ${RegExp.$1}`;
  else if (/Firefox\/(\d+)/.test(ua))         browser=`Firefox ${RegExp.$1}`;
  else if (/Version\/(\d+).*Safari/.test(ua)) browser=`Safari ${RegExp.$1}`;
  else if (/Safari/.test(ua))                 browser='Safari';
  if      (/iPhone/.test(ua))   { os='iPhone';          device='Mobile';  }
  else if (/iPad/.test(ua))     { os='iPad';             device='Tablet';  }
  else if (/Android/.test(ua))  { const v=/Android ([0-9.]+)/.exec(ua); os=`Android${v?` ${v[1]}`:''}`;  device='Mobile'; }
  else if (/Windows NT 1[01]/.test(ua)) os='Windows 10/11';
  else if (/Windows/.test(ua))  os='Windows';
  else if (/Mac OS X/.test(ua)) os='macOS';
  else if (/Linux/.test(ua))    os='Linux';
  return { browser, os, device };
}

async function geolocate(ip) {
  if (!ip || ip.startsWith('127.') || ip==='::1' || ip==='unknown') return 'Local';
  try {
    const ctrl  = new AbortController();
    const timer = setTimeout(()=>ctrl.abort(), 2000);
    const r     = await fetch(`http://ip-api.com/json/${ip}?fields=city,country,status`, { signal:ctrl.signal });
    clearTimeout(timer);
    const d = await r.json();
    if (d.status==='success') return [d.city, d.country].filter(Boolean).join(', ');
  } catch { /* timeout or network error */ }
  return null;
}

async function logAccess(entry) {
  const logs = await storage.get('access_log');
  logs.unshift(entry);
  await storage.set('access_log', logs.slice(0, 500));
  // Enrich with geo in background — non-blocking
  geolocate(entry.ip).then(async loc => {
    if (!loc) return;
    const l = await storage.get('access_log');
    const i = l.findIndex(e => e.timestamp === entry.timestamp && e.ip === entry.ip);
    if (i >= 0) { l[i].location = loc; await storage.set('access_log', l); }
  }).catch(() => {});
}

// ── Auth ──────────────────────────────────────────────────────────────────────
const AUTH_SECRET = process.env.AUTH_SECRET || 'fl4wl3ss-inv-2026-secret';
const USERS       = { flawless: 'Flawless123' };

function makeToken(username) {
  const data = `${username}:${Date.now()}`;
  const sig  = crypto.createHmac('sha256', AUTH_SECRET).update(data).digest('hex');
  return Buffer.from(`${data}:${sig}`).toString('base64');
}

function verifyToken(token) {
  if (!token) return null;
  try {
    const decoded = Buffer.from(token, 'base64').toString();
    const parts   = decoded.split(':');
    if (parts.length < 3) return null;
    const sig      = parts.pop();
    const data     = parts.join(':');
    const expected = crypto.createHmac('sha256', AUTH_SECRET).update(data).digest('hex');
    if (sig !== expected) return null;
    const [username, ts] = parts;
    if (Date.now() - Number(ts) > 7 * 24 * 60 * 60 * 1000) return null;
    return username;
  } catch { return null; }
}

function requireAuth(req, res, next) {
  const token = (req.headers.authorization || '').replace('Bearer ', '').trim();
  const user  = verifyToken(token);
  if (!user) return res.status(401).json({ error: 'Unauthorized' });
  req.user = user;
  next();
}

const loginAttempts = new Map();
const MAX_ATTEMPTS  = 5;
const LOCKOUT_MS    = 15 * 60 * 1000;

function checkRateLimit(ip) {
  const now = Date.now();
  const rec = loginAttempts.get(ip) || { count: 0, first: now, locked: 0 };
  if (rec.locked && now < rec.locked) return { blocked: true, secsLeft: Math.ceil((rec.locked - now) / 1000) };
  if (now - rec.first > LOCKOUT_MS) { rec.count = 0; rec.first = now; rec.locked = 0; }
  loginAttempts.set(ip, rec);
  return { blocked: false };
}
function recordFailure(ip) {
  const rec = loginAttempts.get(ip) || { count: 0, first: Date.now(), locked: 0 };
  rec.count++;
  if (rec.count >= MAX_ATTEMPTS) rec.locked = Date.now() + LOCKOUT_MS;
  loginAttempts.set(ip, rec);
}
function clearAttempts(ip) { loginAttempts.delete(ip); }
setInterval(() => { const now = Date.now(); for (const [ip, r] of loginAttempts) if (now - r.first > LOCKOUT_MS * 2) loginAttempts.delete(ip); }, 30 * 60 * 1000);

app.post('/api/auth', (req, res) => {
  const ip = req.headers['x-forwarded-for']?.split(',')[0]?.trim() || req.socket.remoteAddress || 'unknown';
  const { blocked, secsLeft } = checkRateLimit(ip);
  if (blocked) return res.status(429).json({ error: `Too many attempts. Try again in ${secsLeft}s.` });
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'Missing credentials' });
  if (USERS[username.toLowerCase()] !== password) { recordFailure(ip); return res.status(401).json({ error: 'Invalid username or password' }); }
  clearAttempts(ip);
  const { browser, os, device } = parseUA(req.headers['user-agent'] || '');
  const { timezone='', screen='', lang='' } = req.body || {};
  const entry = {
    timestamp: new Date().toISOString(),
    user: username.toLowerCase(),
    ip, browser, os, device,
    timezone, screen, lang,
    location: null,
  };
  logAccess(entry).catch(() => {}); // fire-and-forget
  res.json({ success: true, token: makeToken(username.toLowerCase()) });
});

app.post('/api/auth/verify', (req, res) => {
  const token = (req.headers.authorization || '').replace('Bearer ', '').trim();
  res.json({ valid: !!verifyToken(token) });
});

// ── Week helpers ──────────────────────────────────────────────────────────────
function getWeekId(date = new Date()) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  const w1 = new Date(d.getFullYear(), 0, 4);
  const n  = 1 + Math.round(((d - w1) / 86400000 - 3 + (w1.getDay() + 6) % 7) / 7);
  return `${d.getFullYear()}-W${String(n).padStart(2, '0')}`;
}

function weekRange(weekId) {
  const [yr, wn] = weekId.split('-W').map(Number);
  const w1s = new Date(yr, 0, 4);
  w1s.setDate(w1s.getDate() - ((w1s.getDay() + 6) % 7));
  const start = new Date(w1s);
  start.setDate(start.getDate() + (wn - 1) * 7);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const fmt = d => d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  return `${fmt(start)} – ${fmt(end)}`;
}

// ── Stock helpers ─────────────────────────────────────────────────────────────
function normalizeEntry(e) {
  if (!e) return { initialQty: 0, soldQty: 0, topUp: 0 };
  if ('provided' in e) return { initialQty: e.provided || 0, soldQty: e.sold || 0, topUp: e.topUp || 0 };
  return { initialQty: e.initialQty || 0, soldQty: e.soldQty || 0, topUp: e.topUp || 0 };
}
function currentStock(e) { const n = normalizeEntry(e); return n.initialQty + n.topUp - n.soldQty; }

// ── WhatsApp ──────────────────────────────────────────────────────────────────
function buildWAMessage(shop, week, products) {
  const range = weekRange(week);
  const lines = [];
  for (const [product, vals] of Object.entries(products)) {
    const n = normalizeEntry(vals), cs = n.initialQty + n.topUp - n.soldQty;
    if (!n.initialQty && !n.soldQty && !n.topUp) continue;
    const parts = [];
    if (n.initialQty) parts.push(`Init: ${n.initialQty}`);
    if (n.soldQty)    parts.push(`Sold: ${n.soldQty}`);
    if (n.topUp)      parts.push(`+${n.topUp}`);
    const icon = cs < 0 ? '⚠️' : cs <= LOW_STOCK_THRESHOLD ? '🔴' : '✅';
    lines.push(`${icon} *${product}*\n   ${parts.join(' | ')} → ${cs < 0 ? `${cs} OVER-SOLD` : cs === 0 ? 'Cleared' : `${cs} left`}`);
  }
  if (!lines.length) return null;
  const time = new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
  return `🌸 *Flawless Inventory Update*\n\n📍 *${shop}*\n📅 ${week}  ·  ${range}\n\n${lines.join('\n\n')}\n\n_Saved at ${time}_`;
}

async function sendWA(message) {
  const key = process.env.WA_API_KEY;
  if (!key) return;
  try { await fetch(`https://api.callmebot.com/whatsapp.php?phone=${WA_NUMBER}&text=${encodeURIComponent(message)}&apikey=${key}`); } catch { /* silent */ }
}

async function addLog(entry) {
  const logs = await storage.get('logs');
  logs.unshift({ ...entry, timestamp: new Date().toISOString() });
  await storage.set('logs', logs.slice(0, 2000));
}

// ── API: Access Log ───────────────────────────────────────────────────────────
app.get('/api/access-log', requireAuth, async (_, res) => {
  try { res.json(await storage.get('access_log') || []); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/access-log', requireAuth, async (_, res) => {
  try { await storage.set('access_log', []); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Shops ────────────────────────────────────────────────────────────────
function filterShops(shops) {
  return (shops || []).filter(s => s && !REMOVED_SHOPS.includes(s));
}

app.get('/api/shops', requireAuth, async (_, res) => {
  try {
    let shops = await storage.get('shops') || [];
    const originalShops = [...shops];
    shops = filterShops(shops);
    const init = await storage.get('shops_initialized');
    if (!shops || shops.length === 0) {
      if (!init || originalShops.some(s => REMOVED_SHOPS.includes(s))) {
        shops = [...DEFAULT_SHOPS];
        await storage.set('shops', shops);
        await storage.set('shops_initialized', true);
        await addLog({ action: 'INIT_SHOPS', count: shops.length });
      }
    }
    res.json(shops || []);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/shops', requireAuth, async (req, res) => {
  try {
    const name = req.body.name?.trim();
    if (!name) return res.status(400).json({ error: 'Shop name required' });
    const shops = await storage.get('shops');
    if (shops.includes(name)) return res.status(400).json({ error: 'Shop already exists' });
    shops.push(name);
    await storage.set('shops', shops);
    await addLog({ action: 'ADD_SHOP', shop: name });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/shops/:name', requireAuth, async (req, res) => {
  try {
    const name = decodeURIComponent(req.params.name);
    const shops = await storage.get('shops');
    await storage.set('shops', shops.filter(s => s !== name));
    await addLog({ action: 'REMOVE_SHOP', shop: name });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

function normalizeWarehouse(warehouse = {}) {
  const normalized = { ...warehouse };
  for (const item of WAREHOUSE_ITEMS) {
    if (!Object.prototype.hasOwnProperty.call(normalized, item)) {
      normalized[item] = { quantity: 0, savedAt: null };
    }
  }
  return normalized;
}

app.get('/api/warehouse', requireAuth, async (_, res) => {
  try {
    let warehouse = await storage.get('warehouse');
    warehouse = normalizeWarehouse(warehouse);
    await storage.set('warehouse', warehouse);
    res.json(warehouse);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/warehouse', requireAuth, async (req, res) => {
  try {
    const { items } = req.body || {};
    if (!items || typeof items !== 'object') return res.status(400).json({ error: 'Items object required' });
    const warehouse = await storage.get('warehouse') || {};
    const savedAt = new Date().toISOString();
    for (const [name, vals] of Object.entries(items)) {
      warehouse[name] = {
        quantity: Math.max(0, Number(vals.quantity) || 0),
        savedAt,
      };
    }
    await storage.set('warehouse', warehouse);
    await addLog({ action: 'WAREHOUSE_SAVE', items: Object.keys(items).length });
    res.json({ success: true, warehouse });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Contacts ─────────────────────────────────────────────────────────────
app.get('/api/contacts', requireAuth, async (_, res) => {
  try { res.json((await storage.get('contacts')) || {}); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/contacts', requireAuth, async (req, res) => {
  try {
    const { shop, contact, email } = req.body || {};
    if (!shop) return res.status(400).json({ error: 'Shop required' });
    const contacts = (await storage.get('contacts')) || {};
    contacts[shop] = { contact: contact || '', email: email || '', updatedAt: new Date().toISOString() };
    await storage.set('contacts', contacts);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Products ─────────────────────────────────────────────────────────────
app.get('/api/products', requireAuth, (_, res) => res.json(PRODUCTS));

// ── API: Stock ────────────────────────────────────────────────────────────────
app.get('/api/stock/:week', requireAuth, async (req, res) => {
  try { const stock = await storage.get('stock'); res.json(stock[req.params.week] || {}); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/stock/bulk', requireAuth, async (req, res) => {
  try {
    const { shop, week, products } = req.body;
    if (!shop || !week || !products) return res.status(400).json({ error: 'Missing fields' });
    const stock = await storage.get('stock');
    if (!stock[week])       stock[week] = {};
    if (!stock[week][shop]) stock[week][shop] = {};
    const saved = new Date().toISOString();
    for (const [product, vals] of Object.entries(products)) {
      stock[week][shop][product] = {
        initialQty: Math.max(0, Number(vals.initialQty) || 0),
        soldQty:    Math.max(0, Number(vals.soldQty)    || 0),
        topUp:      Math.max(0, Number(vals.topUp)      || 0),
        savedAt: saved,
      };
    }
    await storage.set('stock', stock);
    await addLog({ action: 'STOCK_SAVE', shop, week, products: Object.keys(products).length });
    const waMsg = buildWAMessage(shop, week, products);
    if (waMsg) sendWA(waMsg).catch(() => {});
    const waUrl = waMsg ? `https://wa.me/${WA_NUMBER}?text=${encodeURIComponent(waMsg)}` : null;
    res.json({ success: true, waUrl });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Dashboard ────────────────────────────────────────────────────────────
app.get('/api/dashboard/:week', requireAuth, async (req, res) => {
  try {
    const { week } = req.params;
    const [stock, shops] = await Promise.all([storage.get('stock'), storage.get('shops')]);
    const filteredShops = filterShops(shops);
    const weekData = stock[week] || {};

    let totalDelivered=0, totalSold=0, totalRemaining=0, totalRevenue=0, totalCOGS=0;
    const lowStock=[], overSold=[], productSummary={}, shopStatuses=[];

    for (const shop of filteredShops) {
      const shopData = weekData[shop] || {};
      let hasAny=false, shopSold=0, shopRemain=0, shopOk=true, lastSaved=null;

      for (const p of PRODUCTS) {
        const raw  = shopData[p.name];
        const norm = normalizeEntry(raw);
        const rem  = currentStock(raw);
        const hd   = norm.initialQty > 0 || norm.soldQty > 0 || norm.topUp > 0;
        if (hd) {
          hasAny = true;
          totalDelivered += norm.initialQty + norm.topUp;
          totalSold      += norm.soldQty;
          totalRevenue   += norm.soldQty * (p.sellPrice || 0);
          totalCOGS      += norm.soldQty * (p.costPrice || 0);
          shopSold       += norm.soldQty;
          shopRemain     += rem;
          if (!productSummary[p.name]) productSummary[p.name] = { delivered:0, sold:0, remaining:0, revenue:0, cogs:0, category:p.category, image:p.image };
          productSummary[p.name].delivered  += norm.initialQty + norm.topUp;
          productSummary[p.name].sold       += norm.soldQty;
          productSummary[p.name].remaining  += rem;
          productSummary[p.name].revenue    += norm.soldQty * (p.sellPrice || 0);
          productSummary[p.name].cogs       += norm.soldQty * (p.costPrice || 0);
          if (rem < 0)                        { overSold.push({ shop, product:p.name, remaining:rem }); shopOk=false; }
          else if (rem <= LOW_STOCK_THRESHOLD){ lowStock.push({ shop, product:p.name, remaining:rem }); shopOk=false; }
          if (!lastSaved || (raw?.savedAt && raw.savedAt > lastSaved)) lastSaved = raw?.savedAt;
        }
      }
      totalRemaining += shopRemain;
      shopStatuses.push({ shop, hasData:hasAny, sold:shopSold, remaining:shopRemain, ok:shopOk, lastSaved });
    }

    const grossProfit = totalRevenue - totalCOGS;
    const margin = totalRevenue > 0 ? Math.round((grossProfit / totalRevenue) * 100) : 0;

    res.json({ week, range:weekRange(week), totalShops:filteredShops.length,
      shopsReporting: shopStatuses.filter(s=>s.hasData).length,
      totalDelivered, totalSold, totalRemaining,
      totalRevenue, totalCOGS, grossProfit, margin,
      lowStockCount:lowStock.length, overSoldCount:overSold.length,
      lowStock, overSold, productSummary, shopStatuses });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Logs ─────────────────────────────────────────────────────────────────
app.get('/api/logs', requireAuth, async (req, res) => {
  try {
    const limit = Math.min(Number(req.query.limit) || 200, 2000);
    const logs  = await storage.get('logs');
    res.json(logs.slice(0, limit));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/week', (_, res) => { const week = getWeekId(); res.json({ week, range: weekRange(week) }); });

// ── Report: Excel ─────────────────────────────────────────────────────────────
async function buildWorkbook(week) {
  const [stock, shops] = await Promise.all([storage.get('stock'), storage.get('shops')]);
  const filteredShops = filterShops(shops);
  const weekData = stock[week] || {};

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Flawless Inventory'; wb.created = new Date();

  const ws = wb.addWorksheet(`Week ${week}`, { pageSetup: { fitToPage:true, fitToWidth:1 } });
  ws.views = [{ state:'frozen', ySplit:2 }];
  ws.properties.defaultRowHeight = 18;

  ws.mergeCells('A1:J1');
  const title = ws.getCell('A1');
  title.value = `FLAWLESS PERFUMES — Stock Report ${week}  (${weekRange(week)})`;
  title.font  = { name:'Calibri', size:14, bold:true, color:{argb:'FF1A1A1A'} };
  title.fill  = { type:'pattern', pattern:'solid', fgColor:{argb:'FFE1CBBD'} };
  title.alignment = { horizontal:'center', vertical:'middle' };
  ws.getRow(1).height = 28;

  const headers = ['Shop','Product','Category','Initial Qty','Qty Sold','Top-Up','Remaining','Sell Price (€)','Revenue (€)','Status'];
  const widths = [28,26,16,13,11,11,13,14,13,14];
  ws.columns = headers.map((h,i) => ({ header:h, key:h, width:widths[i], style:{ alignment:{ horizontal:i<=2?'left':'center', vertical:'middle' } } }));
  const hRow = ws.getRow(2);
  hRow.values = headers;
  hRow.eachCell(c => {
    c.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF1A1A1A'} };
    c.font = { name:'Calibri', bold:true, color:{argb:'FFE1CBBD'}, size:11 };
    c.alignment = { horizontal:'center', vertical:'middle' };
    c.border = { top:{style:'thin',color:{argb:'FFB99A75'}}, left:{style:'thin',color:{argb:'FFB99A75'}}, bottom:{style:'thin',color:{argb:'FFB99A75'}}, right:{style:'thin',color:{argb:'FFB99A75'}} };
  });
  ws.getRow(2).height = 22;

  let rowIdx=3, hasMismatches=false;
  filteredShops.forEach(shop => {
    PRODUCT_NAMES.forEach(product => {
      const pObj = PRODUCTS.find(p => p.name === product);
      const raw  = weekData[shop]?.[product];
      const e    = normalizeEntry(raw);
      const cs   = currentStock(raw);
      const hasD = e.initialQty > 0 || e.soldQty > 0 || e.topUp > 0;
      const bad  = hasD && cs < 0;
      const low  = hasD && !bad && cs <= LOW_STOCK_THRESHOLD;
      const rev  = e.soldQty * (pObj?.sellPrice || 0);
      if (bad) hasMismatches = true;

      const row = ws.getRow(rowIdx);
      row.values = [shop, product, pObj?.category||'', e.initialQty, e.soldQty, e.topUp,
        hasD ? cs : '—', pObj?.sellPrice||0, hasD ? rev : '—',
        !hasD ? 'Not entered' : bad ? '⚠ OVER-SOLD' : low ? '⚠ LOW' : '✓ OK'];
      row.eachCell((c, col) => {
        c.font = { name:'Calibri', size:10, color:{argb: bad?'FFFFFFFF':'FF1A1A1A'} };
        c.alignment = { horizontal: col<=3?'left':'center', vertical:'middle' };
        c.fill = bad
          ? { type:'pattern', pattern:'solid', fgColor:{argb:'FFF83A3A'} }
          : low
            ? { type:'pattern', pattern:'solid', fgColor:{argb:'FFFFF3CD'} }
            : { type:'pattern', pattern:'solid', fgColor:{argb: rowIdx%2===0?'FFF9F0EC':'FFFFFFFF'} };
        c.border = { top:{style:'thin',color:{argb:'FFB99A75'}}, left:{style:'thin',color:{argb:'FFB99A75'}}, bottom:{style:'thin',color:{argb:'FFB99A75'}}, right:{style:'thin',color:{argb:'FFB99A75'}} };
        if (bad) c.font = { ...c.font, bold:true };
      });
      row.getCell(10).font = { name:'Calibri', size:10, bold:bad||low,
        color:{argb: bad?'FFFFFFFF':low?'FFCC8800':!hasD?'FF888888':'FF00A341'} };
      row.height = 18;
      rowIdx++;
    });
  });

  // Financial summary row
  rowIdx++;
  let totalRev=0, totalCOGS_=0;
  filteredShops.forEach(shop => {
    PRODUCT_NAMES.forEach(product => {
      const pObj = PRODUCTS.find(p=>p.name===product);
      const raw  = weekData[shop]?.[product];
      const e    = normalizeEntry(raw);
      totalRev   += e.soldQty * (pObj?.sellPrice||0);
      totalCOGS_ += e.soldQty * (pObj?.costPrice||0);
    });
  });

  ws.mergeCells(`A${rowIdx}:G${rowIdx}`);
  ws.getCell(`A${rowIdx}`).value = `Week ${week} Financial Summary`;
  ws.getCell(`A${rowIdx}`).font  = { name:'Calibri', bold:true, size:11, color:{argb:'FF1A1A1A'} };
  ws.getCell(`A${rowIdx}`).fill  = { type:'pattern', pattern:'solid', fgColor:{argb:'FFE1CBBD'} };
  ws.getCell(`A${rowIdx}`).alignment = { horizontal:'center' };

  rowIdx++;
  const summaryRows = [
    ['Total Revenue', `€${totalRev.toFixed(2)}`],
    ['Total COGS', `€${totalCOGS_.toFixed(2)}`],
    ['Gross Profit', `€${(totalRev-totalCOGS_).toFixed(2)}`],
    ['Margin %', totalRev>0 ? `${Math.round((totalRev-totalCOGS_)/totalRev*100)}%` : '—'],
  ];
  summaryRows.forEach(([label, val]) => {
    ws.getCell(`A${rowIdx}`).value = label;
    ws.getCell(`B${rowIdx}`).value = val;
    ws.getCell(`A${rowIdx}`).font = { name:'Calibri', size:10, bold:true };
    ws.getCell(`B${rowIdx}`).font = { name:'Calibri', size:10 };
    rowIdx++;
  });

  ws.mergeCells(`A${rowIdx}:J${rowIdx}`);
  const footer = ws.getCell(`A${rowIdx}`);
  footer.value = hasMismatches
    ? '⚠  Rows in RED = over-sold. Please review immediately.'
    : `✓  All stock counts balance for week ${week}.`;
  footer.font = { name:'Calibri', italic:true, size:10, color:{argb: hasMismatches?'FFF83A3A':'FF00A341'} };
  footer.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFE1CBBD'} };
  footer.alignment = { horizontal:'center' };

  return wb;
}

app.get('/api/report/:week', requireAuth, async (req, res) => {
  try {
    const wb = await buildWorkbook(req.params.week);
    res.setHeader('Content-Disposition', `attachment; filename="Flawless-Stock-${req.params.week}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Email ─────────────────────────────────────────────────────────────────────
async function sendWeeklyReport(week) {
  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) throw new Error('EMAIL_USER / EMAIL_PASS not configured.');
  const wb = await buildWorkbook(week), buffer = await wb.xlsx.writeBuffer(), range = weekRange(week);
  const transporter = nodemailer.createTransport({ service:'gmail', auth:{ user:process.env.EMAIL_USER, pass:process.env.EMAIL_PASS } });
  await transporter.sendMail({
    from: `"Flawless Inventory" <${process.env.EMAIL_USER}>`,
    to:   REPORT_EMAIL,
    subject: `Flawless Perfumes — Weekly Stock Report ${week}`,
    html: `<div style="font-family:sans-serif;max-width:640px;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e0d0c8">
      <div style="background:#E1CBBD;padding:28px;text-align:center"><img src="https://flawlessperfumes.com/cdn/shop/files/Flawless_Logo.svg?v=1711183368&width=220" style="height:56px;width:auto" alt="Flawless"></div>
      <div style="padding:28px">
        <h2 style="font-size:20px;color:#1A1A1A;margin:0 0 6px">Weekly Stock Report — ${week}</h2>
        <p style="color:#9B6D57;margin:0 0 20px;font-size:13px">${range}</p>
        <p style="line-height:1.7;color:#1A1A1A">Find the full stock report attached. Red rows = over-sold · Yellow rows = low stock (≤${LOW_STOCK_THRESHOLD} units).</p>
        <p style="line-height:1.7;color:#1A1A1A">The Excel file includes revenue and cost of sales per product.</p>
      </div>
      <div style="background:#E1CBBD;padding:14px;text-align:center;font-size:11px;color:#9B6D57">Flawless Inventory · <a href="https://wa.me/${WA_NUMBER}" style="color:#9B6D57">WhatsApp</a></div>
    </div>`,
    attachments: [{ filename:`Flawless-Stock-${week}.xlsx`, content:buffer, contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }],
  });
  await addLog({ action:'EMAIL_SENT', week, to:REPORT_EMAIL });
}

app.post('/api/send-report', requireAuth, async (req, res) => {
  const week = req.body.week || getWeekId();
  try { await sendWeeklyReport(week); res.json({ success:true, message:`Report for ${week} sent to ${REPORT_EMAIL}` }); }
  catch (e) { res.status(500).json({ error:e.message }); }
});

// ── API: Admin Reset ──────────────────────────────────────────────────────────
app.post('/api/admin/reset', requireAuth, async (_, res) => {
  try {
    await Promise.all([
      storage.set('stock', {}),
      storage.set('logs', []),
      storage.set('contacts', {}),
      storage.set('shops', []),
      storage.set('shops_initialized', false),
      storage.set('access_log', []),
      storage.set('warehouse', {}),
    ]);
    await addLog({ action: 'ADMIN_RESET', note: 'All data cleared' });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/cron', async (req, res) => {
  const secret = process.env.CRON_SECRET;
  if (secret && req.headers.authorization !== `Bearer ${secret}`) return res.status(401).json({ error:'Unauthorized' });
  const week = getWeekId();
  try { await sendWeeklyReport(week); await addLog({ action:'CRON_EMAIL', week }); res.json({ success:true, week }); }
  catch (e) { console.error('[cron]', e.message); res.status(500).json({ error:e.message }); }
});

module.exports = { app, getWeekId, sendWeeklyReport };
