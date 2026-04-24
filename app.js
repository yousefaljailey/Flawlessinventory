const express   = require('express');
const nodemailer = require('nodemailer');
const ExcelJS   = require('exceljs');
const crypto    = require('crypto');
const path      = require('path');
const storage   = require('./lib/storage');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Constants ─────────────────────────────────────────────────────────────────
const PRODUCTS = [
  'Monte-Carlo 30ml',
  'Monte-Carlo 100ml',
  'Summer in Paris 30ml',
  'Bedouin 100ml',
  'Rub Al Khali 30ml',
  'Discovery Set'
];

const DEFAULT_SHOPS = [
  "Ardi's Barbershop",
  "Family's Barber",
  "Hair Essence",
  "Romana Paulen",
  "Pilateez",
  "Elevated Studio"
];

const WA_NUMBER = '33758033774';
const REPORT_EMAIL = 'flawlessperfumesmanagement@gmail.com';

// ── Auth ──────────────────────────────────────────────────────────────────────
const AUTH_SECRET = process.env.AUTH_SECRET || 'fl4wl3ss-inv-2026-secret';
const USERS = { flawless: 'Flawless123' };

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
    const sig  = parts.pop();
    const data = parts.join(':');
    const expected = crypto.createHmac('sha256', AUTH_SECRET).update(data).digest('hex');
    if (sig !== expected) return null;
    const [username, ts] = parts;
    if (Date.now() - Number(ts) > 7 * 24 * 60 * 60 * 1000) return null; // 7 days
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

app.post('/api/auth', (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'Missing credentials' });
  if (USERS[username.toLowerCase()] !== password) {
    return res.status(401).json({ error: 'Invalid username or password' });
  }
  res.json({ success: true, token: makeToken(username.toLowerCase()) });
});

app.post('/api/auth/verify', (req, res) => {
  const token = (req.headers.authorization || '').replace('Bearer ', '').trim();
  const user  = verifyToken(token);
  res.json({ valid: !!user, user });
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

// ── Normalize stock entry (handles old + new format) ─────────────────────────
function normalizeEntry(e) {
  if (!e) return { initialQty: 0, soldQty: 0, topUp: 0 };
  // Old format migration: provided → initialQty, sold → soldQty
  if ('provided' in e) return { initialQty: e.provided || 0, soldQty: e.sold || 0, topUp: e.topUp || 0 };
  return { initialQty: e.initialQty || 0, soldQty: e.soldQty || 0, topUp: e.topUp || 0 };
}

function currentStock(e) {
  const n = normalizeEntry(e);
  return n.initialQty + n.topUp - n.soldQty;
}

async function addLog(entry) {
  const logs = await storage.get('logs');
  logs.unshift({ ...entry, timestamp: new Date().toISOString() });
  await storage.set('logs', logs.slice(0, 2000));
}

// ── API: Shops ────────────────────────────────────────────────────────────────
app.get('/api/shops', requireAuth, async (_, res) => {
  try {
    let shops = await storage.get('shops');
    // Seed defaults on first run
    if (!shops || shops.length === 0) {
      const initialized = await storage.get('shops_initialized');
      if (!initialized) {
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
    const name  = decodeURIComponent(req.params.name);
    const shops = await storage.get('shops');
    await storage.set('shops', shops.filter(s => s !== name));
    await addLog({ action: 'REMOVE_SHOP', shop: name });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Products ─────────────────────────────────────────────────────────────
app.get('/api/products', requireAuth, (_, res) => res.json(PRODUCTS));

// ── API: Stock ────────────────────────────────────────────────────────────────
app.get('/api/stock/:week', requireAuth, async (req, res) => {
  try {
    const stock = await storage.get('stock');
    res.json(stock[req.params.week] || {});
  } catch (e) { res.status(500).json({ error: e.message }); }
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
        savedAt: saved
      };
    }
    await storage.set('stock', stock);
    await addLog({ action: 'STOCK_SAVE', shop, week, products: Object.keys(products).length });
    res.json({ success: true });
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

// ── API: Week info ────────────────────────────────────────────────────────────
app.get('/api/week', (_, res) => {
  const week = getWeekId();
  res.json({ week, range: weekRange(week) });
});

// ── Report: build Excel workbook ──────────────────────────────────────────────
async function buildWorkbook(week) {
  const [stock, shops] = await Promise.all([storage.get('stock'), storage.get('shops')]);
  const weekData = stock[week] || {};

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Flawless Inventory';
  wb.created = new Date();

  const ws = wb.addWorksheet(`Week ${week}`, { pageSetup: { fitToPage: true, fitToWidth: 1 } });

  // Title
  ws.mergeCells('A1:G1');
  const title = ws.getCell('A1');
  title.value     = `FLAWLESS PERFUMES — Stock Report ${week}  (${weekRange(week)})`;
  title.font      = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF1A1A1A' } };
  title.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE1CBBD' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;

  const headers = ['Shop', 'Product', 'Initial Qty Given', 'Qty Sold', 'Top-Up Qty', 'Current Stock', 'Status'];
  ws.columns = headers.map((h, i) => ({ header: h, key: h, width: [28, 26, 18, 14, 14, 16, 14][i] }));
  const hRow = ws.getRow(2);
  hRow.values = headers;
  hRow.eachCell(c => {
    c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A1A' } };
    c.font      = { name: 'Calibri', bold: true, color: { argb: 'FFE1CBBD' }, size: 11 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.getRow(2).height = 22;

  let rowIdx = 3, hasMismatches = false;
  shops.forEach(shop => {
    PRODUCTS.forEach(product => {
      const raw = weekData[shop]?.[product];
      const e   = normalizeEntry(raw);
      const cs  = currentStock(raw);
      const hasData   = e.initialQty > 0 || e.soldQty > 0 || e.topUp > 0;
      const bad       = hasData && cs < 0;
      if (bad) hasMismatches = true;

      const row = ws.getRow(rowIdx);
      row.values = [shop, product, e.initialQty, e.soldQty, e.topUp, hasData ? cs : '—',
        !hasData ? 'Not entered' : bad ? '⚠ OVER-SOLD' : '✓ OK'];
      row.eachCell((c, col) => {
        c.font      = { name: 'Calibri', size: 10, color: { argb: bad ? 'FFFFFFFF' : 'FF1A1A1A' } };
        c.alignment = { horizontal: col <= 2 ? 'left' : 'center', vertical: 'middle' };
        c.fill      = bad
          ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF83A3A' } }
          : { type: 'pattern', pattern: 'solid', fgColor: { argb: rowIdx % 2 === 0 ? 'FFF9F0EC' : 'FFFFFFFF' } };
        if (bad) c.font = { ...c.font, bold: true };
      });
      row.getCell(7).font = {
        name: 'Calibri', size: 10, bold: bad,
        color: { argb: bad ? 'FFFFFFFF' : !hasData ? 'FF888888' : 'FF00A341' }
      };
      row.height = 18;
      rowIdx++;
    });
  });

  rowIdx++;
  ws.mergeCells(`A${rowIdx}:G${rowIdx}`);
  const footer = ws.getCell(`A${rowIdx}`);
  footer.value     = hasMismatches
    ? '⚠  Rows in RED indicate over-sold stock. Please review immediately.'
    : `✓  All stock counts balance correctly for week ${week}.`;
  footer.font      = { name: 'Calibri', italic: true, size: 10, color: { argb: hasMismatches ? 'FFF83A3A' : 'FF00A341' } };
  footer.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE1CBBD' } };
  footer.alignment = { horizontal: 'center' };

  return wb;
}

// ── API: Download Excel ───────────────────────────────────────────────────────
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
  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
    throw new Error('EMAIL_USER / EMAIL_PASS not set in environment variables.');
  }
  const wb     = await buildWorkbook(week);
  const buffer = await wb.xlsx.writeBuffer();
  const range  = weekRange(week);

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS }
  });
  await transporter.sendMail({
    from:    `"Flawless Inventory" <${process.env.EMAIL_USER}>`,
    to:      REPORT_EMAIL,
    subject: `Flawless Perfumes — Weekly Stock Report ${week}`,
    html: `
      <div style="font-family:sans-serif;max-width:640px;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e0d0c8">
        <div style="background:#E1CBBD;padding:32px;text-align:center">
          <img src="https://flawlessperfumes.com/cdn/shop/files/Flawless_Logo.svg?v=1711183368&width=220" style="height:60px;width:auto" alt="Flawless">
        </div>
        <div style="padding:32px">
          <h2 style="font-size:20px;font-weight:700;color:#1A1A1A;margin:0 0 8px">Weekly Stock Report — ${week}</h2>
          <p style="color:#9B6D57;margin:0 0 24px;font-size:14px">${range}</p>
          <p style="line-height:1.7;color:#1A1A1A">Please find the weekly inventory report attached as an Excel spreadsheet.</p>
          <p style="line-height:1.7;color:#1A1A1A">Rows <strong style="color:#F83A3A">highlighted in red</strong> indicate over-sold stock that requires immediate attention.</p>
          <div style="margin:24px 0;padding:16px;background:#fdf7f4;border-radius:8px;border:1px solid #e0d0c8">
            <p style="margin:0;font-size:13px;color:#5A3A2A">
              📧 ${REPORT_EMAIL}<br>
              💬 WhatsApp: <a href="https://wa.me/${WA_NUMBER}" style="color:#9B6D57">+33 7 58 03 37 74</a>
            </p>
          </div>
        </div>
        <div style="background:#E1CBBD;padding:16px;text-align:center">
          <p style="margin:0;font-size:11px;color:#9B6D57">Sent automatically by Flawless Inventory Management</p>
        </div>
      </div>`,
    attachments: [{ filename: `Flawless-Stock-${week}.xlsx`, content: buffer,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
  });
  await addLog({ action: 'EMAIL_SENT', week, to: REPORT_EMAIL });
}

app.post('/api/send-report', requireAuth, async (req, res) => {
  const week = req.body.week || getWeekId();
  try {
    await sendWeeklyReport(week);
    res.json({ success: true, message: `Report for ${week} sent to ${REPORT_EMAIL}` });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Vercel Cron (every Sunday 9 AM) ──────────────────────────────────────────
app.get('/api/cron', async (req, res) => {
  const secret = process.env.CRON_SECRET;
  if (secret && req.headers.authorization !== `Bearer ${secret}`) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const week = getWeekId();
  try {
    await sendWeeklyReport(week);
    await addLog({ action: 'CRON_EMAIL', week });
    res.json({ success: true, week });
  } catch (e) {
    console.error('[cron]', e.message);
    res.status(500).json({ error: e.message });
  }
});

module.exports = { app, getWeekId, sendWeeklyReport };
