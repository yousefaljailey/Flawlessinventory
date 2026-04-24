const express  = require('express');
const nodemailer = require('nodemailer');
const ExcelJS  = require('exceljs');
const path     = require('path');
const storage  = require('./lib/storage');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const PRODUCTS = [
  'Monte-Carlo 30ml',
  'Monte-Carlo 100ml',
  'Summer in Paris 30ml',
  'Bedouin 100ml',
  'Rub Al Khali 30ml',
  'Discovery Set'
];

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

async function addLog(entry) {
  const logs = await storage.get('logs');
  logs.unshift({ ...entry, timestamp: new Date().toISOString() });
  await storage.set('logs', logs.slice(0, 2000));
}

// ── API: Shops ────────────────────────────────────────────────────────────────
app.get('/api/shops', async (_, res) => {
  try { res.json(await storage.get('shops')); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/shops', async (req, res) => {
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

app.delete('/api/shops/:name', async (req, res) => {
  try {
    const name  = decodeURIComponent(req.params.name);
    const shops = await storage.get('shops');
    await storage.set('shops', shops.filter(s => s !== name));
    await addLog({ action: 'REMOVE_SHOP', shop: name });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Products ─────────────────────────────────────────────────────────────
app.get('/api/products', (_, res) => res.json(PRODUCTS));

// ── API: Stock ────────────────────────────────────────────────────────────────
app.get('/api/stock/:week', async (req, res) => {
  try {
    const stock = await storage.get('stock');
    res.json(stock[req.params.week] || {});
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/stock/bulk', async (req, res) => {
  try {
    const { shop, week, products } = req.body;
    if (!shop || !week || !products) return res.status(400).json({ error: 'Missing fields' });

    const stock = await storage.get('stock');
    if (!stock[week])       stock[week] = {};
    if (!stock[week][shop]) stock[week][shop] = {};

    const saved = new Date().toISOString();
    for (const [product, vals] of Object.entries(products)) {
      stock[week][shop][product] = {
        provided:  Math.max(0, Number(vals.provided)  || 0),
        sold:      Math.max(0, Number(vals.sold)       || 0),
        remaining: Math.max(0, Number(vals.remaining)  || 0),
        savedAt: saved
      };
    }
    await storage.set('stock', stock);
    await addLog({ action: 'STOCK_SAVE', shop, week, products: Object.keys(products).length });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── API: Logs ─────────────────────────────────────────────────────────────────
app.get('/api/logs', async (req, res) => {
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

  ws.mergeCells('A1:G1');
  const title = ws.getCell('A1');
  title.value     = `FLAWLESS PERFUMES — Stock Report ${week}  (${weekRange(week)})`;
  title.font      = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFC9A84C' } };
  title.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A0806' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;

  const headers = ['Shop', 'Product', 'Provided', 'Sold', 'Remaining', 'Expected', 'Status'];
  ws.columns = headers.map((h, i) => ({ header: h, key: h.toLowerCase(), width: [28,26,12,12,12,12,14][i] }));
  const hRow = ws.getRow(2);
  hRow.values = headers;
  hRow.eachCell(c => {
    c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1510' } };
    c.font      = { name: 'Calibri', bold: true, color: { argb: 'FFC9A84C' }, size: 11 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border    = { bottom: { style: 'thin', color: { argb: 'FFC9A84C' } } };
  });
  ws.getRow(2).height = 22;

  let rowIdx = 3, hasMismatches = false;
  shops.forEach(shop => {
    PRODUCTS.forEach(product => {
      const e        = weekData[shop]?.[product] || {};
      const provided  = e.provided  || 0;
      const sold      = e.sold      || 0;
      const remaining = e.remaining || 0;
      const expected  = provided - sold;
      const hasData   = provided > 0 || sold > 0 || remaining > 0;
      const bad       = hasData && remaining !== expected;
      if (bad) hasMismatches = true;

      const row = ws.getRow(rowIdx);
      row.values = [shop, product, provided, sold, remaining, expected,
        !hasData ? 'Not entered' : bad ? '⚠ MISMATCH' : '✓ OK'];
      row.eachCell((c, col) => {
        c.font      = { name: 'Calibri', size: 10, color: { argb: bad ? 'FFFFFFFF' : 'FFF0E6D0' } };
        c.alignment = { horizontal: col <= 2 ? 'left' : 'center', vertical: 'middle' };
        c.fill      = bad
          ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC2222' } }
          : { type: 'pattern', pattern: 'solid', fgColor: { argb: rowIdx % 2 === 0 ? 'FF100D0A' : 'FF0A0806' } };
        if (bad) c.font = { ...c.font, bold: true };
      });
      row.getCell(7).font = {
        name: 'Calibri', size: 10, bold: bad,
        color: { argb: bad ? 'FFFFFFFF' : !hasData ? 'FF888888' : 'FF4CAF8A' }
      };
      row.height = 18;
      rowIdx++;
    });
  });

  rowIdx++;
  ws.mergeCells(`A${rowIdx}:G${rowIdx}`);
  const footer = ws.getCell(`A${rowIdx}`);
  footer.value     = hasMismatches
    ? '⚠  This report contains discrepancies. Rows highlighted in RED indicate provided − sold ≠ remaining.'
    : `✓  All stock counts balance correctly for week ${week}.`;
  footer.font      = { name: 'Calibri', italic: true, size: 10, color: { argb: hasMismatches ? 'FFFF5252' : 'FF4CAF8A' } };
  footer.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1510' } };
  footer.alignment = { horizontal: 'center' };

  return wb;
}

// ── API: Download Excel report ────────────────────────────────────────────────
app.get('/api/report/:week', async (req, res) => {
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
    to:      'flawlessperfumesmanagement@gmail.com',
    subject: `Flawless Perfumes — Weekly Stock Report ${week}`,
    html: `
      <div style="font-family:sans-serif;max-width:600px;background:#0A0806;color:#F0E6D0;padding:32px;border-radius:10px">
        <h1 style="font-family:Georgia,serif;color:#C9A84C;letter-spacing:.15em;font-size:24px;margin:0 0 4px">FLAWLESS</h1>
        <p style="color:#8A7860;margin:0 0 24px;font-size:12px;letter-spacing:.1em">PERFUMES</p>
        <h2 style="font-size:18px;font-weight:600;color:#F0E6D0;margin:0 0 8px">Weekly Stock Report — ${week}</h2>
        <p style="color:#8A7860;margin:0 0 24px">${range}</p>
        <p style="line-height:1.7">Please find the weekly inventory report attached as an Excel spreadsheet.</p>
        <p style="line-height:1.7">Rows <strong style="color:#FF5252">highlighted in red</strong> indicate a discrepancy between provided, sold and remaining stock.</p>
        <hr style="border:none;border-top:1px solid rgba(201,168,76,.2);margin:24px 0">
        <p style="color:#5A4A38;font-size:12px">Sent automatically by Flawless Inventory Management</p>
      </div>`,
    attachments: [{ filename: `Flawless-Stock-${week}.xlsx`, content: buffer,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
  });
  await addLog({ action: 'EMAIL_SENT', week, to: 'flawlessperfumesmanagement@gmail.com' });
}

app.post('/api/send-report', async (req, res) => {
  const week = req.body.week || getWeekId();
  try {
    await sendWeeklyReport(week);
    res.json({ success: true, message: `Report for ${week} sent to flawlessperfumesmanagement@gmail.com` });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Vercel Cron endpoint (called every Sunday 9 AM by vercel.json) ────────────
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
