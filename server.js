require('dotenv').config();
const express = require('express');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const cron = require('node-cron');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const DATA_DIR = path.join(__dirname, 'data');
const SHOPS_FILE = path.join(DATA_DIR, 'shops.json');
const STOCK_FILE = path.join(DATA_DIR, 'stock.json');
const LOGS_FILE  = path.join(DATA_DIR, 'logs.json');

const PRODUCTS = [
  'Monte-Carlo 30ml',
  'Monte-Carlo 100ml',
  'Summer in Paris 30ml',
  'Bedouin 100ml',
  'Rub Al Khali 30ml',
  'Discovery Set'
];

// ── Data helpers ──────────────────────────────────────────────────────────────
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const INIT_VALUES = { [SHOPS_FILE]: [], [STOCK_FILE]: {}, [LOGS_FILE]: [] };
Object.entries(INIT_VALUES).forEach(([f, def]) => {
  if (!fs.existsSync(f)) fs.writeFileSync(f, JSON.stringify(def));
  else {
    // Ensure logs and shops are always arrays, stock is always an object
    try {
      const parsed = JSON.parse(fs.readFileSync(f, 'utf-8'));
      const shouldBeArray = f !== STOCK_FILE;
      if (shouldBeArray && !Array.isArray(parsed)) fs.writeFileSync(f, JSON.stringify(def));
      if (!shouldBeArray && (Array.isArray(parsed) || typeof parsed !== 'object')) fs.writeFileSync(f, JSON.stringify(def));
    } catch { fs.writeFileSync(f, JSON.stringify(def)); }
  }
});

const read  = f => JSON.parse(fs.readFileSync(f, 'utf-8'));
const write = (f, d) => fs.writeFileSync(f, JSON.stringify(d, null, 2));

function addLog(entry) {
  const logs = read(LOGS_FILE);
  logs.unshift({ ...entry, timestamp: new Date().toISOString() });
  write(LOGS_FILE, logs.slice(0, 2000));
}

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
  const w1start  = new Date(yr, 0, 4);
  w1start.setDate(w1start.getDate() - ((w1start.getDay() + 6) % 7));
  const start = new Date(w1start);
  start.setDate(start.getDate() + (wn - 1) * 7);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const fmt = d => d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  return `${fmt(start)} – ${fmt(end)}`;
}

// ── API: Shops ────────────────────────────────────────────────────────────────
app.get('/api/shops', (_, res) => res.json(read(SHOPS_FILE)));

app.post('/api/shops', (req, res) => {
  const name = req.body.name?.trim();
  if (!name) return res.status(400).json({ error: 'Shop name required' });
  const shops = read(SHOPS_FILE);
  if (shops.includes(name)) return res.status(400).json({ error: 'Shop already exists' });
  shops.push(name);
  write(SHOPS_FILE, shops);
  addLog({ action: 'ADD_SHOP', shop: name });
  res.json({ success: true });
});

app.delete('/api/shops/:name', (req, res) => {
  const name = decodeURIComponent(req.params.name);
  write(SHOPS_FILE, read(SHOPS_FILE).filter(s => s !== name));
  addLog({ action: 'REMOVE_SHOP', shop: name });
  res.json({ success: true });
});

// ── API: Products ─────────────────────────────────────────────────────────────
app.get('/api/products', (_, res) => res.json(PRODUCTS));

// ── API: Stock ────────────────────────────────────────────────────────────────
app.get('/api/stock/:week', (req, res) => {
  const stock = read(STOCK_FILE);
  res.json(stock[req.params.week] || {});
});

app.post('/api/stock/bulk', (req, res) => {
  const { shop, week, products } = req.body;
  if (!shop || !week || !products) return res.status(400).json({ error: 'Missing fields' });

  const stock = read(STOCK_FILE);
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
  write(STOCK_FILE, stock);
  addLog({ action: 'STOCK_SAVE', shop, week, products: Object.keys(products).length });
  res.json({ success: true });
});

// ── API: Logs ─────────────────────────────────────────────────────────────────
app.get('/api/logs', (req, res) => {
  const limit = Math.min(Number(req.query.limit) || 200, 2000);
  res.json(read(LOGS_FILE).slice(0, limit));
});

// ── API: Week info ────────────────────────────────────────────────────────────
app.get('/api/week', (_, res) => {
  const week = getWeekId();
  res.json({ week, range: weekRange(week) });
});

// ── Report generation ─────────────────────────────────────────────────────────
async function buildWorkbook(week) {
  const stock    = read(STOCK_FILE);
  const shops    = read(SHOPS_FILE);
  const weekData = stock[week] || {};

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Flawless Inventory';
  wb.created = new Date();

  const ws = wb.addWorksheet(`Week ${week}`, {
    pageSetup: { fitToPage: true, fitToWidth: 1 }
  });

  // Title row
  ws.mergeCells('A1:G1');
  const titleCell = ws.getCell('A1');
  titleCell.value = `FLAWLESS PERFUMES — Stock Report ${week}  (${weekRange(week)})`;
  titleCell.font  = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFC9A84C' } };
  titleCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A0806' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;

  // Header row
  const headers = ['Shop', 'Product', 'Provided', 'Sold', 'Remaining', 'Expected', 'Status'];
  const widths  = [28, 26, 12, 12, 12, 12, 14];
  ws.columns = headers.map((h, i) => ({ header: h, key: h.toLowerCase(), width: widths[i] }));
  const headerRow = ws.getRow(2);
  headerRow.values = headers;
  headerRow.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1510' } };
    cell.font      = { name: 'Calibri', bold: true, color: { argb: 'FFC9A84C' }, size: 11 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border    = { bottom: { style: 'thin', color: { argb: 'FFC9A84C' } } };
  });
  ws.getRow(2).height = 22;

  let rowIdx = 3;
  let hasMismatches = false;

  shops.forEach(shop => {
    PRODUCTS.forEach(product => {
      const e        = weekData[shop]?.[product] || { provided: 0, sold: 0, remaining: 0 };
      const provided  = e.provided  || 0;
      const sold      = e.sold      || 0;
      const remaining = e.remaining || 0;
      const expected  = provided - sold;
      const hasData   = provided > 0 || sold > 0 || remaining > 0;
      const isMismatch = hasData && remaining !== expected;
      if (isMismatch) hasMismatches = true;

      const status = !hasData ? 'Not entered' : isMismatch ? '⚠ MISMATCH' : '✓ OK';
      const row = ws.getRow(rowIdx);
      row.values = [shop, product, provided, sold, remaining, expected, status];

      row.eachCell((cell, colNum) => {
        cell.font      = { name: 'Calibri', size: 10, color: { argb: isMismatch ? 'FFFFFFFF' : 'FFF0E6D0' } };
        cell.alignment = { horizontal: colNum <= 2 ? 'left' : 'center', vertical: 'middle' };
        if (isMismatch) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC2222' } };
          cell.font = { ...cell.font, bold: true };
        } else if (rowIdx % 2 === 0) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF100D0A' } };
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A0806' } };
        }
      });
      row.getCell(7).font = {
        ...row.getCell(7).font,
        color: { argb: isMismatch ? 'FFFFFFFF' : (!hasData ? 'FF888888' : 'FF4CAF8A') }
      };
      row.height = 18;
      rowIdx++;
    });
  });

  // Summary footer
  rowIdx++;
  ws.mergeCells(`A${rowIdx}:G${rowIdx}`);
  const footerCell = ws.getCell(`A${rowIdx}`);
  footerCell.value = hasMismatches
    ? `⚠  This report contains discrepancies. Rows highlighted in RED indicate provided − sold ≠ remaining.`
    : `✓  All stock counts balance correctly for week ${week}.`;
  footerCell.font  = { name: 'Calibri', italic: true, color: { argb: hasMismatches ? 'FFFF5252' : 'FF4CAF8A' }, size: 10 };
  footerCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1510' } };
  footerCell.alignment = { horizontal: 'center' };

  return wb;
}

// ── API: Download report ──────────────────────────────────────────────────────
app.get('/api/report/:week', async (req, res) => {
  try {
    const wb = await buildWorkbook(req.params.week);
    res.setHeader('Content-Disposition', `attachment; filename="Flawless-Stock-${req.params.week}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Email ─────────────────────────────────────────────────────────────────────
async function sendWeeklyReport(week) {
  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
    throw new Error('Email credentials not set. Copy .env.example to .env and fill in EMAIL_USER and EMAIL_PASS.');
  }

  const wb      = await buildWorkbook(week);
  const buffer  = await wb.xlsx.writeBuffer();
  const range   = weekRange(week);

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS }
  });

  await transporter.sendMail({
    from: `"Flawless Inventory" <${process.env.EMAIL_USER}>`,
    to:   'flawlessperfumesmanagement@gmail.com',
    subject: `Flawless Perfumes — Weekly Stock Report ${week}`,
    html: `
      <div style="font-family:sans-serif;max-width:600px;background:#0A0806;color:#F0E6D0;padding:32px;border-radius:10px">
        <h1 style="font-family:Georgia,serif;color:#C9A84C;letter-spacing:0.15em;font-size:24px;margin:0 0 4px">FLAWLESS</h1>
        <p style="color:#8A7860;margin:0 0 24px;font-size:12px;letter-spacing:0.1em">PERFUMES</p>
        <h2 style="font-size:18px;font-weight:600;color:#F0E6D0;margin:0 0 8px">Weekly Stock Report</h2>
        <p style="color:#8A7860;margin:0 0 24px">${week} &nbsp;·&nbsp; ${range}</p>
        <p style="line-height:1.7">Please find the weekly inventory report attached as an Excel spreadsheet.</p>
        <p style="line-height:1.7">Any rows <strong style="color:#FF5252">highlighted in red</strong> indicate a discrepancy — the remaining stock does not equal provided minus sold.</p>
        <hr style="border:none;border-top:1px solid rgba(201,168,76,0.2);margin:24px 0">
        <p style="color:#5A4A38;font-size:12px">Generated automatically by Flawless Inventory Management System</p>
      </div>
    `,
    attachments: [{
      filename:    `Flawless-Stock-${week}.xlsx`,
      content:     buffer,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }]
  });

  addLog({ action: 'EMAIL_SENT', week, to: 'flawlessperfumesmanagement@gmail.com' });
}

app.post('/api/send-report', async (req, res) => {
  const week = req.body.week || getWeekId();
  try {
    await sendWeeklyReport(week);
    res.json({ success: true, message: `Report for ${week} sent to flawlessperfumesmanagement@gmail.com` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Cron: every Sunday at 9:00 AM ────────────────────────────────────────────
cron.schedule('0 9 * * 0', async () => {
  const week = getWeekId();
  console.log(`[CRON ${new Date().toISOString()}] Sending weekly report for ${week}…`);
  try {
    await sendWeeklyReport(week);
    console.log('[CRON] Report sent successfully.');
  } catch (err) {
    console.error('[CRON] Failed to send report:', err.message);
  }
});

// ── Start ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n  Flawless Inventory  →  http://localhost:${PORT}`);
  console.log(`  Weekly report cron  →  every Sunday at 9:00 AM\n`);
});
