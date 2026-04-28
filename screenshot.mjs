#!/usr/bin/env node
import puppeteer from 'puppeteer';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

(async () => {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });
  await page.goto('http://localhost:3001', { waitUntil: 'networkidle2', timeout: 10000 });

  const dir = path.join(__dirname, 'temporary screenshots');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  const file = path.join(dir, 'login-page.png');
  await page.screenshot({ path: file, fullPage: true });
  console.log(`Screenshot saved: ${file}`);

  await browser.close();
})().catch(e => { console.error('Screenshot error:', e.message); process.exit(1); });
