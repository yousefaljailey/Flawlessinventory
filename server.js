// Local development server — NOT used by Vercel
require('dotenv').config();
const cron = require('node-cron');
const { app, getWeekId, sendWeeklyReport } = require('./app');

// Weekly cron: every Sunday at 9:00 AM (local only; Vercel uses /api/cron)
cron.schedule('0 9 * * 0', async () => {
  const week = getWeekId();
  console.log(`[CRON ${new Date().toISOString()}] Sending weekly report for ${week}…`);
  try {
    await sendWeeklyReport(week);
    console.log('[CRON] Report sent successfully.');
  } catch (err) {
    console.error('[CRON] Failed:', err.message);
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n  Flawless Inventory  →  http://localhost:${PORT}`);
  console.log(`  Weekly report cron  →  every Sunday at 9:00 AM\n`);
});
