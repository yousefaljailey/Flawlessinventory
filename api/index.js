// Vercel serverless entry point — exports the Express app
require('dotenv').config();
const { app } = require('../app');
module.exports = app;
