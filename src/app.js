const path = require('path');
const express = require('express');
const cors = require('cors');
const { notFound, errorHandler } = require('./middleware/errorHandler');

const authRoutes = require('./routes/authRoutes');
const userRoutes = require('./routes/userRoutes');
const siteRoutes = require('./routes/siteRoutes');
const clientRoutes = require('./routes/clientRoutes');
const bookingRoutes = require('./routes/bookingRoutes');
const proposalRoutes = require('./routes/proposalRoutes');
const dashboardRoutes = require('./routes/dashboardRoutes');
const locationRoutes = require('./routes/locationRoutes');
const { pptRouter, excelRouter } = require('./routes/templateRoutes');

const app = express();

app.use(cors({ origin: process.env.CLIENT_URL || 'http://localhost:3000', credentials: true }));
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

// Generated proposal files are stored under a randomized, collision-safe physical filename
// (see storageService.js) — a `?download=` query param lets a link force the browser's Save
// dialog to use the real human-readable name instead, straight from the server, so it works
// regardless of which frontend code path served the link (plain <a>, fetch+blob, cached JS).
function withDownloadName(req, res, next) {
  const name = req.query.download;
  if (name) {
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(String(name))}"`);
  }
  next();
}

app.use('/generated', withDownloadName, express.static(path.join(__dirname, '..', 'generated')));
app.use('/uploads', withDownloadName, express.static(path.join(__dirname, '..', 'uploads')));

app.use('/admin', authRoutes);
app.use('/api/auth', authRoutes);
app.use('/api/users', userRoutes);
app.use('/api/sites', siteRoutes);
app.use('/api/clients', clientRoutes);
app.use('/api/bookings', bookingRoutes);
app.use('/api/proposals', proposalRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/locations', locationRoutes);
app.use('/api/ppt-templates', pptRouter);
app.use('/api/excel-templates', excelRouter);

app.use(notFound);
app.use(errorHandler);

module.exports = app;
