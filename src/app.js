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
const siteInfoRoutes = require('./routes/siteInfoRoutes');
const { pptRouter, excelRouter } = require('./routes/templateRoutes');

const app = express();

// Browsers send Origin as scheme://host[:port] only — no path, no trailing slash.
const allowedOrigins = [
  'http://localhost:3000',
  'http://localhost:5000',
  'http://localhost:8080',
  'https://outdoor-client-proposal-frontend.vercel.app',
  'https://outdoor-client-proposal-frontend-m3.vercel.app',
  'https://adinn-space.sgp1.cdn.digitaloceanspaces.com',
  'https://adinntech.in',
  'https://www.adinntech.in',
  ...(process.env.CLIENT_URL || '').split(',').map((o) => o.trim().replace(/\/+$/, '')).filter(Boolean),
];

const corsOptions = {
  origin(origin, callback) {
    // Allow non-browser requests (Postman, server-to-server) which send no Origin
    if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
    console.warn('Blocked by CORS:', origin);
    return callback(null, false);
  },
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
};

app.use(cors(corsOptions));
app.options('*', cors(corsOptions));
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

app.get('/api/health', (req, res) => res.json({ status: 'ok' }));
app.use('/api/image-proxy', require('./routes/imageProxyRoutes'));

app.use('/generated', express.static(path.join(__dirname, '..', 'generated')));
app.use('/uploads', express.static(path.join(__dirname, '..', 'uploads')));
app.use('/reference', express.static(path.join(__dirname, '..', 'reference')));

app.use('/admin', authRoutes);
app.use('/api/auth', authRoutes);
app.use('/api/users', userRoutes);
app.use('/api/sites', siteRoutes);
app.use('/api/clients', clientRoutes);
app.use('/api/bookings', bookingRoutes);
app.use('/api/proposals', proposalRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/locations', locationRoutes);
app.use('/api/site-info', siteInfoRoutes);
app.use('/api/ppt-templates', pptRouter);
app.use('/api/excel-templates', excelRouter);

app.use(notFound);
app.use(errorHandler);

module.exports = app;
