require('dotenv').config();
const app = require('./app');
const connectDB = require('./config/db');
const { reconcileAllSites } = require('./services/bookingScheduler');

const PORT = process.env.PORT || 5000;
const RECONCILE_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes

connectDB()
  .then(() => {
    app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

    // Automatic booking status reconciliation — without this, a site's live status would
    // only flip Booked<->Available on a booking's start/end date if someone happened to
    // save it that day. No existing scheduler/cron was present in this app to reuse.
    reconcileAllSites().catch((err) => console.error('[bookingScheduler] initial reconcile failed:', err.message));
    setInterval(() => {
      reconcileAllSites().catch((err) => console.error('[bookingScheduler] reconcile failed:', err.message));
    }, RECONCILE_INTERVAL_MS);
  })
  .catch((err) => {
    console.error('Failed to connect to MongoDB', err);
    process.exit(1);
  });


  