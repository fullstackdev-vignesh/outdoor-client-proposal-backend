// One-time data update for the Immediate / Confirmed / Hold / Issue statuses. Safe to run any number
// of times — once applied, every query below matches nothing. Runs automatically on server start
// (src/server.js), or by hand: `npm run migrate:statuses`.
//   1. Site.mediaStatus 'available' -> 'immediate'
//   2. Blocked sites are no longer forced Inactive (Active/Inactive is the site's own setting now),
//      so blocked sites that were made Inactive only by the old block logic become Active again.
//   3. Inventory Timeline rows: status / previousStatus 'available' -> 'immediate'
//   4. Confirmed sites from the old .env-duration version (no dated record) become Immediate
//      (booking reconciliation then moves any with a running booking to Booked). Today's Confirmed,
//      which has its own Start/End Date like Blocked, is never touched.
// Uses the raw collections, so no model hooks bump timestamps or "updated by".
const Site = require('../models/Site');
const InventoryHistory = require('../models/InventoryHistory');

async function migrateSiteStatuses() {
  const renamed = await Site.collection.updateMany({ mediaStatus: 'available' }, { $set: { mediaStatus: 'immediate' } });
  const reactivated = await Site.collection.updateMany({ mediaStatus: 'blocked', isActive: false }, { $set: { isActive: true } });
  const history = await InventoryHistory.collection.updateMany({ status: 'available' }, { $set: { status: 'immediate' } });
  const historyPrev = await InventoryHistory.collection.updateMany({ previousStatus: 'available' }, { $set: { previousStatus: 'immediate' } });
  const unconfirmed = await Site.collection.updateMany(
    { mediaStatus: 'confirmed', 'blockInfo.startDate': { $exists: false } },
    { $set: { mediaStatus: 'immediate' }, $unset: { statusInfo: '' } }
  );
  // 5. A site with no isActive at all is Active (the schema default) — set it, so the list sort
  //    (Active first, Inactive last) puts it with the Active sites.
  const activeDefault = await Site.collection.updateMany({ isActive: { $exists: false } }, { $set: { isActive: true } });
  const changed = activeDefault.modifiedCount + renamed.modifiedCount + reactivated.modifiedCount + history.modifiedCount + historyPrev.modifiedCount + unconfirmed.modifiedCount;
  if (changed) {
    // eslint-disable-next-line no-console
    console.log(
      `[migrateSiteStatuses] sites available->immediate: ${renamed.modifiedCount}, blocked sites re-activated: ${reactivated.modifiedCount}, ` +
        `timeline rows: ${history.modifiedCount} (+${historyPrev.modifiedCount} previous-status), confirmed->immediate: ${unconfirmed.modifiedCount}`
    );
  }
  return changed;
}

module.exports = { migrateSiteStatuses };

if (require.main === module) {
  require('dotenv').config();
  const connectDB = require('../config/db');
  const mongoose = require('mongoose');
  connectDB()
    .then(migrateSiteStatuses)
    .then((n) => console.log(`Done — ${n} document(s) updated.`))
    .catch((err) => {
      console.error(err);
      process.exitCode = 1;
    })
    .finally(() => mongoose.disconnect());
}
