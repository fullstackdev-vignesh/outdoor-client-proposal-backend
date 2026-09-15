require('dotenv').config();
const mongoose = require('mongoose');
const Site = require('../models/Site');
const InventoryHistory = require('../models/InventoryHistory');
const Client = require('../models/Client');

async function run() {
  await mongoose.connect(process.env.MONGO_URI);

  const sites = await Site.find({});
  let created = 0;

  for (const site of sites) {
    const existing = await InventoryHistory.countDocuments({ site: site._id });
    if (existing > 0) continue; // never overwrite/duplicate if already backfilled or has real history

    const effectiveFrom = site.mediaStatus === 'booked' && site.bookingInfo?.startDate ? new Date(site.bookingInfo.startDate) : site.createdAt;
    const effectiveTo = site.mediaStatus === 'booked' && site.bookingInfo?.endDate ? new Date(site.bookingInfo.endDate) : null;

    const doc = {
      site: site._id,
      mediaId: site.mediaId,
      mediaType: site.mediaType,
      state: site.state,
      city: site.city,
      mediaImage: site.mediaImage,
      status: site.mediaStatus,
      previousStatus: null,
      isActive: site.isActive,
      effectiveFrom,
      effectiveTo,
      changedAt: site.updatedAt || site.createdAt,
      changedBy: site.createdBy,
      source: 'sites',
    };

    if (site.mediaStatus === 'booked' && site.bookingInfo) {
      let customerName;
      if (site.bookingInfo.client) {
        const client = await Client.findById(site.bookingInfo.client).select('name');
        customerName = client?.name;
      }
      doc.bookingSnapshot = {
        customerType: site.bookingInfo.customerType,
        client: site.bookingInfo.client,
        customerName,
        startDate: site.bookingInfo.startDate,
        endDate: site.bookingInfo.endDate,
        durationDays: site.bookingInfo.durationDays,
        monthlyTotalCost: site.bookingInfo.monthlyTotalCost,
        amount: site.bookingInfo.amount,
      };
    } else if (site.mediaStatus === 'blocked' && site.blockInfo) {
      doc.blockSnapshot = {
        reason: site.blockInfo.reason,
        notes: site.blockInfo.notes,
        blockedDate: site.blockInfo.blockedDate,
      };
    }

    await InventoryHistory.create(doc);
    created += 1;
  }

  console.log(`Backfilled an initial timeline period for ${created} site(s) (skipped sites that already had history).`);
  await mongoose.disconnect();
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
