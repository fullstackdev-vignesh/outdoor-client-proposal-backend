const InventoryHistory = require('../models/InventoryHistory');
const Client = require('../models/Client');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

// Centralized status-period writer, reused by /sites create/update/status-change and
// /inventory row + bulk updates so every source produces identical timeline behaviour.
async function recordStatusPeriod({ site, previousStatus, source, userId }) {
  const now = nowIST();
  const newStatus = site.mediaStatus;

  const effectiveFrom = newStatus === 'booked' && site.bookingInfo?.startDate ? new Date(site.bookingInfo.startDate) : now;
  const effectiveTo = newStatus === 'booked' && site.bookingInfo?.endDate ? new Date(site.bookingInfo.endDate) : null;

  // Close whatever period was previously open for this site (no-op for a brand-new site).
  await InventoryHistory.updateMany({ site: site._id, effectiveTo: null }, { $set: { effectiveTo: effectiveFrom } });

  const doc = {
    site: site._id,
    mediaId: site.mediaId,
    mediaType: site.mediaType,
    state: site.state,
    city: site.city,
    image: site.image,
    status: newStatus,
    previousStatus: previousStatus || null,
    isActive: site.isActive,
    effectiveFrom,
    effectiveTo,
    changedAt: now,
    changedBy: userId,
    source,
  };

  if (newStatus === 'booked' && site.bookingInfo) {
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
  } else if (newStatus === 'blocked' && site.blockInfo) {
    doc.blockSnapshot = {
      reason: site.blockInfo.reason,
      notes: site.blockInfo.notes,
      blockedDate: site.blockInfo.blockedDate,
    };
  }

  await InventoryHistory.create(doc);
}

// Period-overlap filter: a history record is included if its period intersects [from, to].
// An open-ended record (effectiveTo === null) is treated as ongoing until now.
function buildOverlapFilter(query) {
  const filter = {};
  if (query.state) filter.state = query.state;
  if (query.city) filter.city = query.city;
  if (query.mediaStatus) filter.status = query.mediaStatus;
  if (query.isActive !== undefined && query.isActive !== '') filter.isActive = query.isActive === 'true';
  if (query.search) {
    filter.$or = [{ mediaId: new RegExp(query.search, 'i') }, { mediaType: new RegExp(query.search, 'i') }, { city: new RegExp(query.search, 'i') }, { state: new RegExp(query.search, 'i') }];
  }

  if (query.from || query.to) {
    const from = query.from ? new Date(query.from) : null;
    const to = query.to ? new Date(query.to) : null;
    if (to) filter.effectiveFrom = { $lte: to };
    if (from) {
      filter.$and = (filter.$and || []).concat([{ $or: [{ effectiveTo: null }, { effectiveTo: { $gte: from } }] }]);
    }
  }

  return filter;
}

module.exports = { recordStatusPeriod, buildOverlapFilter, nowIST };
