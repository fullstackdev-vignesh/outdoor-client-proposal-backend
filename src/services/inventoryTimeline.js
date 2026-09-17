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
    mediaImage: site.mediaImage,
    siteOwner: site.siteOwner,
    status: newStatus,
    previousStatus: previousStatus || null,
    isActive: site.isActive,
    effectiveFrom,
    effectiveTo,
    changedAt: now,
    changedBy: userId,
    source,
  };

  // Every booking (current, upcoming or completed) gets its own Timeline row via
  // syncBookingTimelineRecords, called unconditionally on every save — so a 'booked' row
  // here would just duplicate whichever booking happens to be live right now. Still run the
  // open-period-closing update above (e.g. closing an 'available' row when a booking starts).
  if (newStatus === 'booked') return;

  if (newStatus === 'blocked' && site.blockInfo) {
    doc.blockSnapshot = {
      reason: site.blockInfo.reason,
      notes: site.blockInfo.notes,
      blockedDate: site.blockInfo.blockedDate,
    };
  }

  await InventoryHistory.create(doc);
}

// Upserts ONE Timeline row per non-cancelled booking in `site.bookings`, keyed by
// (site, bookingId) — independent of whichever booking is currently "live". This is what
// makes an Upcoming booking (e.g. starting next month) visible in Inventory Timeline the
// moment it's saved, without waiting for its Start Date, and without disturbing the site's
// live mediaStatus (computed separately by bookingScheduler.resolveSiteStatus). Called
// unconditionally after every save that touches bookings, so editing a booking's dates just
// updates its existing row in place instead of creating a new one.
async function syncBookingTimelineRecords(site, userId, source) {
  const now = nowIST();
  const bookings = (site.bookings || []).filter((b) => b.status !== 'cancelled');
  for (const b of bookings) {
    let customerName = b.customerName;
    if (!customerName && b.client) {
      const client = await Client.findById(b.client).select('name');
      customerName = client?.name;
    }
    await InventoryHistory.findOneAndUpdate(
      { site: site._id, bookingId: b.bookingId },
      {
        $set: {
          site: site._id,
          bookingId: b.bookingId,
          mediaId: site.mediaId,
          mediaType: site.mediaType,
          state: site.state,
          city: site.city,
          mediaImage: site.mediaImage,
          siteOwner: site.siteOwner,
          status: 'booked',
          isActive: site.isActive,
          effectiveFrom: new Date(b.startDate),
          effectiveTo: new Date(b.endDate),
          changedAt: now,
          changedBy: userId,
          source,
          bookingSnapshot: {
            customerType: b.customerType,
            client: b.client,
            customerName,
            startDate: b.startDate,
            endDate: b.endDate,
            durationDays: b.durationDays,
            monthlyTotalCost: b.monthlyTotalCost,
            amount: b.amount,
          },
        },
      },
      { upsert: true }
    );
  }
}

// Dynamic per-booking lifecycle for Timeline display — kept separate from the stored
// 'booked' status so it always reflects today's date without needing a write. Non-booking
// (available/blocked) rows have no lifecycle.
function computeBookingLifecycle(item) {
  if (item.status !== 'booked') return null;
  const now = Date.now();
  const start = item.effectiveFrom ? new Date(item.effectiveFrom).getTime() : null;
  const end = item.effectiveTo ? new Date(item.effectiveTo).getTime() : null;
  if (start != null && now < start) return 'upcoming';
  if (end != null && now > end) return 'completed';
  return 'active';
}

// Period-overlap filter: a history record is included if its period intersects [from, to].
// An open-ended record (effectiveTo === null) is treated as ongoing until now.
function buildOverlapFilter(query) {
  const filter = {};
  if (query.state) filter.state = query.state;
  if (query.city) filter.city = query.city;
  if (query.siteOwner) filter.siteOwner = query.siteOwner;
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

module.exports = { recordStatusPeriod, buildOverlapFilter, nowIST, syncBookingTimelineRecords, computeBookingLifecycle };
