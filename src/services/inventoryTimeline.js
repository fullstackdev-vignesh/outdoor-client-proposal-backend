const InventoryHistory = require('../models/InventoryHistory');
const { normalizeStatus } = require('../config/siteStatus');
const Client = require('../models/Client');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

// Centralized status-period writer, reused by /sites create/update/status-change and
// /inventory row + bulk updates so every source produces identical timeline behaviour.
// `effectiveAt` backdates a non-booked change to when it really happened (the scheduler passes 12:01 AM
// after an expired booking's end date); omitted, the change is stamped now.
async function recordStatusPeriod({ site, previousStatus, source, userId, effectiveAt }) {
  const now = effectiveAt || nowIST();
  const newStatus = site.mediaStatus;

  // Booked and dated Blocked periods run over their own dates; everything else starts now.
  const datedBlock = (newStatus === 'blocked' || newStatus === 'confirmed') && site.blockInfo?.startDate && site.blockInfo?.endDate;
  const effectiveFrom =
    newStatus === 'booked' && site.bookingInfo?.startDate ? new Date(site.bookingInfo.startDate) :
    datedBlock ? new Date(site.blockInfo.startDate) : now;
  const effectiveTo =
    newStatus === 'booked' && site.bookingInfo?.endDate ? new Date(site.bookingInfo.endDate) :
    datedBlock ? new Date(site.blockInfo.endDate) : null;

  // Close whatever period was previously open for this site (no-op for a brand-new site).
  await InventoryHistory.updateMany({ site: site._id, effectiveTo: null }, { $set: { effectiveTo: effectiveFrom, updatedAt: now } });

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
    updatedAt: now,
    changedBy: userId,
    source,
  };

  // Every booking (current, upcoming or completed) gets its own Timeline row via
  // syncBookingTimelineRecords, called unconditionally on every save — so a 'booked' row
  // here would just duplicate whichever booking happens to be live right now. Still run the
  // open-period-closing update above (e.g. closing an 'immediate' row when a booking starts).
  if (newStatus === 'booked') return;

  if (['hold', 'issue'].includes(newStatus) && site.statusInfo) {
    doc.statusSnapshot = { reason: site.statusInfo.reason, notes: site.statusInfo.notes, date: site.statusInfo.date };
  }

  if ((newStatus === 'blocked' || newStatus === 'confirmed') && site.blockInfo) {
    doc.blockSnapshot = {
      kind: site.blockInfo.kind || 'blocked',
      reason: site.blockInfo.reason,
      notes: site.blockInfo.notes,
      blockedDate: site.blockInfo.blockedDate,
      customerType: site.blockInfo.customerType,
      customerName: site.blockInfo.customerName,
      startDate: site.blockInfo.startDate,
      endDate: site.blockInfo.endDate,
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
  const bookings = site.bookings || [];
  for (const b of bookings) {
    let customerName = b.customerName;
    if (!customerName && b.client) {
      const client = await Client.findById(b.client).select('name');
      customerName = client?.name;
    }
    const isCancelled = b.status === 'cancelled';
    const update = {
      site: site._id,
      bookingId: b.bookingId,
      mediaId: site.mediaId,
      mediaType: site.mediaType,
      state: site.state,
      city: site.city,
      mediaImage: site.mediaImage,
      siteOwner: site.siteOwner,
      // Cancelling a booking never erases its own Timeline row — it just relabels it so the
      // original "Booked Period" (bookingSnapshot below) stays visible alongside why/when/by
      // whom it was cancelled.
      status: isCancelled ? 'cancelled' : 'booked',
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
    };
    if (isCancelled) {
      update.cancellationSnapshot = {
        reason: b.cancellationReason,
        cancelledAt: b.cancelledAt,
        cancelledByName: b.cancelledByName,
        cancelledByRole: b.cancelledByRole,
        cancellationType: b.cancellationType,
      };
    }
    // This runs on every save, so skip rows whose data is unchanged — otherwise every booking of
    // a site would get a fresh changedAt/updatedAt and jump to the top of the Timeline list.
    const existing = await InventoryHistory.findOne({ site: site._id, bookingId: b.bookingId }).lean();
    if (existing && !bookingRowChanged(existing, update)) continue;
    update.updatedAt = now;
    const ops = { $set: update, $setOnInsert: { bookedAt: b.createdAt || now } };
    // The row is overwritten with the latest booking values, so keep what it was before —
    // lets the Timeline show "Booking Updated: old dates → new dates" as its own step.
    if (existing && existing.status !== 'cancelled' && !isCancelled && fieldsDiffer(existing.bookingSnapshot, update.bookingSnapshot, BOOKING_EDIT_FIELDS)) {
      ops.$push = {
        edits: {
          editedAt: now,
          editedBy: userId,
          source,
          previous: pickEditFields(existing.bookingSnapshot),
          next: pickEditFields(update.bookingSnapshot),
        },
      };
    }
    await InventoryHistory.findOneAndUpdate({ site: site._id, bookingId: b.bookingId }, ops, { upsert: true });
  }
}

// Compares only the data fields of a booking Timeline row (not who/when/source metadata).
const BOOKING_ROW_FIELDS = ['mediaId', 'mediaType', 'state', 'city', 'mediaImage', 'siteOwner', 'status', 'isActive', 'effectiveFrom', 'effectiveTo'];
const BOOKING_SNAPSHOT_FIELDS = ['customerType', 'client', 'customerName', 'startDate', 'endDate', 'durationDays', 'monthlyTotalCost', 'amount'];
const CANCELLATION_SNAPSHOT_FIELDS = ['reason', 'cancelledAt', 'cancelledByName', 'cancelledByRole', 'cancellationType'];
// Changes that count as the user editing the booking (customerName is left out — older rows may
// only now be getting it filled in from the client, which isn't an edit).
const BOOKING_EDIT_FIELDS = ['customerType', 'client', 'startDate', 'endDate', 'amount'];
const EDIT_SNAPSHOT_FIELDS = ['customerType', 'client', 'customerName', 'startDate', 'endDate', 'durationDays', 'amount'];

function pickEditFields(snapshot) {
  return Object.fromEntries(EDIT_SNAPSHOT_FIELDS.map((f) => [f, snapshot?.[f]]));
}

// Dates compare by time value, ObjectIds by hex string, and empty values as equal.
function normalizeValue(v) {
  if (v === undefined || v === null || v === '') return null;
  if (v instanceof Date) return v.getTime();
  return String(v);
}

function fieldsDiffer(a, b, fields) {
  return fields.some((f) => normalizeValue(a?.[f]) !== normalizeValue(b?.[f]));
}

function bookingRowChanged(existing, update) {
  return (
    fieldsDiffer(existing, update, BOOKING_ROW_FIELDS) ||
    fieldsDiffer(existing.bookingSnapshot, update.bookingSnapshot, BOOKING_SNAPSHOT_FIELDS) ||
    Boolean(update.cancellationSnapshot && fieldsDiffer(existing.cancellationSnapshot, update.cancellationSnapshot, CANCELLATION_SNAPSHOT_FIELDS))
  );
}

// Dynamic per-booking lifecycle for Timeline display — kept separate from the stored
// 'booked' status so it always reflects today's date without needing a write. Non-booking
// (available/blocked) rows have no lifecycle.
function computeBookingLifecycle(item) {
  if (item.status !== 'booked') return null;
  // Compare IST calendar days (booking dates are date-only at UTC midnight), so the end date itself
  // still counts as active and the switch happens at IST midnight.
  const day = (v) => {
    const d = new Date(v);
    return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  };
  const today = day(nowIST());
  const start = item.effectiveFrom ? day(item.effectiveFrom) : null;
  const end = item.effectiveTo ? day(item.effectiveTo) : null;
  if (start != null && today < start) return 'upcoming';
  if (end != null && today > end) return 'completed';
  return 'active';
}

function escapeRegex(text) {
  return text.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
}

// Period-overlap filter: a history record is included if its period intersects [from, to].
// An open-ended record (effectiveTo === null) is treated as ongoing until now.
function buildOverlapFilter(query) {
  const filter = {};
  if (query.state) filter.state = new RegExp(`^${escapeRegex(query.state.trim())}$`, 'i');
  if (query.city) filter.city = new RegExp(escapeRegex(query.city.trim()), 'i');
  if (query.siteOwner) filter.siteOwner = query.siteOwner;
  if (query.mediaStatus) filter.status = normalizeStatus(query.mediaStatus);
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
