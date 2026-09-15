const crypto = require('crypto');

function genBookingId() {
  return `BK-${Date.now().toString(36).toUpperCase()}-${crypto.randomBytes(2).toString('hex').toUpperCase()}`;
}

// Calendar-day comparison (booking dates are date-only, stored at UTC midnight) — strips
// any time-of-day so "today" always matches a booking whose range includes it.
function toUtcMidnight(value) {
  const d = new Date(value);
  return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

/**
 * Recomputes a site's live mediaStatus/bookingInfo/isActive from its `bookings` array as of
 * `asOf` (defaults to now) and mutates the site in place. This is the single source of truth
 * for "does today fall inside an active, non-cancelled booking" — reused by every mutation
 * path (add/edit booking, status change, bulk status change, Add/Edit Site) AND by the
 * periodic reconciliation job, so a booking's start/end date alone is enough to flip status
 * automatically without any manual action.
 *
 * Manually Blocked sites are left completely untouched — a blocked state is never silently
 * overridden by booking-date reconciliation.
 *
 * Does NOT decide whether anything "changed" — callers compare their own before/after
 * snapshot (mediaStatus + bookingInfo.bookingId) since they already hold both.
 */
function resolveSiteStatus(site, asOf = new Date()) {
  if (site.mediaStatus === 'blocked') return { activeBooking: null };

  const today = toUtcMidnight(asOf);
  const bookings = site.bookings || [];

  let active = null;
  for (const b of bookings) {
    if (b.status === 'cancelled') continue;
    const start = toUtcMidnight(b.startDate);
    const end = toUtcMidnight(b.endDate);
    if (today >= start && today <= end) {
      active = b;
      break;
    }
  }

  // Keep each booking's own lifecycle status in sync with today's date.
  bookings.forEach((b) => {
    if (b.status === 'cancelled') return;
    if (active && b.bookingId === active.bookingId) {
      b.status = 'active';
      return;
    }
    const end = toUtcMidnight(b.endDate);
    const start = toUtcMidnight(b.startDate);
    b.status = today > end ? 'completed' : today < start ? 'upcoming' : b.status;
  });
  if (typeof site.markModified === 'function') site.markModified('bookings');

  if (active) {
    site.bookingInfo = {
      bookingId: active.bookingId,
      customerType: active.customerType,
      client: active.client,
      startDate: active.startDate,
      endDate: active.endDate,
      durationDays: active.durationDays,
      monthlyTotalCost: active.monthlyTotalCost,
      amount: active.amount,
      bookedBy: active.updatedBy || active.createdBy,
    };
  } else {
    site.bookingInfo = undefined;
  }
  site.mediaStatus = active ? 'booked' : 'available';
  site.isActive = true;

  return { activeBooking: active };
}

/**
 * Periodic reconciliation for sites with no incoming request today — without this, a site
 * would stay "Booked" forever after its booking's end date passes (or stay "Available"
 * after a future booking's start date arrives) until someone happens to save it again.
 * Safe to call repeatedly; only sites whose computed status actually changed are saved.
 */
async function reconcileAllSites() {
  const Site = require('../models/Site');
  const { recordStatusPeriod } = require('./inventoryTimeline');

  const candidates = await Site.find({ mediaStatus: { $ne: 'blocked' }, 'bookings.0': { $exists: true } });
  let updated = 0;
  for (const site of candidates) {
    const previousStatus = site.mediaStatus;
    const previousBookingId = site.bookingInfo?.bookingId;
    resolveSiteStatus(site);
    const changed = previousStatus !== site.mediaStatus || previousBookingId !== site.bookingInfo?.bookingId;
    if (!changed) continue;
    try {
      await site.save();
      await recordStatusPeriod({ site, previousStatus, source: 'inventory', userId: null });
      updated += 1;
    } catch (err) {
      // Don't let one bad site abort reconciliation for the rest.
      // eslint-disable-next-line no-console
      console.error(`[bookingScheduler] failed to reconcile site ${site._id}:`, err.message);
    }
  }
  return updated;
}

module.exports = { genBookingId, resolveSiteStatus, reconcileAllSites };
