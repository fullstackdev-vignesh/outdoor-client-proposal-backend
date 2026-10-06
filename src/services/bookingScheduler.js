const crypto = require('crypto');
const { OVERRIDE_STATUSES, DATED_STATUSES } = require('../config/siteStatus');

function genBookingId() {
  return `BK-${Date.now().toString(36).toUpperCase()}-${crypto.randomBytes(2).toString('hex').toUpperCase()}`;
}

// Calendar-day comparison (booking dates are date-only, stored at UTC midnight) — strips
// any time-of-day so "today" always matches a booking whose range includes it.
function toUtcMidnight(value) {
  const d = new Date(value);
  return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

const IST_OFFSET_MS = 330 * 60000;
const DAY_MS = 24 * 60 * 60 * 1000;
// IST wall-clock stored as UTC (same convention as every model's nowIST), so its UTC date is the IST date —
// booking days therefore roll over at IST midnight, not at 05:30 IST (UTC midnight).
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

// The moment a booking (or block) stops being live: 12:01 AM IST on the day after its end date. Used to
// stamp the change at the real transition time, even when the server was down then and the reconcile
// only ran later (e.g. on startup).
function bookingEndTransitionAt(endDate) {
  return new Date(toUtcMidnight(endDate) + DAY_MS + 60 * 1000);
}

// The moment a scheduled block starts: 12:01 AM IST on its Start Date.
function blockStartTransitionAt(startDate) {
  return new Date(toUtcMidnight(startDate) + 60 * 1000);
}

// A Blocked/Confirmed period with its own Start/End Date (older blocks have neither and stay until changed by hand).
const isDatedBlock = (block) => !!(block && block.startDate && block.endDate);

/**
 * Recomputes a site's live mediaStatus/bookingInfo from its `bookings` array and its dated block
 * (blockInfo.startDate..endDate) as of `asOf` (defaults to now) and mutates the site in place.
 * This is the single source of truth for "does today fall inside an active booking / block" —
 * reused by every mutation path (add/edit booking, status change, bulk status change, Add/Edit
 * Site) AND by the periodic reconciliation job, so the dates alone flip the status automatically.
 *
 *   • Hold / Issue are set by hand and left completely untouched.
 *   • An older block without dates is also left untouched (it stays Blocked until changed by hand).
 *   • A dated Blocked/Confirmed period gives the site that status from its Start Date through its
 *     End Date; once the End Date has passed it is cleared. One scheduled ahead leaves the site as it
 *     is until then. Its dates and booking dates never overlap (checked when either is saved).
 *
 *   • An Inactive site is always Immediate: any Blocked/Confirmed period and Hold/Issue reason is
 *     dropped. Its bookings are kept (their own lifecycle still moves with the dates) but don't make
 *     it Booked; once it's Active again they count as usual.
 *
 * Active/Inactive is the site's own setting and is never changed here. Does NOT decide whether
 * anything "changed" — callers compare their own before/after snapshot.
 */
function resolveSiteStatus(site, asOf = nowIST()) {
  if (site.isActive === false) {
    site.blockInfo = undefined;
    site.statusInfo = undefined;
    site.mediaStatus = 'immediate'; // fall through: bookings' own lifecycle still updates below
  }
  if (OVERRIDE_STATUSES.includes(site.mediaStatus)) return { activeBooking: null };
  const block = site.blockInfo;
  if (site.mediaStatus === 'blocked' && block && !isDatedBlock(block)) return { activeBooking: null };

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

  let blockActive = false;
  if (isDatedBlock(block)) {
    if (today > toUtcMidnight(block.endDate)) site.blockInfo = undefined; // block finished
    // Blocked and Confirmed both show straight away (from when they're saved) through their End Date —
    // except while a booking is running, which stays Booked until it ends.
    else blockActive = !active;
  }

  // Inactive: bookings never make the site Booked.
  if (site.isActive === false) active = null;

  if (active && !blockActive) {
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
  site.mediaStatus = blockActive ? block.kind || 'blocked' : active ? 'booked' : 'immediate';

  return { activeBooking: blockActive ? null : active };
}

/**
 * Periodic reconciliation for sites with no incoming request today — without this, a site
 * would stay "Booked"/"Blocked" forever after its booking's / block's end date passes (or stay
 * "Immediate" after a future one's start date arrives) until someone happens to save it again.
 * Safe to call repeatedly; only sites whose computed status actually changed are saved.
 */
async function reconcileAllSites() {
  const Site = require('../models/Site');
  const { recordStatusPeriod } = require('./inventoryTimeline');

  const candidates = await Site.find({
    mediaStatus: { $nin: OVERRIDE_STATUSES },
    $or: [{ 'bookings.0': { $exists: true } }, { 'blockInfo.startDate': { $exists: true } }],
  });
  let updated = 0;
  for (const site of candidates) {
    const previousStatus = site.mediaStatus;
    const previousBookingId = site.bookingInfo?.bookingId;
    const previousEndDate = DATED_STATUSES.includes(previousStatus) ? site.blockInfo?.endDate : site.bookingInfo?.endDate;
    const pendingBlockStart = site.blockInfo?.startDate;
    resolveSiteStatus(site);
    const changed = previousStatus !== site.mediaStatus || previousBookingId !== site.bookingInfo?.bookingId;
    if (!changed) continue;
    // Stamp the change when it really happened, not whenever this ran: a booking/block that ran out
    // ended at 12:01 AM the day after its End Date; a scheduled block started at 12:01 AM on its Start Date.
    const at =
      DATED_STATUSES.includes(site.mediaStatus) && pendingBlockStart ? blockStartTransitionAt(pendingBlockStart) :
      (previousStatus === 'booked' || DATED_STATUSES.includes(previousStatus)) && site.mediaStatus === 'immediate' && previousEndDate
        ? bookingEndTransitionAt(previousEndDate)
        : null;
    const effectiveAt = at ? new Date(Math.min(at.getTime(), nowIST().getTime())) : undefined;
    try {
      await site.save();
      await recordStatusPeriod({ site, previousStatus, source: 'inventory', userId: null, effectiveAt });
      updated += 1;
    } catch (err) {
      // Don't let one bad site abort reconciliation for the rest.
      // eslint-disable-next-line no-console
      console.error(`[bookingScheduler] failed to reconcile site ${site._id}:`, err.message);
    }
  }
  return updated;
}

module.exports = { genBookingId, resolveSiteStatus, reconcileAllSites, isDatedBlock };
