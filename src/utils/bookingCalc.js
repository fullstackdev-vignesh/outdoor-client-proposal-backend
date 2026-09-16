const MS_PER_DAY = 24 * 60 * 60 * 1000;
const IST_OFFSET_MS = 5.5 * 60 * 60 * 1000;
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// Plain 'YYYY-MM-DD' for "today" in IST calendar terms — matches how the rest of the
// backend stamps dates (nowIST()) so a straight string comparison against a booking's
// plain YYYY-MM-DD startDate never drifts across a UTC/local boundary.
function todayDateOnly() {
  return new Date(Date.now() + IST_OFFSET_MS).toISOString().slice(0, 10);
}

function calcDurationDays(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const diff = Math.round((end - start) / MS_PER_DAY);
  return diff >= 0 ? diff + 1 : 0;
}

function calcBookingAmount(monthlyTotalCost, durationDays) {
  const cost = Number(monthlyTotalCost) || 0;
  const days = Number(durationDays) || 0;
  return Math.round(((cost / 30) * days + Number.EPSILON) * 100) / 100;
}

// Date-only label for friendly messages, e.g. "17-Sep-2026". Booking dates are plain
// calendar dates (no time-of-day meaning), so UTC getters avoid any timezone drift.
function formatDateLabel(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return '';
  const day = String(d.getUTCDate()).padStart(2, '0');
  const month = MONTHS[d.getUTCMonth()];
  return `${day}-${month}-${d.getUTCFullYear()}`;
}

// Standard overlap rule: two ranges [aStart,aEnd] and [bStart,bEnd] overlap when
// aStart <= bEnd AND aEnd >= bStart. `excludeBookingId` lets an edit ignore its own row.
function findOverlappingBooking(bookings, newStart, newEnd, excludeBookingId) {
  const start = new Date(newStart);
  const end = new Date(newEnd);
  return (bookings || []).find((b) => {
    if (b.status === 'cancelled') return false;
    if (excludeBookingId && b.bookingId === excludeBookingId) return false;
    const bStart = new Date(b.startDate);
    const bEnd = new Date(b.endDate);
    return start <= bEnd && end >= bStart;
  });
}

module.exports = { calcDurationDays, calcBookingAmount, formatDateLabel, findOverlappingBooking, todayDateOnly };
