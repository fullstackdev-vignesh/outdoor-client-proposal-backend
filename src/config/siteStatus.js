// Site media statuses — the single list every model, controller and job uses.
//   immediate — free; the only status a new proposal can use (formerly "available")
//   booked    — driven by the site's bookings (see services/bookingScheduler.js)
//   blocked   — blocked for a customer between its own Start and End Date (date-driven, like a
//               booking); older blocks saved without dates stay until changed by hand
//   confirmed — confirmed by a customer between its own Start and End Date — same flow as blocked
//   hold      — pending client query (amount, clarification…); stays until changed manually
//   issue     — non-client/site problem (corporation…); stays until changed, hidden from proposals
const MEDIA_STATUSES = ['immediate', 'blocked', 'confirmed', 'booked', 'hold', 'issue'];

// Blocked and Confirmed share one dated record (Site.blockInfo, with kind = the status).
const DATED_STATUSES = ['blocked', 'confirmed'];

// Statuses whose details (customer/dates/reason, hold/issue reason) come from the request.
const MANUAL_STATUSES = ['blocked', 'confirmed', 'hold', 'issue'];

// Statuses set by hand that booking/block-date reconciliation never overrides.
const OVERRIDE_STATUSES = ['hold', 'issue'];

const HISTORY_STATUSES = MEDIA_STATUSES;

// Any case/spacing ("Available", " Hold "); older clients/data may still send the pre-rename value.
function normalizeStatus(status) {
  if (typeof status !== 'string') return status;
  const s = status.trim().toLowerCase();
  return s === 'available' ? 'immediate' : s;
}

module.exports = { MEDIA_STATUSES, DATED_STATUSES, MANUAL_STATUSES, OVERRIDE_STATUSES, HISTORY_STATUSES, normalizeStatus };
