const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const inventoryHistorySchema = new mongoose.Schema({
  site: { type: mongoose.Schema.Types.ObjectId, ref: 'Site', required: true, index: true },
  mediaId: { type: String, required: true },
  mediaType: String,
  state: String,
  city: String,
  mediaImage: String,
  siteOwner: String,

  status: { type: String, enum: ['available', 'booked', 'blocked', 'cancelled'], required: true, index: true },
  previousStatus: { type: String, enum: ['available', 'booked', 'blocked', null], default: null },
  isActive: { type: Boolean, default: true },

  // Identifies which booking (Site.bookings[].bookingId) this row represents — lets a site
  // with several bookings get one independent Timeline row per booking instead of only its
  // current/active one. Unset for non-booking (available/blocked) rows.
  bookingId: { type: String, default: null, index: true },

  // Period the status applies to. effectiveTo = null means the period is still open/ongoing.
  effectiveFrom: { type: Date, required: true, index: true },
  effectiveTo: { type: Date, default: null, index: true },

  // When the user actually performed the update (distinct from the effective period).
  changedAt: { type: Date, default: nowIST },
  // Booking rows only: when the booking was first made. Unlike changedAt it is never
  // overwritten, so the Timeline can still show the original "Booked" step after a cancel.
  bookedAt: Date,
  // Last time this row's data actually changed (created, period closed, booking edited or
  // cancelled). Drives the Timeline list order. Older rows may lack it — readers fall back to
  // changedAt.
  updatedAt: { type: Date, default: nowIST, index: true },
  changedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  source: { type: String, enum: ['sites', 'inventory'], required: true },

  bookingSnapshot: {
    customerType: String,
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client' },
    customerName: String,
    startDate: Date,
    endDate: Date,
    durationDays: Number,
    monthlyTotalCost: Number,
    amount: Number,
  },
  blockSnapshot: {
    reason: String,
    notes: String,
    blockedDate: Date,
  },
  // Set only on the row for a booking that was cancelled — keeps `bookingSnapshot`'s original
  // dates intact (still visible as the "Booked Period") alongside why/when/by-whom it ended.
  cancellationSnapshot: {
    reason: String,
    cancelledAt: Date,
    cancelledByName: String,
    cancelledByRole: String,
    cancellationType: { type: String, enum: ['manual', 'blocked'] },
  },
});

inventoryHistorySchema.index({ site: 1, effectiveTo: 1 });
inventoryHistorySchema.index({ effectiveFrom: 1, effectiveTo: 1 });
// One Timeline row per booking — the upsert in syncBookingTimelineRecords relies on this to
// update in place instead of ever inserting a duplicate for the same bookingId.
inventoryHistorySchema.index({ site: 1, bookingId: 1 }, { unique: true, partialFilterExpression: { bookingId: { $type: 'string' } } });

module.exports = mongoose.model('InventoryHistory', inventoryHistorySchema);
