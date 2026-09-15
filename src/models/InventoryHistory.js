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

  status: { type: String, enum: ['available', 'booked', 'blocked'], required: true, index: true },
  previousStatus: { type: String, enum: ['available', 'booked', 'blocked', null], default: null },
  isActive: { type: Boolean, default: true },

  // Period the status applies to. effectiveTo = null means the period is still open/ongoing.
  effectiveFrom: { type: Date, required: true, index: true },
  effectiveTo: { type: Date, default: null, index: true },

  // When the user actually performed the update (distinct from the effective period).
  changedAt: { type: Date, default: nowIST },
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
});

inventoryHistorySchema.index({ site: 1, effectiveTo: 1 });
inventoryHistorySchema.index({ effectiveFrom: 1, effectiveTo: 1 });

module.exports = mongoose.model('InventoryHistory', inventoryHistorySchema);
