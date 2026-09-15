const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const bookingInfoSchema = new mongoose.Schema(
  {
    customerType: { type: String, enum: ['client', 'agency'], default: 'client' },
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client' },
    booking: { type: mongoose.Schema.Types.ObjectId, ref: 'Booking' },
    bookingRef: String,
    startDate: Date,
    endDate: Date,
    durationDays: Number,
    monthlyTotalCost: Number,
    amount: Number,
    bookedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

const blockInfoSchema = new mongoose.Schema(
  {
    reason: String,
    notes: String,
    blockedDate: { type: Date, default: nowIST },
    blockedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

const siteSchema = new mongoose.Schema(
  {
    mediaId: { type: String, required: true, unique: true, trim: true },
    mediaName: { type: String, trim: true },
    mediaType: { type: String, required: true, trim: true },
    quantity: { type: Number, default: 1, min: 0 },
    state: { type: String, required: true, trim: true },
    city: { type: String, required: true, trim: true },
    location: { type: String, trim: true },
    areaName: { type: String, trim: true },
    locationDetails: { type: String, trim: true },
    siteOwner: { type: String, trim: true, index: true },
    latitude: { type: Number, min: -90, max: 90 },
    longitude: { type: Number, min: -180, max: 180 },
    illumination: { type: String, trim: true },
    width: { type: Number, min: 0 },
    height: { type: Number, min: 0 },
    sizeUnit: { type: String, default: 'ft' },
    autoSize: { type: Number, min: 0 },
    amount: { type: Number, min: 0 },
    gstAmount: { type: Number, min: 0 },
    monthlyAmount: { type: Number, min: 0 },
    printingCost: { type: Number, min: 0, default: 0 },
    mountingCost: { type: Number, min: 0, default: 0 },
    totalCost: { type: Number, min: 0 },
    mediaImage: String,
    isActive: { type: Boolean, default: true },
    mediaStatus: {
      type: String,
      enum: ['available', 'booked', 'blocked'],
      default: 'available',
      index: true,
    },
    bookingInfo: bookingInfoSchema,
    blockInfo: blockInfoSchema,
    assignedTL: { type: mongoose.Schema.Types.ObjectId, ref: 'User', default: null },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
    inventoryUpdatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

function applyComputedFields(doc) {
  if (doc.width != null && doc.height != null) {
    doc.autoSize = Number(doc.width) * Number(doc.height);
  }
  const display = Number(doc.monthlyAmount) || 0;
  const printing = Number(doc.printingCost) || 0;
  const mounting = Number(doc.mountingCost) || 0;
  doc.totalCost = display + printing + mounting;
}

siteSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  applyComputedFields(this);
  if (this.$locals.inventoryOnly) {
    this.inventoryUpdatedAt = now;
  } else {
    this.updatedAt = now;
  }
  next();
});

siteSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

siteSchema.index({ mediaId: 'text', location: 'text', city: 'text', state: 'text', areaName: 'text' });

const Site = mongoose.model('Site', siteSchema);
Site.applyComputedFields = applyComputedFields;
module.exports = Site;
