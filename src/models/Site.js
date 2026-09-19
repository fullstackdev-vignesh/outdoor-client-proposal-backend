const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const bookingInfoSchema = new mongoose.Schema(
  {
    bookingId: String,
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

// One row per booking order. A site can hold many, as long as their date ranges never
// overlap (enforced in the controller/service layer, not here). `bookingInfo` above stays
// as a cached snapshot of whichever booking is CURRENTLY ACTIVE (or null), recomputed by
// services/bookingScheduler.js — every existing consumer of site.bookingInfo keeps working
// unchanged.
const bookingRecordSchema = new mongoose.Schema(
  {
    bookingId: { type: String, required: true },
    customerType: { type: String, enum: ['client', 'agency'], default: 'client' },
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client' },
    customerName: String,
    startDate: { type: Date, required: true },
    endDate: { type: Date, required: true },
    durationDays: Number,
    monthlyTotalCost: Number,
    amount: Number,
    status: { type: String, enum: ['upcoming', 'active', 'completed', 'cancelled'], default: 'upcoming' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
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

// Fields that represent Site MASTER data — changing any of these bumps `updatedAt` only.
const MASTER_FIELDS = [
  'mediaId', 'mediaName', 'mediaType', 'quantity', 'state', 'city', 'location', 'areaName',
  'locationDetails', 'siteOwner', 'latitude', 'longitude', 'illumination', 'width', 'height',
  'sizeUnit', 'autoSize', 'amount', 'gstAmount', 'monthlyAmount', 'printingCost', 'mountingCost',
  'totalCost', 'mediaImage', 'siteInfoId',
];

// Fields that represent live Inventory/status/booking state — changing any of these bumps
// `inventoryUpdatedAt` only. A single save can bump BOTH if it touches both groups.
const INVENTORY_FIELDS = ['mediaStatus', 'bookingInfo', 'bookings', 'blockInfo', 'isActive'];

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
    // Optional link to a reusable Site Info card (title + description) shown on PPT templates
    // that support it (e.g. Adinn-Direct-Client-format). Only the reference is stored here —
    // the actual title/description text lives on the SiteInfo document, never duplicated here.
    siteInfoId: { type: mongoose.Schema.Types.ObjectId, ref: 'SiteInfo', default: null },
    isActive: { type: Boolean, default: true },
    mediaStatus: {
      type: String,
      enum: ['available', 'booked', 'blocked'],
      default: 'available',
      index: true,
    },
    // Cached snapshot of the currently-active booking (or undefined). Kept in sync by
    // services/bookingScheduler.js#resolveSiteStatus — never edited directly by controllers.
    bookingInfo: bookingInfoSchema,
    bookings: { type: [bookingRecordSchema], default: [] },
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
    const autoSize = Number(doc.width) * Number(doc.height);
    if (doc.autoSize !== autoSize) doc.autoSize = autoSize;
  }
  const display = Number(doc.monthlyAmount) || 0;
  const printing = Number(doc.printingCost) || 0;
  const mounting = Number(doc.mountingCost) || 0;
  const totalCost = display + printing + mounting;
  if (doc.totalCost !== totalCost) doc.totalCost = totalCost;
}

// Timestamp rule: driven by WHAT DATA changed, never by which page/endpoint was used.
siteSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  applyComputedFields(this);

  const masterChanged = this.isNew || MASTER_FIELDS.some((f) => this.isModified(f));
  const inventoryChanged = this.isNew || INVENTORY_FIELDS.some((f) => this.isModified(f)) || this.$locals.forceInventoryTouch;

  if (masterChanged) this.updatedAt = now;
  if (inventoryChanged) this.inventoryUpdatedAt = now;
  next();
});

siteSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

siteSchema.index({ mediaId: 'text', location: 'text', city: 'text', state: 'text', areaName: 'text' });

const Site = mongoose.model('Site', siteSchema);
Site.applyComputedFields = applyComputedFields;
Site.MASTER_FIELDS = MASTER_FIELDS;
Site.INVENTORY_FIELDS = INVENTORY_FIELDS;
module.exports = Site;
