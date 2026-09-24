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
    // Cancellation audit trail — set only when status is flipped to 'cancelled'. The booking
    // record itself is never deleted so Timeline/history can keep showing what was booked.
    cancellationReason: String,
    cancelledAt: Date,
    cancelledBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    cancelledByName: String,
    cancelledByRole: String,
    // 'blocked' = ended automatically because the site was blocked mid-booking; 'manual' = a user
    // cancelled it with their own reason. Lets the Timeline show the two differently.
    cancellationType: { type: String, enum: ['manual', 'blocked'] },
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
    quantity: { type: Number, default: 1, min: 0 }  ,
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

// Derived from other master fields (width/height, monthlyAmount/printingCost/mountingCost), so a
// recalculation alone is never treated as a user edit.
const DERIVED_FIELDS = ['autoSize', 'totalCost'];
const MASTER_COMPARE_FIELDS = MASTER_FIELDS.filter((f) => !DERIVED_FIELDS.includes(f));

// Audit metadata that controllers rewrite on every save (e.g. every booking gets a fresh
// updatedAt/updatedBy when the Edit Site form resubmits the bookings array) — not real data.
const IGNORED_NESTED_KEYS = new Set(['createdAt', 'updatedAt', 'createdBy', 'updatedBy', 'bookedBy', 'blockedDate', 'blockedBy', '_id']);

// Order-independent, type-normalized serialization: Dates by time value, ObjectIds by hex,
// and null/undefined/'' all treated as "empty".
function canonical(value, nested = false) {
  if (value === undefined || value === null || value === '') return null;
  if (value instanceof Date) return value.getTime();
  if (value && typeof value.toHexString === 'function') return value.toHexString();
  if (value && typeof value.toObject === 'function') value = value.toObject({ depopulate: true });
  if (Array.isArray(value)) return value.map((v) => canonical(v, true));
  if (typeof value === 'object') {
    const out = {};
    for (const key of Object.keys(value).sort()) {
      if (nested && IGNORED_NESTED_KEYS.has(key)) continue;
      const v = canonical(value[key], true);
      if (v !== null) out[key] = v;
    }
    return Object.keys(out).length ? out : null;
  }
  return value;
}

function snapshot(doc, fields) {
  return JSON.stringify(fields.map((f) => canonical(doc.get(f), true)));
}

// Remember the values as loaded from the DB so pre('save') can tell what really changed —
// Mongoose's isModified() reports a change whenever a subdocument/array is reassigned, even
// with identical content (resolveSiteStatus reassigns bookingInfo/bookings on every call).
siteSchema.post('init', function () {
  this.$locals.masterSnapshot = snapshot(this, MASTER_COMPARE_FIELDS);
  this.$locals.inventorySnapshot = snapshot(this, INVENTORY_FIELDS);
});

// Timestamp rule: driven by WHAT DATA changed, never by which page/endpoint was used.
// A Site edit bumps only `updatedAt`; an Inventory/status/booking change bumps only
// `inventoryUpdatedAt`; a single save touching both groups bumps both.
siteSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  applyComputedFields(this);

  const { masterSnapshot, inventorySnapshot } = this.$locals;
  const masterChanged =
    this.isNew || masterSnapshot === undefined || snapshot(this, MASTER_COMPARE_FIELDS) !== masterSnapshot;
  const inventoryChanged =
    this.isNew ||
    inventorySnapshot === undefined ||
    snapshot(this, INVENTORY_FIELDS) !== inventorySnapshot ||
    this.$locals.forceInventoryTouch;

  if (masterChanged) this.updatedAt = now;
  if (inventoryChanged) this.inventoryUpdatedAt = now;
  next();
});

// Refresh the baseline after a save so a second save of the same document compares against
// what is now in the DB.
siteSchema.post('save', function () {
  this.$locals.masterSnapshot = snapshot(this, MASTER_COMPARE_FIELDS);
  this.$locals.inventorySnapshot = snapshot(this, INVENTORY_FIELDS);
  this.$locals.forceInventoryTouch = false;
});

// Query-style updates: bump whichever timestamp matches the fields being written (e.g. the
// legacy Booking API only sets mediaStatus/bookingInfo, which is an inventory change).
function touchedRootFields(update) {
  const fields = new Set();
  for (const [key, val] of Object.entries(update || {})) {
    if (key.startsWith('$')) {
      if (val && typeof val === 'object') Object.keys(val).forEach((p) => fields.add(p.split('.')[0]));
    } else {
      fields.add(key.split('.')[0]);
    }
  }
  return fields;
}

siteSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  const fields = touchedRootFields(this.getUpdate());
  const now = nowIST();
  const set = {};
  if (MASTER_FIELDS.some((f) => fields.has(f))) set.updatedAt = now;
  if (INVENTORY_FIELDS.some((f) => fields.has(f))) set.inventoryUpdatedAt = now;
  if (Object.keys(set).length) this.set(set);
  next();
});

siteSchema.index({ mediaId: 'text', location: 'text', city: 'text', state: 'text', areaName: 'text' });

const Site = mongoose.model('Site', siteSchema);
Site.applyComputedFields = applyComputedFields;
Site.MASTER_FIELDS = MASTER_FIELDS;
Site.INVENTORY_FIELDS = INVENTORY_FIELDS;
module.exports = Site;
