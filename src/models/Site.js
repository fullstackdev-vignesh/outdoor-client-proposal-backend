const mongoose = require('mongoose');
const { MEDIA_STATUSES } = require('../config/siteStatus');

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

// Details for the Hold / Issue statuses (Blocked keeps its own blockInfo).
const statusInfoSchema = new mongoose.Schema(
  {
    reason: String,
    notes: String,
    date: { type: Date, default: nowIST },
    by: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

// A Blocked or Confirmed period for one customer over its own date range (`kind` says which).
// startDate/endDate are date-only (UTC midnight, same as bookings): the site has that status from
// the Start Date through the End Date, and the period may be scheduled ahead. Older blocks have no
// dates and stay Blocked until changed by hand.
const blockInfoSchema = new mongoose.Schema(
  {
    kind: { type: String, enum: ['blocked', 'confirmed'], default: 'blocked' },
    reason: String,
    notes: String,
    customerType: { type: String, enum: ['client', 'agency'] },
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client' },
    customerName: String,
    startDate: Date,
    endDate: Date,
    blockedDate: { type: Date, default: nowIST },
    blockedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

// Further Blocked/Confirmed periods added with "+ Add Blocked" / "+ Add Confirmed" while the site
// already has one (blockInfo). Each waits here until its Start Date, then becomes the site's blockInfo
// (services/bookingScheduler.js#resolveSiteStatus) — like an Upcoming booking becoming Active.
const upcomingBlockSchema = new mongoose.Schema(
  {
    blockId: { type: String, required: true },
    kind: { type: String, enum: ['blocked', 'confirmed'], default: 'blocked' },
    reason: String,
    notes: String,
    customerType: { type: String, enum: ['client', 'agency'] },
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client' },
    customerName: String,
    startDate: { type: Date, required: true },
    endDate: { type: Date, required: true },
    blockedDate: { type: Date, default: nowIST },
    blockedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

// Fields that represent Site MASTER data — changing any of these bumps `updatedAt` only.
const MASTER_FIELDS = [
  'mediaId', 'mediaName', 'mediaType', 'quantity', 'state', 'city', 'location', 'areaName',
  'locationDetails', 'trafficViewFrom', 'trafficViewTo', 'specification', 'siteOwner', 'latitude', 'longitude', 'illumination', 'width', 'height',
  'sizeUnit', 'autoSize', 'amount', 'gstAmount', 'monthlyAmount', 'printingCost', 'mountingCost',
  'totalCost', 'mediaImage', 'mediaImages', 'siteInfoId',
  // Active/Inactive is the site's own setting (changed in Site Management), not inventory state.
  'isActive', 'inactiveReason',
];

// Fields that represent live Inventory/status/booking state — changing any of these bumps
// `inventoryUpdatedAt` only. A single save can bump BOTH if it touches both groups.
const INVENTORY_FIELDS = ['mediaStatus', 'bookingInfo', 'bookings', 'blockInfo', 'upcomingBlocks', 'statusInfo'];

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
    // Traffic direction the site faces, e.g. "Madurai" -> "Trichy".
    trafficViewFrom: { type: String, trim: true },
    trafficViewTo: { type: String, trim: true },
    // Size as shown to clients, e.g. "30 x 20" — filled from Width x Height in the form but editable.
    specification: { type: String, trim: true },
    siteOwner: { type: String, trim: true, index: true },
    latitude: { type: Number, min: -90, max: 90 },
    longitude: { type: Number, min: -180, max: 180 },
    illumination: { type: String, trim: true },
    width: { type: Number, min: 0 },
    height: { type: Number, min: 0 },
    sizeUnit: { type: String, default: 'Sq.ft' },
    autoSize: { type: Number, min: 0 },
    amount: { type: Number, min: 0 },
    gstAmount: { type: Number, min: 0 },
    monthlyAmount: { type: Number, min: 0 },
    printingCost: { type: Number, min: 0, default: 0 },
    mountingCost: { type: Number, min: 0, default: 0 },
    totalCost: { type: Number, min: 0 },
    // Default image — the one every list, PPT and export shows.
    mediaImage: String,
    // Every image saved for the site (includes the default). Older sites only have mediaImage.
    mediaImages: { type: [String], default: undefined },
    // Optional link to a reusable Site Info card (title + description) shown on PPT templates
    // that support it (e.g. Adinn-Direct-Client-format). Only the reference is stored here —
    // the actual title/description text lives on the SiteInfo document, never duplicated here.
    siteInfoId: { type: mongoose.Schema.Types.ObjectId, ref: 'SiteInfo', default: null },
    isActive: { type: Boolean, default: true },
    // Why the site was made Inactive (required then); cleared when it's made Active again.
    inactiveReason: { type: String, trim: true },
    mediaStatus: {
      type: String,
      enum: MEDIA_STATUSES,
      default: 'immediate',
      index: true,
    },
    // Cached snapshot of the currently-active booking (or undefined). Kept in sync by
    // services/bookingScheduler.js#resolveSiteStatus — never edited directly by controllers.
    bookingInfo: bookingInfoSchema,
    bookings: { type: [bookingRecordSchema], default: [] },
    blockInfo: blockInfoSchema,
    upcomingBlocks: { type: [upcomingBlockSchema], default: [] },
    statusInfo: statusInfoSchema,
    assignedTL: { type: mongoose.Schema.Types.ObjectId, ref: 'User', default: null },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedBy: { type: String, default: 'System' },
    inventoryUpdatedBy: { type: String, default: 'System' },
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
const IGNORED_NESTED_KEYS = new Set(['createdAt', 'updatedAt', 'inventoryUpdatedAt', 'createdBy', 'updatedBy', 'inventoryUpdatedBy', 'bookedBy', 'blockedDate', 'blockedBy', '_id']);

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

  // A status/booking change made from Site Management ($locals.changeSource = 'sites') is a site
  // update — it bumps "Site Updated", not "Inventory Updated". From the Inventory page, bulk status
  // and the automatic date checks it stays an inventory update.
  const fromSitePage = this.$locals.changeSource === 'sites';
  if (masterChanged || (inventoryChanged && fromSitePage && !this.isNew)) {
    this.updatedAt = now;
    if (this.$locals.currentUserName) this.updatedBy = this.$locals.currentUserName;
  }
  if (inventoryChanged && (!fromSitePage || this.isNew)) {
    this.inventoryUpdatedAt = now;
    if (this.$locals.currentUserName) this.inventoryUpdatedBy = this.$locals.currentUserName;
  }
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
