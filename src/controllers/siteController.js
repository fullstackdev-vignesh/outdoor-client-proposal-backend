const asyncHandler = require('express-async-handler');
const ExcelJS = require('exceljs');
const Site = require('../models/Site');
const SiteHistory = require('../models/SiteHistory');
const Client = require('../models/Client');
const { saveMediaImage } = require('../utils/imageStorage');
const { calcDurationDays, calcBookingAmount, formatDateLabel, findOverlappingBooking, todayDateOnly } = require('../utils/bookingCalc');
const { formatIST } = require('../utils/formatDate');
const InventoryHistory = require('../models/InventoryHistory');
const { recordStatusPeriod, buildOverlapFilter, syncBookingTimelineRecords, computeBookingLifecycle } = require('../services/inventoryTimeline');
const { genBookingId, resolveSiteStatus } = require('../services/bookingScheduler');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const TRACKED_FIELDS = [
  'mediaId', 'mediaType', 'quantity', 'state', 'city', 'location', 'areaName', 'locationDetails', 'siteOwner',
  'latitude', 'longitude', 'illumination', 'width', 'height', 'sizeUnit', 'amount', 'gstAmount',
  'monthlyAmount', 'printingCost', 'mountingCost', 'totalCost', 'mediaImage', 'isActive', 'mediaStatus',
  'blockInfo.reason', 'blockInfo.notes',
];

function getPath(obj, path) {
  return path.split('.').reduce((o, k) => (o == null ? undefined : o[k]), obj);
}

// `changedAt` is hoisted by the caller so every SiteHistory row written for the SAME save
// operation shares the exact same timestamp — the Edit History UI groups rows into one
// event/card by matching this value exactly, rather than guessing from near-identical times.
async function logFieldChanges(siteId, before, after, userId, changedAt) {
  const ts = changedAt || nowIST();
  const entries = [];
  for (const field of TRACKED_FIELDS) {
    const oldVal = before ? getPath(before, field) : undefined;
    const newVal = after ? getPath(after, field) : undefined;
    const oldStr = oldVal === undefined || oldVal === null ? '' : String(oldVal);
    const newStr = newVal === undefined || newVal === null ? '' : String(newVal);
    if (oldStr !== newStr) {
      entries.push({ site: siteId, field, oldValue: oldVal, newValue: newVal, changedBy: userId, changedAt: ts });
    }
  }
  if (entries.length) await SiteHistory.insertMany(entries);
}

const money = (n) => (n == null ? '' : `₹${Number(n).toLocaleString('en-IN')}`);
const days = (n) => (n == null ? '' : `${n} Days`);

// Diffs the site's booking array (by bookingId) into readable audit rows: a brand-new
// bookingId becomes one "Booking Added" row; an existing bookingId whose fields changed
// becomes one row per changed field (Customer/Start Date/End Date/Duration/Booking Amount).
async function logBookingChanges(siteId, beforeBookings, afterBookings, userId, changedAt) {
  const ts = changedAt || nowIST();
  const entries = [];
  const beforeById = new Map((beforeBookings || []).map((b) => [b.bookingId, b]));

  (afterBookings || []).forEach((b, idx) => {
    const prev = beforeById.get(b.bookingId);
    if (!prev) {
      entries.push({
        site: siteId,
        field: 'Booking Added',
        oldValue: null,
        newValue: `Booking #${idx + 1}: ${b.customerName || 'Customer'} (${formatDateLabel(b.startDate)} → ${formatDateLabel(b.endDate)})`,
        changedBy: userId,
        changedAt: ts,
      });
      return;
    }
    const pairs = [
      ['Booking Customer', prev.customerName, b.customerName],
      ['Booking Start Date', formatDateLabel(prev.startDate), formatDateLabel(b.startDate)],
      ['Booking End Date', formatDateLabel(prev.endDate), formatDateLabel(b.endDate)],
      ['Booking Duration', days(prev.durationDays), days(b.durationDays)],
      ['Booking Amount', money(prev.amount), money(b.amount)],
    ];
    for (const [field, oldValue, newValue] of pairs) {
      if ((oldValue || '') !== (newValue || '')) {
        entries.push({ site: siteId, field, oldValue, newValue, changedBy: userId, changedAt: ts });
      }
    }
  });

  if (entries.length) await SiteHistory.insertMany(entries);
}

// Marks a single booking record cancelled in place (never removes it — Timeline/history need
// the original booking to still exist). Caller is responsible for site.markModified('bookings'),
// resolveSiteStatus(site) and site.save() afterwards.
function applyBookingCancellation(booking, reason, user, type = 'manual') {
  booking.status = 'cancelled';
  booking.cancellationType = type;
  booking.cancellationReason = reason;
  booking.cancelledAt = nowIST();
  booking.cancelledBy = user._id;
  booking.cancelledByName = user.name;
  booking.cancelledByRole = user.role;
}

const utcDay = (value) => {
  const d = new Date(value);
  return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
};

// Leaving the Blocked state (to Booked or Available). Must run BEFORE resolveSiteStatus, which
// deliberately leaves blocked sites untouched — otherwise a blocked site could never be
// re-booked or made available again.
//
// Blocking interrupts whatever booking was running, so the booking that covers TODAY is ended
// (recorded as cancelled) on the way out — otherwise status recomputation would silently flip
// the site back to Booked on "Available", and a new booking entered on "Booked" would be
// rejected as overlapping it. Future (upcoming) bookings are kept.
// `wasBlocked` must come from the saved (before) state, never from stale fields on the doc.
// Returns the bookings it cancelled so the caller can log them.
function unblockSite(site, { wasBlocked, user, cancelCurrentBooking = true }) {
  if (!wasBlocked) return [];
  site.mediaStatus = 'available';
  site.blockInfo = undefined;
  site.isActive = true;

  if (!cancelCurrentBooking) return [];
  const today = utcDay(new Date());
  const cancelled = [];
  for (const b of site.bookings || []) {
    if (b.status === 'cancelled') continue;
    if (utcDay(b.startDate) <= today && utcDay(b.endDate) >= today) {
      applyBookingCancellation(b, 'Booking ended — site was blocked', user, 'blocked');
      cancelled.push(b);
    }
  }
  if (cancelled.length) site.markModified('bookings');
  return cancelled;
}

// Bookings made from the status popup may arrive with only a client id. Store the client's name
// on every booking that lacks one (the new booking and any older ones), so lists and the
// Timeline can always show "Client: <name>".
async function fillMissingCustomerNames(site) {
  const missing = (site.bookings || []).filter((b) => !b.customerName && b.client);
  if (!missing.length) return;
  const ids = [...new Set(missing.map((b) => String(b.client)))];
  const clients = await Client.find({ _id: { $in: ids } }).select('name').lean();
  const nameById = new Map(clients.map((c) => [String(c._id), c.name]));
  let changed = false;
  for (const b of missing) {
    const name = nameById.get(String(b.client));
    if (name) {
      b.customerName = name;
      changed = true;
    }
  }
  if (changed) site.markModified('bookings');
}

// One explicit, readable SiteHistory row for a cancellation — deliberately not run through
// logBookingChanges' generic field-diff format, since "Booking Cancelled" is a business event,
// not a plain old-value → new-value edit.
async function logBookingCancellation(siteId, booking, userId, changedAt) {
  const ts = changedAt || nowIST();
  const cancelledByLabel = booking.cancelledByName
    ? `${booking.cancelledByName}${booking.cancelledByRole ? ` (${booking.cancelledByRole.toUpperCase()})` : ''}`
    : '';
  const newValue = [
    `Booking: ${formatDateLabel(booking.startDate)} → ${formatDateLabel(booking.endDate)}`,
    `Reason: ${booking.cancellationReason}`,
    cancelledByLabel && `Cancelled By: ${cancelledByLabel}`,
    `Cancelled At: ${formatIST(booking.cancelledAt)}`,
  ]
    .filter(Boolean)
    .join(' | ');
  await SiteHistory.insertMany([
    { site: siteId, field: 'Booking Cancelled', oldValue: null, newValue, changedBy: userId, changedAt: ts },
  ]);
}

const buildFilter = (query) => {
  const filter = {};
  if (query.search) {
    filter.$or = [
      { mediaId: new RegExp(query.search, 'i') },
      { mediaType: new RegExp(query.search, 'i') },
      { city: new RegExp(query.search, 'i') },
      { state: new RegExp(query.search, 'i') },
      { location: new RegExp(query.search, 'i') },
      { areaName: new RegExp(query.search, 'i') },
    ];
  }
  if (query.mediaType) filter.mediaType = query.mediaType;
  if (query.state) filter.state = new RegExp(`^${escapeRegex(query.state.trim())}$`, 'i');
  if (query.city) filter.city = new RegExp(escapeRegex(query.city.trim()), 'i');
  if (query.mediaStatus) filter.mediaStatus = query.mediaStatus;
  if (query.siteOwner) filter.siteOwner = query.siteOwner;
  if (query.isActive !== undefined && query.isActive !== '') filter.isActive = query.isActive === 'true';
  if (query.minPrice || query.maxPrice) {
    filter.monthlyAmount = {};
    if (query.minPrice) filter.monthlyAmount.$gte = Number(query.minPrice);
    if (query.maxPrice) filter.monthlyAmount.$lte = Number(query.maxPrice);
  }
  return filter;
};

// Each page orders by its own timestamp: the Inventory page passes `sortBy=inventory` (latest
// status/booking/block change first); everything else — the Sites page — orders by the latest
// master-data edit (`updatedAt`). The two never mix.
function siteListSort(query) {
  return query.sortBy === 'inventory' ? { inventoryUpdatedAt: -1, _id: -1 } : { updatedAt: -1, _id: -1 };
}

const getSites = asyncHandler(async (req, res) => {
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  const filter = buildFilter(req.query);

  const [items, total] = await Promise.all([
    Site.find(filter).sort(siteListSort(req.query)).skip((page - 1) * limit).limit(limit),
    Site.countDocuments(filter),
  ]);

  res.json({ items: items.map((s) => ({ ...s.toObject(), mediaCode: s.mediaId })), total, page, pages: Math.ceil(total / limit) });
});

const getAvailableSites = asyncHandler(async (req, res) => {
  const filter = { ...buildFilter(req.query), mediaStatus: 'available', isActive: true };
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);

  const [items, total] = await Promise.all([
    Site.find(filter).sort(siteListSort(req.query)).skip((page - 1) * limit).limit(limit),
    Site.countDocuments(filter),
  ]);

  res.json({ items, total, page, pages: Math.ceil(total / limit) });
});

const getSite = asyncHandler(async (req, res) => {
  const site = await Site.findById(req.params.id)
    .populate('bookingInfo.client', 'name phone email')
    .populate('blockInfo.blockedBy', 'name');
  if (!site) {
    res.status(404);
    throw new Error('Site not found');
  }
  res.json({ ...site.toObject(), mediaCode: site.mediaId });
});

function validateSitePayload(body) {
  const errors = [];
  const num = (v) => (v === '' || v === undefined || v === null ? undefined : Number(v));

  if (body.width !== undefined && (isNaN(num(body.width)) || num(body.width) <= 0)) errors.push('Width must be a number greater than 0');
  if (body.height !== undefined && (isNaN(num(body.height)) || num(body.height) <= 0)) errors.push('Height must be a number greater than 0');
  // Blank means "not given" (e.g. the optional Display Cost Per Month cleared on Edit).
  ['amount', 'gstAmount', 'monthlyAmount', 'printingCost', 'mountingCost'].forEach((f) => {
    if (num(body[f]) !== undefined && (isNaN(num(body[f])) || num(body[f]) < 0)) errors.push(`${f} must be a number greater than or equal to 0`);
  });
  if (body.latitude !== undefined && body.latitude !== '' && (isNaN(num(body.latitude)) || num(body.latitude) < -90 || num(body.latitude) > 90)) {
    errors.push('Latitude must be between -90 and 90');
  }
  if (body.longitude !== undefined && body.longitude !== '' && (isNaN(num(body.longitude)) || num(body.longitude) < -180 || num(body.longitude) > 180)) {
    errors.push('Longitude must be between -180 and 180');
  }
  if (body.mediaStatus && !['available', 'booked', 'blocked'].includes(body.mediaStatus)) {
    errors.push('Invalid media status');
  }
  return errors;
}

function normalizeSiteBody(body) {
  const payload = { ...body };
  if (payload.mediaCode && !payload.mediaId) payload.mediaId = payload.mediaCode;
  delete payload.mediaCode;
  // multipart/form-data (used by Add/Edit Site so the image uploads in the same request)
  // sends nested objects/arrays as a JSON string.
  if (typeof payload.bookingInfo === 'string' && payload.bookingInfo) {
    try {
      payload.bookingInfo = JSON.parse(payload.bookingInfo);
    } catch {
      delete payload.bookingInfo;
    }
  }
  if (typeof payload.bookings === 'string' && payload.bookings) {
    try {
      payload.bookings = JSON.parse(payload.bookings);
    } catch {
      delete payload.bookings;
    }
  }
  // Optional Site Info dropdown submits '' when left unselected or 'None' is chosen.
  // Convert empty string / 'none' / 'null' to null so Mongoose sets siteInfoId = null in MongoDB.
  if (
    payload.siteInfoId === '' ||
    payload.siteInfoId === 'none' ||
    payload.siteInfoId === 'null' ||
    payload.siteInfoId === null
  ) {
    payload.siteInfoId = null;
  }
  return payload;
}

function extractUploadedFile(req) {
  if (req.file) return req.file;
  if (req.files) {
    if (req.files.mediaImage && req.files.mediaImage[0]) return req.files.mediaImage[0];
    if (req.files.image && req.files.image[0]) return req.files.image[0];
    if (req.files.file && req.files.file[0]) return req.files.file[0];
    const keys = Object.keys(req.files);
    if (keys.length > 0 && req.files[keys[0]][0]) return req.files[keys[0]][0];
  }
  return null;
}

// Uploads the newly selected image (if any) via the existing storage logic and sets
// payload.mediaImage to its public URL. If no new file was sent, mediaImage is left
// untouched so an update never clears/overwrites the site's existing image.
async function applyUploadedImage(payload, req) {
  const file = extractUploadedFile(req);
  if (file) {
    payload.mediaImage = await saveMediaImage(file);
  } else {
    delete payload.mediaImage;
  }
}

// Builds one booking record from raw input, checking it against `existingBookings` for a
// date overlap (excluding its own bookingId when editing). Throws a friendly, user-facing
// message — never allows an overlapping save.
function buildBookingRecord(site, input, userId, existingBooking) {
  const { customerType, client, customerName, startDate, endDate } = input || {};
  if (!client) throw new Error('Customer is required');
  if (!startDate || !endDate) throw new Error('Start Date and End Date are required');
  if (new Date(endDate) < new Date(startDate)) throw new Error('End Date must be on or after Start Date');
  // Only a brand-new booking is floored at today — editing an already-saved booking (it
  // has an existingBooking record) keeps its own historical start date valid.
  if (!existingBooking && startDate < todayDateOnly()) throw new Error('Start Date cannot be before today.');

  const overlap = findOverlappingBooking(site.bookings, startDate, endDate, existingBooking?.bookingId);
  if (overlap) {
    throw new Error(`This site is already booked from ${formatDateLabel(overlap.startDate)} to ${formatDateLabel(overlap.endDate)}.`);
  }

  const durationDays = calcDurationDays(startDate, endDate);
  const monthlyTotalCost = site.totalCost || site.monthlyAmount || 0;
  const amount = calcBookingAmount(monthlyTotalCost, durationDays);

  return {
    bookingId: existingBooking?.bookingId || genBookingId(),
    customerType: customerType === 'agency' ? 'agency' : 'client',
    client,
    customerName,
    startDate,
    endDate,
    durationDays,
    monthlyTotalCost,
    amount,
    status: existingBooking?.status || 'upcoming',
    createdAt: existingBooking?.createdAt || nowIST(),
    updatedAt: nowIST(),
    createdBy: existingBooking?.createdBy || userId,
    updatedBy: userId,
  };
}

// Table quick-action path (StatusChangeModal/BulkStatusModal) — one booking at a time.
// Edits the site's currently ACTIVE booking in place if there is one (so adjusting dates on
// today's campaign doesn't create a duplicate row); otherwise appends a new booking.
// `mode: 'new'` always appends a separate booking (e.g. a second client for later dates) — it
// never touches the existing ones; the overlap check still rejects clashing dates.
function upsertActiveBooking(site, input, userId, { mode = 'edit' } = {}) {
  if (mode === 'new') {
    const record = buildBookingRecord(site, input, userId, null);
    site.bookings = [...(site.bookings || []), record];
    return record;
  }
  const activeId = site.mediaStatus === 'booked' ? site.bookingInfo?.bookingId : undefined;
  const bookings = site.bookings || [];
  const existingIndex = activeId ? bookings.findIndex((b) => b.bookingId === activeId) : -1;
  const existing = existingIndex >= 0 ? bookings[existingIndex] : null;

  const record = buildBookingRecord(site, input, userId, existing);
  if (existingIndex >= 0) {
    bookings[existingIndex] = record;
  } else {
    bookings.push(record);
  }
  site.bookings = bookings;
  return record;
}

// Add/Edit Site's full "+ Add Booking" array path — the form always submits the complete,
// authoritative list of bookings for the site. Every entry is re-validated (including
// pairwise overlap across the whole set) and rebuilt; existing bookingIds are preserved so
// history/createdAt/createdBy survive the round-trip.
function buildBookingsArray(site, incomingBookings, userId) {
  const existingById = new Map((site.bookings || []).map((b) => [b.bookingId, b]));
  const built = [];

  for (const input of incomingBookings) {
    const existing = input.bookingId ? existingById.get(input.bookingId) : null;
    const overlap = findOverlappingBooking(built, input.startDate, input.endDate, input.bookingId);
    if (overlap) {
      throw new Error(`This site is already booked from ${formatDateLabel(overlap.startDate)} to ${formatDateLabel(overlap.endDate)}.`);
    }
    if (!input.client) throw new Error('Customer is required for every booking');
    if (!input.startDate || !input.endDate) throw new Error('Start Date and End Date are required for every booking');
    if (new Date(input.endDate) < new Date(input.startDate)) throw new Error('End Date must be on or after Start Date');
    // Only a brand-new booking (no matching existing record) is floored at today — an
    // already-saved booking round-tripping through this array keeps its historical date.
    if (!existing && input.startDate < todayDateOnly()) throw new Error('Start Date cannot be before today.');

    const durationDays = calcDurationDays(input.startDate, input.endDate);
    const monthlyTotalCost = site.totalCost || site.monthlyAmount || 0;
    const amount = calcBookingAmount(monthlyTotalCost, durationDays);

    built.push({
      bookingId: existing?.bookingId || input.bookingId || genBookingId(),
      customerType: input.customerType === 'agency' ? 'agency' : 'client',
      client: input.client,
      customerName: input.customerName,
      startDate: input.startDate,
      endDate: input.endDate,
      durationDays,
      monthlyTotalCost,
      amount,
      status: existing?.status || 'upcoming',
      createdAt: existing?.createdAt || nowIST(),
      updatedAt: nowIST(),
      createdBy: existing?.createdBy || userId,
      updatedBy: userId,
    });
  }

  return built;
}

const createSite = asyncHandler(async (req, res) => {
  const payload = normalizeSiteBody(req.body);
  const errors = validateSitePayload(payload);
  if (errors.length) {
    res.status(400);
    throw new Error(errors.join('; '));
  }
  await applyUploadedImage(payload, req);
  const userName = req.user?.name || 'System';
  payload.createdBy = req.user._id;
  payload.updatedBy = userName;
  payload.inventoryUpdatedBy = userName;
  if (!payload.mediaStatus) payload.mediaStatus = 'available';
  if (!payload.illumination) payload.illumination = 'Front Lit';

  const { blockReason, blockNotes, bookingInfo, bookings: incomingBookings } = payload;
  delete payload.blockReason;
  delete payload.blockNotes;
  delete payload.bookingInfo;
  delete payload.bookings;

  Site.applyComputedFields(payload);

  if (payload.mediaStatus === 'blocked') {
    if (!blockReason) {
      res.status(400);
      throw new Error('Block reason is required');
    }
    payload.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
    payload.isActive = false;
  } else if (payload.mediaStatus === 'booked') {
    const bookingList = Array.isArray(incomingBookings) && incomingBookings.length ? incomingBookings : bookingInfo ? [bookingInfo] : [];
    if (!bookingList.length) {
      res.status(400);
      throw new Error('At least one booking is required');
    }
    try {
      payload.bookings = buildBookingsArray({ bookings: [], totalCost: payload.totalCost, monthlyAmount: payload.monthlyAmount }, bookingList, req.user._id);
    } catch (err) {
      res.status(400);
      throw err;
    }
    payload.isActive = true;
  } else {
    payload.isActive = true;
  }

  let site;
  try {
    site = await Site.create(payload);
  } catch (err) {
    if (err.code === 11000 && err.keyPattern?.mediaId) {
      res.status(400);
      throw new Error('MediaCode already exists. Please use a different MediaCode.');
    }
    throw err;
  }

  // The live status must reflect TODAY against the booking dates just saved — a booking
  // starting in the future keeps the site Available until its start date arrives.
  resolveSiteStatus(site);
  await site.save();

  await recordStatusPeriod({ site, previousStatus: null, source: 'sites', userId: req.user._id });
  await syncBookingTimelineRecords(site, req.user._id, 'sites');
  res.status(201).json({ ...site.toObject(), mediaCode: site.mediaId });
});

const updateSite = asyncHandler(async (req, res) => {
  const site = await Site.findById(req.params.id);
  if (!site) {
    res.status(404);
    throw new Error('Site not found');
  }
  const payload = normalizeSiteBody(req.body);
  const errors = validateSitePayload(payload);
  if (errors.length) {
    res.status(400);
    throw new Error(errors.join('; '));
  }
  await applyUploadedImage(payload, req);
  const before = site.toObject();
  const beforeBookings = before.bookings || [];
  const beforeStatus = before.mediaStatus;
  const beforeBookingId = before.bookingInfo?.bookingId;

  const { blockReason, blockNotes, bookingInfo, bookings: incomingBookings, ...siteFields } = payload;
  let unblockCancelled = [];
  site.$locals.currentUserName = req.user?.name || 'System';
  Object.assign(site, siteFields);
  Site.applyComputedFields(site);

  if (site.mediaStatus === 'blocked') {
    if (!blockReason) {
      res.status(400);
      throw new Error('Block reason is required');
    }
    // Re-saving an already-blocked site from Edit Site keeps its original block record (date/by)
    // unless the reason or notes were actually changed.
    const sameBlock =
      before.mediaStatus === 'blocked' &&
      before.blockInfo?.reason === blockReason &&
      (before.blockInfo?.notes || '') === (blockNotes || '');
    if (!sameBlock) {
      site.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
    }
    site.isActive = false;
  } else {
    const targetStatus = site.mediaStatus;
    const bookingList = Array.isArray(incomingBookings) ? incomingBookings : bookingInfo ? [bookingInfo] : null;
    if (targetStatus === 'booked' && bookingList) {
      try {
        site.bookings = buildBookingsArray(site, bookingList, req.user._id);
      } catch (err) {
        res.status(400);
        throw err;
      }
    }
    // Coming out of Blocked. → Available: the booking that was running when the site got blocked
    // ends, so the site really becomes Available. → Booked: the form's booking rows (which the
    // user sees and edits) are the source of truth, so nothing is auto-cancelled.
    unblockCancelled = unblockSite(site, {
      wasBlocked: beforeStatus === 'blocked',
      user: req.user,
      cancelCurrentBooking: targetStatus !== 'booked',
    });
    // Live status/bookingInfo always get recomputed from the (possibly just-edited)
    // bookings array against today's date — never taken at face value from the toggle.
    resolveSiteStatus(site);
  }

  await fillMissingCustomerNames(site);
  try {
    await site.save();
  } catch (err) {
    if (err.code === 11000 && err.keyPattern?.mediaId) {
      res.status(400);
      throw new Error('MediaCode already exists. Please use a different MediaCode.');
    }
    throw err;
  }

  const statusChanged = beforeStatus !== site.mediaStatus || beforeBookingId !== site.bookingInfo?.bookingId;
  if (statusChanged) {
    await recordStatusPeriod({ site, previousStatus: beforeStatus, source: 'sites', userId: req.user._id });
  }
  // Unconditional (not gated by statusChanged): every booking in the array — active,
  // upcoming or completed — must show up/refresh in Timeline as soon as it's saved, not
  // only the one currently driving the site's live status.
  await syncBookingTimelineRecords(site, req.user._id, 'sites');
  const changedAt = nowIST();
  await logFieldChanges(site._id, before, site.toObject(), req.user._id, changedAt);
  await logBookingChanges(site._id, beforeBookings, site.bookings, req.user._id, changedAt);
  for (const b of unblockCancelled) {
    await logBookingCancellation(site._id, b, req.user._id, changedAt);
  }
  res.json({ ...site.toObject(), mediaCode: site.mediaId });
});

const deleteSite = asyncHandler(async (req, res) => {
  const site = await Site.findById(req.params.id);
  if (!site) {
    res.status(404);
    throw new Error('Site not found');
  }
  await site.deleteOne();
  res.json({ message: 'Site deleted' });
});

const changeStatus = asyncHandler(async (req, res) => {
  const site = await Site.findById(req.params.id);
  if (!site) {
    res.status(404);
    throw new Error('Site not found');
  }
  const { mediaStatus, blockReason, blockNotes, bookingInfo, source, cancellationReason, reason, bookingMode } = req.body;
  const before = site.toObject();
  const beforeBookings = before.bookings || [];
  const beforeStatus = before.mediaStatus;
  const beforeBookingId = before.bookingInfo?.bookingId;

  if (!['available', 'booked', 'blocked'].includes(mediaStatus)) {
    res.status(400);
    throw new Error('Invalid media status');
  }

  let cancelledBookings = [];

  if (mediaStatus === 'blocked') {
    if (!blockReason) {
      res.status(400);
      throw new Error('Block reason is required');
    }
    site.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
    site.mediaStatus = 'blocked';
    site.isActive = false;
  } else {
    // Blocked → Booked/Available: unblock first and end the booking that was running when the
    // site got blocked — "Available" then really means Available, and "Booked" uses only the
    // new booking details entered now (no overlap with the interrupted booking).
    cancelledBookings = unblockSite(site, { wasBlocked: beforeStatus === 'blocked', user: req.user });
    if (mediaStatus === 'booked') {
      try {
        upsertActiveBooking(site, bookingInfo, req.user._id, { mode: bookingMode === 'new' ? 'new' : 'edit' });
      } catch (err) {
        res.status(400);
        throw err;
      }
    } else if (mediaStatus === 'available' && beforeStatus === 'booked') {
      // Going straight from Booked to Available is really "cancel the booking that's currently
      // making this site Booked" — never a silent status flip. Require and record a reason,
      // and only cancel the ONE booking driving today's live status (not every booking).
      const cancelReason = (cancellationReason || reason || '').trim();
      if (!cancelReason) {
        res.status(400);
        throw new Error('Cancellation reason is required to change from Booked to Available');
      }
      const activeId = before.bookingInfo?.bookingId;
      const activeBooking = activeId ? (site.bookings || []).find((b) => b.bookingId === activeId) : null;
      if (activeBooking && activeBooking.status !== 'cancelled') {
        applyBookingCancellation(activeBooking, cancelReason, req.user);
        site.markModified('bookings');
        cancelledBookings.push(activeBooking);
      }
    }
    // Live status/bookingInfo are always recomputed from the (remaining valid) bookings array
    // against today's date, not taken at face value from the requested `mediaStatus` — so an
    // Upcoming booking elsewhere in the array keeps the site Booked even after this cancellation.
    resolveSiteStatus(site);
  }

  await fillMissingCustomerNames(site);
  site.$locals.currentUserName = req.user?.name || 'System';
  await site.save();

  const statusChanged = beforeStatus !== site.mediaStatus || beforeBookingId !== site.bookingInfo?.bookingId;
  const resolvedSource = source === 'inventory' ? 'inventory' : 'sites';
  if (statusChanged) {
    // `source` still labels the Timeline entry "via Sites"/"via Inventory" — it no longer
    // decides which timestamp bumps (that's fully data-driven in the Site model now).
    await recordStatusPeriod({ site, previousStatus: beforeStatus, source: resolvedSource, userId: req.user._id });
  }
  await syncBookingTimelineRecords(site, req.user._id, resolvedSource);
  const changedAt = nowIST();
  await logFieldChanges(site._id, before, site.toObject(), req.user._id, changedAt);
  await logBookingChanges(site._id, beforeBookings, site.bookings, req.user._id, changedAt);
  for (const b of cancelledBookings) {
    await logBookingCancellation(site._id, b, req.user._id, changedAt);
  }
  res.json({ ...site.toObject(), mediaCode: site.mediaId });
});

// Individual booking cancellation from `/sites` → Edit Site → Booking Details (one booking at
// a time, identified by bookingId — distinct from changeStatus's whole-site mediaStatus flow).
const cancelBooking = asyncHandler(async (req, res) => {
  const site = await Site.findById(req.params.id);
  if (!site) {
    res.status(404);
    throw new Error('Site not found');
  }
  const booking = (site.bookings || []).find((b) => b.bookingId === req.params.bookingId);
  if (!booking) {
    res.status(404);
    throw new Error('Booking not found');
  }
  if (booking.status === 'cancelled') {
    res.status(400);
    throw new Error('This booking is already cancelled');
  }
  const cancelReason = (req.body.reason || '').trim();
  if (!cancelReason) {
    res.status(400);
    throw new Error('Cancellation reason is required');
  }

  const beforeStatus = site.mediaStatus;
  const beforeBookingId = site.bookingInfo?.bookingId;

  applyBookingCancellation(booking, cancelReason, req.user);
  site.markModified('bookings');
  // Remaining valid bookings (e.g. an Upcoming one) decide the site's status next — cancelling
  // one booking never blindly sets the site to Available.
  resolveSiteStatus(site);
  await site.save();

  const statusChanged = beforeStatus !== site.mediaStatus || beforeBookingId !== site.bookingInfo?.bookingId;
  const resolvedSource = req.body.source === 'inventory' ? 'inventory' : 'sites';
  if (statusChanged) {
    await recordStatusPeriod({ site, previousStatus: beforeStatus, source: resolvedSource, userId: req.user._id });
  }
  await syncBookingTimelineRecords(site, req.user._id, resolvedSource);
  await logBookingCancellation(site._id, booking, req.user._id, nowIST());

  res.json({ ...site.toObject(), mediaCode: site.mediaId });
});

const bulkChangeStatus = asyncHandler(async (req, res) => {
  const { siteIds, mediaStatus, blockReason, blockNotes, bookingInfo } = req.body;
  if (!Array.isArray(siteIds) || siteIds.length === 0) {
    res.status(400);
    throw new Error('No sites selected');
  }
  if (!['available', 'booked', 'blocked'].includes(mediaStatus)) {
    res.status(400);
    throw new Error('Invalid media status');
  }
  if (mediaStatus === 'blocked' && !blockReason) {
    res.status(400);
    throw new Error('Block reason is required');
  }

  const sites = await Site.find({ _id: { $in: siteIds } });
  const results = [];
  const skipped = [];
  for (const site of sites) {
    const before = site.toObject();
    const beforeBookings = before.bookings || [];
    const beforeStatus = before.mediaStatus;
    const beforeBookingId = before.bookingInfo?.bookingId;
    let cancelledBookings = [];

    if (mediaStatus === 'blocked') {
      site.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
      site.mediaStatus = 'blocked';
      site.isActive = false;
    } else {
      // Blocked → Booked/Available: unblock first and end the booking that was running when
      // the site got blocked (nothing is saved if the new booking below is skipped).
      cancelledBookings = unblockSite(site, { wasBlocked: beforeStatus === 'blocked', user: req.user });
      if (mediaStatus === 'booked') {
        try {
          // Bulk has no per-site booking to edit — always add a separate new booking.
          upsertActiveBooking(site, bookingInfo, req.user._id, { mode: 'new' });
        } catch (err) {
          skipped.push({ site: site.mediaId, reason: err.message });
          continue;
        }
      }
      resolveSiteStatus(site);
    }

    await fillMissingCustomerNames(site);
    site.$locals.currentUserName = req.user?.name || 'System';
    await site.save();

    const statusChanged = beforeStatus !== site.mediaStatus || beforeBookingId !== site.bookingInfo?.bookingId;
    if (statusChanged) {
      await recordStatusPeriod({ site, previousStatus: beforeStatus, source: 'inventory', userId: req.user._id });
    }
    await syncBookingTimelineRecords(site, req.user._id, 'inventory');
    const changedAt = nowIST();
    await logFieldChanges(site._id, before, site.toObject(), req.user._id, changedAt);
    await logBookingChanges(site._id, beforeBookings, site.bookings, req.user._id, changedAt);
    for (const b of cancelledBookings) {
      await logBookingCancellation(site._id, b, req.user._id, changedAt);
    }
    results.push(site._id);
  }
  res.json({ updated: results.length, total: siteIds.length, skipped });
});

const getSiteHistory = asyncHandler(async (req, res) => {
  const history = await SiteHistory.find({ site: req.params.id })
    .sort({ changedAt: -1 })
    .populate('changedBy', 'name');
  res.json(history);
});

const AUTO_BLOCK_END_REASON = 'Booking ended — site was blocked';

// A site's full history as a list of EVENTS, newest first (current state on top), in the order
// the user actually made the changes. A booking's history row is updated in place when it's
// cancelled, so it's expanded back into its steps here:
//   • "Booked" at the time the booking was made — always shown, even if it was later cancelled;
//   • "Booking Cancelled" at the cancel time — only for a real (manual) cancellation. A booking
//     that ended automatically because the site was blocked is NOT a separate step; the Booked
//     event just carries `endedEarlyAt` ("Ended early — site was blocked").
const getSiteTimeline = asyncHandler(async (req, res) => {
  const [rows, site] = await Promise.all([
    InventoryHistory.find({ site: req.params.id }).populate('changedBy', 'name').lean(),
    Site.findById(req.params.id).select('bookings.bookingId bookings.createdAt').lean(),
  ]);
  const bookingCreatedAt = new Map((site?.bookings || []).map((b) => [b.bookingId, b.createdAt]));
  // Rows are stored in IST-shifted time (nowIST); an ObjectId's own timestamp is real UTC.
  const insertedAt = (row) => new Date(row._id.getTimestamp().getTime() + IST_OFFSET_MS);

  const events = [];
  for (const row of rows) {
    const isBookingRow = row.status === 'booked' || row.status === 'cancelled';
    if (!isBookingRow) {
      events.push({ ...row, eventKey: `${row._id}`, eventAt: row.changedAt });
      continue;
    }
    const cancel = row.cancellationSnapshot || {};
    const endedByBlock =
      row.status === 'cancelled' && (cancel.cancellationType === 'blocked' || cancel.reason === AUTO_BLOCK_END_REASON);
    const bookedAt = row.bookedAt || bookingCreatedAt.get(row.bookingId) || insertedAt(row);

    events.push({
      ...row,
      status: 'booked',
      eventKey: `${row._id}-booked`,
      eventAt: bookedAt,
      bookingLifecycle: row.status === 'cancelled' ? null : computeBookingLifecycle(row),
      endedEarlyAt: endedByBlock ? cancel.cancelledAt : undefined,
      cancelled: row.status === 'cancelled' && !endedByBlock,
    });
    if (row.status === 'cancelled' && !endedByBlock) {
      events.push({ ...row, eventKey: `${row._id}-cancelled`, eventAt: cancel.cancelledAt || row.changedAt });
    }
  }

  events.sort((a, b) => new Date(b.eventAt) - new Date(a.eventAt) || String(b.eventKey).localeCompare(String(a.eventKey)));
  res.json(events);
});

// Timeline rows are listed by when they were last updated (most recent first). Rows written
// before `updatedAt` existed fall back to `changedAt`.
async function findTimelineByLatestUpdate(filter, { skip = 0, limit } = {}) {
  const pipeline = [
    { $match: filter },
    { $addFields: { _latestUpdateAt: { $ifNull: ['$updatedAt', '$changedAt'] } } },
    { $sort: { _latestUpdateAt: -1, _id: -1 } },
  ];
  if (skip) pipeline.push({ $skip: skip });
  if (limit) pipeline.push({ $limit: limit });
  pipeline.push({ $project: { _latestUpdateAt: 0 } });
  const docs = (await InventoryHistory.aggregate(pipeline)).map((d) => InventoryHistory.hydrate(d));
  return InventoryHistory.populate(docs, { path: 'changedBy', select: 'name' });
}

// Timeline LIST: one row per site (not one per change), showing the site's CURRENT state — the
// same status the Inventory page shows. The row is picked by matching the site's live status:
//   • live Booked  → the history row of the booking that's live today (site.bookingInfo);
//   • otherwise    → the still-open row (effectiveTo = null) of its live status (Available/Blocked).
// Only if no row matches (e.g. the date filter excludes it) does it fall back to the most recent
// non-cancelled row. Picking by "most recently updated" was wrong: blocking a site also touches
// its booking row (isActive → false) a moment later, so the booking row won and the list kept
// saying Booked. Sites are ordered by their latest change of any kind; `changeCount` is how many
// history entries the site has (all shown in its "View" timeline).
function groupedTimelinePipeline(filter) {
  return [
    { $match: filter },
    {
      $lookup: {
        from: Site.collection.name,
        localField: 'site',
        foreignField: '_id',
        as: '_site',
        pipeline: [{ $project: { mediaStatus: 1, 'bookingInfo.bookingId': 1 } }],
      },
    },
    {
      $addFields: {
        _siteStatus: { $arrayElemAt: ['$_site.mediaStatus', 0] },
        _siteBookingId: { $arrayElemAt: ['$_site.bookingInfo.bookingId', 0] },
        _latestUpdateAt: { $ifNull: ['$updatedAt', '$changedAt'] },
        _isCancelled: { $cond: [{ $eq: ['$status', 'cancelled'] }, 1, 0] },
      },
    },
    {
      $addFields: {
        _isCurrent: {
          $cond: [
            {
              $cond: [
                { $eq: ['$_siteStatus', 'booked'] },
                { $and: [{ $eq: ['$status', 'booked'] }, { $eq: ['$bookingId', '$_siteBookingId'] }] },
                { $and: [{ $eq: ['$status', '$_siteStatus'] }, { $eq: [{ $ifNull: ['$effectiveTo', null] }, null] }] },
              ],
            },
            1,
            0,
          ],
        },
      },
    },
    { $sort: { _isCurrent: -1, _isCancelled: 1, _latestUpdateAt: -1, _id: -1 } },
    {
      $group: {
        _id: '$site',
        row: { $first: '$$ROOT' },
        changeCount: { $sum: 1 },
        lastChangedAt: { $max: '$_latestUpdateAt' },
        lastRowId: { $max: '$_id' },
      },
    },
  ];
}

const TIMELINE_HELPER_FIELDS = { _site: 0, _siteStatus: 0, _siteBookingId: 0, _latestUpdateAt: 0, _isCancelled: 0, _isCurrent: 0 };

// `status` (optional) filters by the site's CURRENT status — so the "Booked Sites" card, the
// status dropdown and the list all agree.
async function findTimelineGroupedBySite(filter, { status, skip = 0, limit } = {}) {
  const page = [{ $sort: { lastChangedAt: -1, lastRowId: -1 } }];
  if (skip) page.push({ $skip: skip });
  if (limit) page.push({ $limit: limit });
  page.push(
    { $replaceRoot: { newRoot: { $mergeObjects: ['$row', { changeCount: '$changeCount', lastChangedAt: '$lastChangedAt' }] } } },
    { $project: TIMELINE_HELPER_FIELDS }
  );
  const [result] = await InventoryHistory.aggregate([
    ...groupedTimelinePipeline(filter),
    ...(status ? [{ $match: { 'row.status': status } }] : []),
    { $facet: { items: page, total: [{ $count: 'n' }] } },
  ]);
  const items = await InventoryHistory.populate(result.items, { path: 'changedBy', select: 'name' });
  return { items, total: result.total[0]?.n || 0 };
}

const getTimeline = asyncHandler(async (req, res) => {
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  // Status is applied AFTER grouping (current status), not to individual history rows.
  const filter = buildOverlapFilter({ ...req.query, mediaStatus: undefined });

  const { items: rows, total } = await findTimelineGroupedBySite(filter, {
    status: req.query.mediaStatus || undefined,
    skip: (page - 1) * limit,
    limit,
  });
  const items = rows.map((t) => ({ ...t, bookingLifecycle: computeBookingLifecycle(t) }));

  res.json({ items, total, distinctSiteCount: total, page, pages: Math.ceil(total / limit) });
});

// Cards count each site ONCE, by the same current status the list shows — so
// Available + Booked + Blocked always adds up to Total.
const getTimelineSummary = asyncHandler(async (req, res) => {
  const baseFilter = buildOverlapFilter({ ...req.query, mediaStatus: undefined });
  const counts = await InventoryHistory.aggregate([
    ...groupedTimelinePipeline(baseFilter),
    { $group: { _id: '$row.status', n: { $sum: 1 } } },
  ]);
  const by = Object.fromEntries(counts.map((c) => [c._id, c.n]));
  const total = counts.reduce((sum, c) => sum + c.n, 0);

  res.json({ total, available: by.available || 0, booked: by.booked || 0, blocked: by.blocked || 0 });
});

const bulkImport = asyncHandler(async (req, res) => {
  const { records } = req.body;
  if (!Array.isArray(records) || records.length === 0) {
    res.status(400);
    throw new Error('No records to import');
  }
  // Latitude/Longitude are optional: blank, placeholder ("-", "NA") or out-of-range values are
  // dropped so the site still imports without coordinates, instead of the whole row being
  // rejected by the schema's min/max validators (insertMany would skip it silently).
  const optionalCoord = (value, limit) => {
    if (value === undefined || value === null) return undefined;
    const n = Number(String(value).trim().replace(/°/g, ''));
    return String(value).trim() !== '' && Number.isFinite(n) && Math.abs(n) <= limit ? n : undefined;
  };
  const docs = records.map((r) => {
    const doc = {
      ...normalizeSiteBody(r),
      mediaStatus: r.mediaStatus || 'available',
      createdBy: req.user._id,
      updatedBy: 'System',
      inventoryUpdatedBy: 'System',
    };
    doc.latitude = optionalCoord(doc.latitude, 90);
    doc.longitude = optionalCoord(doc.longitude, 180);
    if (doc.latitude === undefined) delete doc.latitude;
    if (doc.longitude === undefined) delete doc.longitude;
    Site.applyComputedFields(doc);
    return doc;
  });
  const created = await Site.insertMany(docs, { ordered: false });
  res.status(201).json({ imported: created.length });
});

const getSiteOwners = asyncHandler(async (req, res) => {
  const owners = await Site.distinct('siteOwner', { siteOwner: { $nin: [null, ''] } });
  res.json(owners.sort((a, b) => a.localeCompare(b)));
});

const getSummary = asyncHandler(async (req, res) => {
  const filter = buildFilter(req.query);
  const [total, available, booked, blocked] = await Promise.all([
    Site.countDocuments(filter),
    Site.countDocuments({ ...filter, mediaStatus: 'available' }),
    Site.countDocuments({ ...filter, mediaStatus: 'booked' }),
    Site.countDocuments({ ...filter, mediaStatus: 'blocked' }),
  ]);
  res.json({ total, available, booked, blocked });
});

const SITE_EXPORT_COLUMNS = [
  { header: 'MediaCode', key: 'mediaCode', width: 16 },
  { header: 'Media Type', key: 'mediaType', width: 16 },
  { header: 'Quantity', key: 'quantity', width: 10, numeric: true },
  { header: 'State', key: 'state', width: 16 },
  { header: 'City', key: 'city', width: 14 },
  { header: 'Location', key: 'location', width: 24, wrap: true },
  { header: 'Area Name', key: 'areaName', width: 16 },
  { header: 'Site Owner', key: 'siteOwner', width: 18 },
  { header: 'Latitude', key: 'latitude', width: 12, numeric: true },
  { header: 'Longitude', key: 'longitude', width: 12, numeric: true },
  { header: 'Illumination', key: 'illumination', width: 14 },
  { header: 'Width', key: 'width', width: 9, numeric: true },
  { header: 'Height', key: 'height', width: 9, numeric: true },
  { header: 'Auto Size', key: 'autoSize', width: 11, numeric: true },
  { header: 'Display Cost Per Month', key: 'monthlyAmount', width: 20, numeric: true },
  { header: 'Printing Cost', key: 'printingCost', width: 14, numeric: true },
  { header: 'Mounting Cost', key: 'mountingCost', width: 14, numeric: true },
  { header: 'Total Cost', key: 'totalCost', width: 14, numeric: true },
  { header: 'Media Status', key: 'mediaStatus', width: 14 },
  { header: 'Active Status', key: 'activeStatus', width: 14 },
  { header: 'MediaImage', key: 'mediaImage', width: 40, wrap: true },
  { header: 'Inventory Updated On', key: 'inventoryUpdatedOn', width: 22 },
  { header: 'Last Updated On', key: 'lastUpdatedOn', width: 22 },
];

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
const THIN_BORDER = { style: 'thin', color: { argb: 'FFB8C2CC' } };
const ALL_BORDERS = { top: THIN_BORDER, left: THIN_BORDER, bottom: THIN_BORDER, right: THIN_BORDER };

const exportSites = asyncHandler(async (req, res) => {
  const filter = buildFilter(req.query);
  const sites = await Site.find(filter).sort(siteListSort(req.query));

  const now = nowIST();
  const pad = (n) => String(n).padStart(2, '0');
  const dateStr = `${now.getUTCFullYear()}-${pad(now.getUTCMonth() + 1)}-${pad(now.getUTCDate())}`;
  const timeStr = `${pad(now.getUTCHours())}-${pad(now.getUTCMinutes())}`;

  const activeStatusFilterLabel =
    req.query.isActive === '' || req.query.isActive === undefined ? 'All' : req.query.isActive === 'true' ? 'Active' : 'Inactive';

  const wb = new ExcelJS.Workbook();

  // --- Summary sheet ---
  const summarySheet = wb.addWorksheet('Summary');
  summarySheet.columns = [{ width: 22 }, { width: 32 }];
  const summaryRows = [
    ['Generated On', formatIST(now)],
    ['State Filter', req.query.state || 'All'],
    ['City Filter', req.query.city || 'All'],
    ['Media Status Filter', req.query.mediaStatus || 'All'],
    ['Active Status Filter', activeStatusFilterLabel],
    ['Search Filter', req.query.search || 'None'],
  ];
  summaryRows.forEach(([label, value]) => {
    const row = summarySheet.addRow([label, value]);
    const labelCell = row.getCell(1);
    const valueCell = row.getCell(2);
    labelCell.font = { bold: true, color: { argb: 'FF1F2937' } };
    labelCell.fill = HEADER_FILL;
    labelCell.border = ALL_BORDERS;
    valueCell.font = { color: { argb: 'FF1F2937' } };
    valueCell.border = ALL_BORDERS;
    valueCell.alignment = { vertical: 'middle' };
  });

  // --- Sites sheet ---
  const dataSheet = wb.addWorksheet('Sites');
  dataSheet.columns = SITE_EXPORT_COLUMNS.map((c) => ({ header: c.header, key: c.key, width: c.width }));

  const headerRow = dataSheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FF1F2937' } };
    cell.fill = HEADER_FILL;
    cell.border = ALL_BORDERS;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  headerRow.height = 20;
  dataSheet.views = [{ state: 'frozen', ySplit: 1 }];

  sites.forEach((s) => {
    const row = dataSheet.addRow({
      mediaCode: s.mediaId,
      mediaType: s.mediaType,
      quantity: s.quantity,
      state: s.state,
      city: s.city,
      location: s.location,
      areaName: s.areaName,
      siteOwner: s.siteOwner,
      latitude: s.latitude,
      longitude: s.longitude,
      illumination: s.illumination,
      width: s.width,
      height: s.height,
      autoSize: s.autoSize,
      monthlyAmount: s.monthlyAmount,
      printingCost: s.printingCost,
      mountingCost: s.mountingCost,
      totalCost: s.totalCost,
      mediaStatus: s.mediaStatus,
      activeStatus: s.isActive ? 'Active' : 'Inactive',
      mediaImage: s.mediaImage || 'N/A',
      inventoryUpdatedOn: formatIST(s.inventoryUpdatedAt),
      lastUpdatedOn: formatIST(s.updatedAt),
    });
    row.eachCell((cell, colNumber) => {
      const col = SITE_EXPORT_COLUMNS[colNumber - 1];
      cell.border = ALL_BORDERS;
      if (col?.numeric) cell.alignment = { horizontal: 'right', vertical: 'middle' };
      else if (col?.wrap) cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      else cell.alignment = { horizontal: 'left', vertical: 'middle' };
    });
  });

  const stateSlug = req.query.state ? req.query.state.replace(/\s+/g, '') : null;
  const citySlug = req.query.city ? req.query.city.replace(/\s+/g, '') : null;
  const statusSlug = req.query.mediaStatus ? req.query.mediaStatus.replace(/\s+/g, '') : null;
  const prefix = req.query.filenamePrefix === 'inventory' ? 'inventory' : 'sites';
  const namePart = stateSlug || citySlug || statusSlug ? [stateSlug, citySlug, statusSlug].filter(Boolean).join('_') : 'all';
  const filename = `${prefix}_${namePart}_${dateStr}_${timeStr}.xlsx`;

  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(Buffer.from(buffer));
});

const TIMELINE_EXPORT_COLUMNS = [
  { header: 'S.No', key: 'sno', width: 8, numeric: true },
  { header: 'MediaCode', key: 'mediaCode', width: 16 },
  { header: 'Media Image', key: 'mediaImage', width: 40, wrap: true },
  { header: 'Media Type', key: 'mediaType', width: 16 },
  { header: 'State', key: 'state', width: 16 },
  { header: 'City', key: 'city', width: 14 },
  { header: 'Site Owner', key: 'siteOwner', width: 18 },
  { header: 'Status', key: 'status', width: 12 },
  { header: 'Effective From', key: 'effectiveFrom', width: 20 },
  { header: 'Effective To', key: 'effectiveTo', width: 20 },
  { header: 'Duration', key: 'duration', width: 12 },
  { header: 'Customer Type', key: 'customerType', width: 14 },
  { header: 'Customer Name', key: 'customerName', width: 20 },
  { header: 'Monthly Cost', key: 'monthlyCost', width: 14, numeric: true },
  { header: 'Booking Amount', key: 'bookingAmount', width: 14, numeric: true },
  { header: 'Block Reason', key: 'blockReason', width: 20, wrap: true },
  { header: 'Block Notes', key: 'blockNotes', width: 24, wrap: true },
  { header: 'Changed On', key: 'changedOn', width: 22 },
  { header: 'Changed By', key: 'changedBy', width: 18 },
  { header: 'Source', key: 'source', width: 12 },
];

const exportTimeline = asyncHandler(async (req, res) => {
  const filter = buildOverlapFilter(req.query);
  const [records, distinctSites] = await Promise.all([
    findTimelineByLatestUpdate(filter),
    InventoryHistory.distinct('site', filter),
  ]);

  const now = nowIST();
  const pad = (n) => String(n).padStart(2, '0');
  const dateStr = `${now.getUTCFullYear()}-${pad(now.getUTCMonth() + 1)}-${pad(now.getUTCDate())}`;
  const timeStr = `${pad(now.getUTCHours())}-${pad(now.getUTCMinutes())}`;

  const activeStatusFilterLabel =
    req.query.isActive === '' || req.query.isActive === undefined ? 'All' : req.query.isActive === 'true' ? 'Active' : 'Inactive';

  const wb = new ExcelJS.Workbook();

  const summarySheet = wb.addWorksheet('Summary');
  summarySheet.columns = [{ width: 24 }, { width: 32 }];
  const summaryRows = [
    ['Generated On', formatIST(now)],
    ['From Date', req.query.from ? formatIST(req.query.from) : 'All'],
    ['To Date', req.query.to ? formatIST(req.query.to) : 'All'],
    ['State Filter', req.query.state || 'All'],
    ['City Filter', req.query.city || 'All'],
    ['Status Filter', req.query.mediaStatus || 'All'],
    ['Active Status Filter', activeStatusFilterLabel],
    ['Search Filter', req.query.search || 'None'],
    ['Total Matching History Records', records.length],
    ['Distinct Sites', distinctSites.length],
  ];
  summaryRows.forEach(([label, value]) => {
    const row = summarySheet.addRow([label, value]);
    row.getCell(1).font = { bold: true, color: { argb: 'FF1F2937' } };
    row.getCell(1).fill = HEADER_FILL;
    row.getCell(1).border = ALL_BORDERS;
    row.getCell(2).border = ALL_BORDERS;
    row.getCell(2).alignment = { vertical: 'middle' };
  });

  const dataSheet = wb.addWorksheet('History');
  dataSheet.columns = TIMELINE_EXPORT_COLUMNS.map((c) => ({ header: c.header, key: c.key, width: c.width }));
  const headerRow = dataSheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FF1F2937' } };
    cell.fill = HEADER_FILL;
    cell.border = ALL_BORDERS;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  headerRow.height = 20;
  dataSheet.views = [{ state: 'frozen', ySplit: 1 }];

  records.forEach((h, idx) => {
    const durationDays = h.status === 'booked' ? h.bookingSnapshot?.durationDays : null;
    const row = dataSheet.addRow({
      sno: idx + 1,
      mediaCode: h.mediaId,
      mediaImage: h.mediaImage || 'N/A',
      mediaType: h.mediaType,
      state: h.state,
      city: h.city,
      siteOwner: h.siteOwner || '-',
      status: h.status,
      effectiveFrom: formatIST(h.effectiveFrom),
      effectiveTo: h.effectiveTo ? formatIST(h.effectiveTo) : 'Ongoing',
      duration: durationDays ? `${durationDays} Days` : '-',
      customerType: h.bookingSnapshot?.customerType || '-',
      customerName: h.bookingSnapshot?.customerName || '-',
      monthlyCost: h.bookingSnapshot?.monthlyTotalCost ?? '-',
      bookingAmount: h.bookingSnapshot?.amount ?? '-',
      blockReason: h.blockSnapshot?.reason || '-',
      blockNotes: h.blockSnapshot?.notes || '-',
      changedOn: formatIST(h.changedAt),
      changedBy: h.changedBy?.name || '-',
      source: h.source,
    });
    row.eachCell((cell, colNumber) => {
      const col = TIMELINE_EXPORT_COLUMNS[colNumber - 1];
      cell.border = ALL_BORDERS;
      if (col?.numeric) cell.alignment = { horizontal: 'right', vertical: 'middle' };
      else if (col?.wrap) cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      else cell.alignment = { horizontal: 'left', vertical: 'middle' };
    });
  });

  const stateSlug = req.query.state ? req.query.state.replace(/\s+/g, '') : null;
  const citySlug = req.query.city ? req.query.city.replace(/\s+/g, '') : null;
  const statusSlug = req.query.mediaStatus ? req.query.mediaStatus.replace(/\s+/g, '') : null;
  const fromSlug = req.query.from ? formatIST(req.query.from).replace(/,.*/, '').replace(/\//g, '-') : null;
  const toSlug = req.query.to ? formatIST(req.query.to).replace(/,.*/, '').replace(/\//g, '-') : null;
  const locationPart = [stateSlug, citySlug, statusSlug].filter(Boolean).join('_');
  const rangePart = fromSlug && toSlug ? `${fromSlug}_to_${toSlug}` : 'all';
  const filename = `inventory_timeline_${locationPart ? locationPart + '_' : ''}${rangePart}.xlsx`;

  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(Buffer.from(buffer));
});

const uploadImage = asyncHandler(async (req, res) => {
  const file = extractUploadedFile(req);
  if (!file) {
    res.status(400);
    throw new Error('No image file provided');
  }
  const url = await saveMediaImage(file);
  res.json({ url });
});

const STATIC_STATES = ['Tamil Nadu', 'Kerala', 'Karnataka'];

function escapeRegex(text) {
  return text.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
}

const getStates = asyncHandler(async (req, res) => {
  res.json(STATIC_STATES);
});

const getCities = asyncHandler(async (req, res) => {
  const state = req.params.state || req.query.state;
  if (!state || !String(state).trim()) {
    return res.json([]);
  }

  const rawCities = await Site.distinct('city', {
    state: new RegExp(`^${escapeRegex(String(state).trim())}$`, 'i'),
  });

  const cleanCities = rawCities
    .filter((c) => c && String(c).trim().length > 0)
    .map((c) => String(c).trim())
    .filter((v, i, a) => a.indexOf(v) === i)
    .sort((a, b) => a.localeCompare(b));

  res.json(cleanCities);
});

module.exports = {
  getSites,
  getAvailableSites,
  getSite,
  createSite,
  updateSite,
  deleteSite,
  changeStatus,
  cancelBooking,
  bulkChangeStatus,
  bulkImport,
  uploadImage,
  exportSites,
  getSiteHistory,
  getSummary,
  getSiteOwners,
  getSiteTimeline,
  getTimeline,
  getTimelineSummary,
  exportTimeline,
  getStates,
  getCities,
  STATIC_STATES,
};
