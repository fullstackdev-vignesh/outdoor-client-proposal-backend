const asyncHandler = require('express-async-handler');
const ExcelJS = require('exceljs');
const Site = require('../models/Site');
const SiteHistory = require('../models/SiteHistory');
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
function applyBookingCancellation(booking, reason, user) {
  booking.status = 'cancelled';
  booking.cancellationReason = reason;
  booking.cancelledAt = nowIST();
  booking.cancelledBy = user._id;
  booking.cancelledByName = user.name;
  booking.cancelledByRole = user.role;
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
  if (query.state) filter.state = query.state;
  if (query.city) filter.city = query.city;
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

const getSites = asyncHandler(async (req, res) => {
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  const filter = buildFilter(req.query);

  const [items, total] = await Promise.all([
    Site.find(filter)
      .sort({ updatedAt: -1, _id: 1 })
      .skip((page - 1) * limit)
      .limit(limit),
    Site.countDocuments(filter),
  ]);

  res.json({ items: items.map((s) => ({ ...s.toObject(), mediaCode: s.mediaId })), total, page, pages: Math.ceil(total / limit) });
});

const getAvailableSites = asyncHandler(async (req, res) => {
  const filter = { ...buildFilter(req.query), mediaStatus: 'available', isActive: true };
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);

  const [items, total] = await Promise.all([
    Site.find(filter).sort({ updatedAt: -1, _id: 1 }).skip((page - 1) * limit).limit(limit),
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
  ['amount', 'gstAmount', 'monthlyAmount', 'printingCost', 'mountingCost'].forEach((f) => {
    if (body[f] !== undefined && (isNaN(num(body[f])) || num(body[f]) < 0)) errors.push(`${f} must be a number greater than or equal to 0`);
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
  // Optional Site Info dropdown submits '' when left unselected — an empty string fails
  // ObjectId casting, so treat it the same as "not provided".
  if (payload.siteInfoId === '') delete payload.siteInfoId;
  return payload;
}

// Uploads the newly selected image (if any) via the existing storage logic and sets
// payload.mediaImage to its public URL. If no new file was sent, mediaImage is left
// untouched so an update never clears/overwrites the site's existing image.
async function applyUploadedImage(payload, req) {
  if (req.file) {
    payload.mediaImage = await saveMediaImage(req.file);
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
function upsertActiveBooking(site, input, userId) {
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
  payload.createdBy = req.user._id;
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
  Object.assign(site, siteFields);
  Site.applyComputedFields(site);

  if (site.mediaStatus === 'blocked') {
    if (!blockReason) {
      res.status(400);
      throw new Error('Block reason is required');
    }
    site.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
    site.isActive = false;
  } else {
    const bookingList = Array.isArray(incomingBookings) ? incomingBookings : bookingInfo ? [bookingInfo] : null;
    if (site.mediaStatus === 'booked' && bookingList) {
      try {
        site.bookings = buildBookingsArray(site, bookingList, req.user._id);
      } catch (err) {
        res.status(400);
        throw err;
      }
    }
    // Live status/bookingInfo always get recomputed from the (possibly just-edited)
    // bookings array against today's date — never taken at face value from the toggle.
    resolveSiteStatus(site);
  }

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
  const { mediaStatus, blockReason, blockNotes, bookingInfo, source, cancellationReason, reason } = req.body;
  const before = site.toObject();
  const beforeBookings = before.bookings || [];
  const beforeStatus = before.mediaStatus;
  const beforeBookingId = before.bookingInfo?.bookingId;

  if (!['available', 'booked', 'blocked'].includes(mediaStatus)) {
    res.status(400);
    throw new Error('Invalid media status');
  }

  let cancelledBooking = null;

  if (mediaStatus === 'blocked') {
    if (!blockReason) {
      res.status(400);
      throw new Error('Block reason is required');
    }
    site.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
    site.mediaStatus = 'blocked';
    site.isActive = false;
  } else {
    if (mediaStatus === 'booked') {
      try {
        upsertActiveBooking(site, bookingInfo, req.user._id);
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
        cancelledBooking = activeBooking;
      }
    }
    // Live status/bookingInfo are always recomputed from the (remaining valid) bookings array
    // against today's date, not taken at face value from the requested `mediaStatus` — so an
    // Upcoming booking elsewhere in the array keeps the site Booked even after this cancellation.
    resolveSiteStatus(site);
  }

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
  if (cancelledBooking) {
    await logBookingCancellation(site._id, cancelledBooking, req.user._id, changedAt);
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

    if (mediaStatus === 'blocked') {
      site.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
      site.mediaStatus = 'blocked';
      site.isActive = false;
    } else {
      if (mediaStatus === 'booked') {
        try {
          upsertActiveBooking(site, bookingInfo, req.user._id);
        } catch (err) {
          skipped.push({ site: site.mediaId, reason: err.message });
          continue;
        }
      }
      resolveSiteStatus(site);
    }

    await site.save();

    const statusChanged = beforeStatus !== site.mediaStatus || beforeBookingId !== site.bookingInfo?.bookingId;
    if (statusChanged) {
      await recordStatusPeriod({ site, previousStatus: beforeStatus, source: 'inventory', userId: req.user._id });
    }
    await syncBookingTimelineRecords(site, req.user._id, 'inventory');
    const changedAt = nowIST();
    await logFieldChanges(site._id, before, site.toObject(), req.user._id, changedAt);
    await logBookingChanges(site._id, beforeBookings, site.bookings, req.user._id, changedAt);
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

const getSiteTimeline = asyncHandler(async (req, res) => {
  const timeline = await InventoryHistory.find({ site: req.params.id })
    .sort({ effectiveFrom: 1 })
    .populate('changedBy', 'name');
  res.json(timeline.map((t) => ({ ...t.toObject(), bookingLifecycle: computeBookingLifecycle(t) })));
});

const getTimeline = asyncHandler(async (req, res) => {
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  const filter = buildOverlapFilter(req.query);

  const [rows, total, distinctSites] = await Promise.all([
    InventoryHistory.find(filter)
      .sort({ effectiveFrom: -1 })
      .skip((page - 1) * limit)
      .limit(limit)
      .populate('changedBy', 'name'),
    InventoryHistory.countDocuments(filter),
    InventoryHistory.distinct('site', filter),
  ]);
  const items = rows.map((t) => ({ ...t.toObject(), bookingLifecycle: computeBookingLifecycle(t) }));

  res.json({ items, total, distinctSiteCount: distinctSites.length, page, pages: Math.ceil(total / limit) });
});

const getTimelineSummary = asyncHandler(async (req, res) => {
  const baseFilter = buildOverlapFilter({ ...req.query, mediaStatus: undefined });

  const distinctByStatus = async (status) => {
    const ids = await InventoryHistory.distinct('site', { ...baseFilter, status });
    return ids.length;
  };

  const [totalIds, available, booked, blocked] = await Promise.all([
    InventoryHistory.distinct('site', baseFilter),
    distinctByStatus('available'),
    distinctByStatus('booked'),
    distinctByStatus('blocked'),
  ]);

  res.json({ total: totalIds.length, available, booked, blocked });
});

const bulkImport = asyncHandler(async (req, res) => {
  const { records } = req.body;
  if (!Array.isArray(records) || records.length === 0) {
    res.status(400);
    throw new Error('No records to import');
  }
  const docs = records.map((r) => {
    const doc = {
      ...normalizeSiteBody(r),
      mediaStatus: r.mediaStatus || 'available',
      createdBy: req.user._id,
    };
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
  const sites = await Site.find(filter).sort({ createdAt: -1 });

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
    InventoryHistory.find(filter).sort({ effectiveFrom: -1 }).populate('changedBy', 'name'),
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
  if (!req.file) {
    res.status(400);
    throw new Error('No image file provided');
  }
  const url = await saveMediaImage(req.file);
  res.json({ url });
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
};
