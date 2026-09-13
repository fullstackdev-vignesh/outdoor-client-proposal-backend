const asyncHandler = require('express-async-handler');
const XLSX = require('xlsx');
const Site = require('../models/Site');
const SiteHistory = require('../models/SiteHistory');
const { saveMediaImage } = require('../utils/imageStorage');
const { calcDurationDays, calcBookingAmount } = require('../utils/bookingCalc');
const { formatIST } = require('../utils/formatDate');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const TRACKED_FIELDS = [
  'mediaId', 'mediaType', 'quantity', 'state', 'city', 'location', 'areaName', 'locationDetails',
  'latitude', 'longitude', 'illumination', 'width', 'height', 'sizeUnit', 'amount', 'gstAmount',
  'monthlyAmount', 'printingCost', 'mountingCost', 'totalCost', 'image', 'isActive', 'mediaStatus',
];

async function logFieldChanges(siteId, before, after, userId) {
  const entries = [];
  for (const field of TRACKED_FIELDS) {
    const oldVal = before ? before[field] : undefined;
    const newVal = after ? after[field] : undefined;
    const oldStr = oldVal === undefined || oldVal === null ? '' : String(oldVal);
    const newStr = newVal === undefined || newVal === null ? '' : String(newVal);
    if (oldStr !== newStr) {
      entries.push({ site: siteId, field, oldValue: oldVal, newValue: newVal, changedBy: userId, changedAt: nowIST() });
    }
  }
  if (entries.length) await SiteHistory.insertMany(entries);
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
      .sort({ createdAt: -1 })
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
    Site.find(filter).sort({ createdAt: -1 }).skip((page - 1) * limit).limit(limit),
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
  return payload;
}

const createSite = asyncHandler(async (req, res) => {
  const payload = normalizeSiteBody(req.body);
  const errors = validateSitePayload(payload);
  if (errors.length) {
    res.status(400);
    throw new Error(errors.join('; '));
  }
  payload.createdBy = req.user._id;
  if (!payload.mediaStatus) payload.mediaStatus = 'available';
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
  const before = site.toObject();
  Object.assign(site, payload);
  try {
    await site.save();
  } catch (err) {
    if (err.code === 11000 && err.keyPattern?.mediaId) {
      res.status(400);
      throw new Error('MediaCode already exists. Please use a different MediaCode.');
    }
    throw err;
  }
  await logFieldChanges(site._id, before, site.toObject(), req.user._id);
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

function buildBookingInfo(site, bookingInfo, userId) {
  const { customerType, client, startDate, endDate } = bookingInfo || {};
  if (!client) throw new Error('Customer is required');
  if (!startDate || !endDate) throw new Error('Start Date and End Date are required');
  if (new Date(endDate) < new Date(startDate)) throw new Error('End Date must be on or after Start Date');

  const durationDays = calcDurationDays(startDate, endDate);
  const monthlyTotalCost = site.totalCost || site.monthlyAmount || 0;
  const amount = calcBookingAmount(monthlyTotalCost, durationDays);

  return {
    customerType: customerType === 'agency' ? 'agency' : 'client',
    client,
    startDate,
    endDate,
    durationDays,
    monthlyTotalCost,
    amount,
    bookedBy: userId,
  };
}

const changeStatus = asyncHandler(async (req, res) => {
  const site = await Site.findById(req.params.id);
  if (!site) {
    res.status(404);
    throw new Error('Site not found');
  }
  const { mediaStatus, blockReason, blockNotes, bookingInfo } = req.body;
  const before = site.toObject();

  if (!['available', 'booked', 'blocked'].includes(mediaStatus)) {
    res.status(400);
    throw new Error('Invalid media status');
  }

  if (mediaStatus === 'blocked') {
    if (!blockReason) {
      res.status(400);
      throw new Error('Block reason is required');
    }
    site.blockInfo = {
      reason: blockReason,
      notes: blockNotes,
      blockedDate: nowIST(),
      blockedBy: req.user._id,
    };
  } else if (mediaStatus === 'booked') {
    try {
      site.bookingInfo = buildBookingInfo(site, bookingInfo, req.user._id);
    } catch (err) {
      res.status(400);
      throw err;
    }
  }

  site.mediaStatus = mediaStatus;
  site.isActive = mediaStatus !== 'blocked';
  site.$locals.inventoryOnly = true;
  await site.save();
  await logFieldChanges(site._id, before, site.toObject(), req.user._id);
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
  for (const site of sites) {
    const before = site.toObject();
    if (mediaStatus === 'blocked') {
      site.blockInfo = { reason: blockReason, notes: blockNotes, blockedDate: nowIST(), blockedBy: req.user._id };
    } else if (mediaStatus === 'booked') {
      try {
        site.bookingInfo = buildBookingInfo(site, bookingInfo, req.user._id);
      } catch (err) {
        continue;
      }
    }
    site.mediaStatus = mediaStatus;
    site.isActive = mediaStatus !== 'blocked';
    site.$locals.inventoryOnly = true;
    await site.save();
    await logFieldChanges(site._id, before, site.toObject(), req.user._id);
    results.push(site._id);
  }
  res.json({ updated: results.length, total: siteIds.length });
});

const getSiteHistory = asyncHandler(async (req, res) => {
  const history = await SiteHistory.find({ site: req.params.id })
    .sort({ changedAt: -1 })
    .populate('changedBy', 'name');
  res.json(history);
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

const exportSites = asyncHandler(async (req, res) => {
  const filter = buildFilter(req.query);
  const sites = await Site.find(filter).sort({ createdAt: -1 });

  const rows = sites.map((s) => ({
    MediaCode: s.mediaId,
    'Media Type': s.mediaType,
    Quantity: s.quantity,
    State: s.state,
    City: s.city,
    Location: s.location,
    'Area Name': s.areaName,
    Latitude: s.latitude,
    Longitude: s.longitude,
    Illumination: s.illumination,
    Width: s.width,
    Height: s.height,
    'Auto Size': s.autoSize,
    'Display Cost Per Month': s.monthlyAmount,
    'Printing Cost': s.printingCost,
    'Mounting Cost': s.mountingCost,
    'Total Cost': s.totalCost,
    'Media Status': s.mediaStatus,
    'Active Status': s.isActive ? 'Active' : 'Inactive',
    'Inventory Updated On': formatIST(s.inventoryUpdatedAt),
    'Last Updated On': formatIST(s.updatedAt),
  }));

  const now = nowIST();
  const pad = (n) => String(n).padStart(2, '0');
  const dateStr = `${now.getUTCFullYear()}-${pad(now.getUTCMonth() + 1)}-${pad(now.getUTCDate())}`;
  const timeStr = `${pad(now.getUTCHours())}-${pad(now.getUTCMinutes())}`;

  const metaRows = [
    ['Generated On', formatIST(now)],
    ['State Filter', req.query.state || 'All'],
    ['City Filter', req.query.city || 'All'],
    ['Media Status Filter', req.query.mediaStatus || 'All'],
    ['Active Status Filter', req.query.isActive === '' || req.query.isActive === undefined ? 'All' : req.query.isActive === 'true' ? 'Active' : 'Inactive'],
    ['Search Filter', req.query.search || 'None'],
    [],
  ];

  const wb = XLSX.utils.book_new();
  const metaSheet = XLSX.utils.aoa_to_sheet(metaRows);
  metaSheet['!cols'] = [{ wch: 20 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(wb, metaSheet, 'Summary');
  const dataSheet = XLSX.utils.json_to_sheet(rows);
  const headers = rows.length ? Object.keys(rows[0]) : [];
  dataSheet['!cols'] = headers.map((h) => ({ wch: Math.max(h.length + 4, h.includes('On') ? 20 : 12) }));
  XLSX.utils.book_append_sheet(wb, dataSheet, 'Sites');

  const stateSlug = req.query.state ? req.query.state.replace(/\s+/g, '') : null;
  const citySlug = req.query.city ? req.query.city.replace(/\s+/g, '') : null;
  const statusSlug = req.query.mediaStatus ? req.query.mediaStatus.replace(/\s+/g, '') : null;
  const prefix = req.query.filenamePrefix === 'inventory' ? 'inventory' : 'sites';
  const namePart = stateSlug || citySlug || statusSlug ? [stateSlug, citySlug, statusSlug].filter(Boolean).join('_') : 'all';
  const filename = `${prefix}_${namePart}_${dateStr}_${timeStr}.xlsx`;

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
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
  bulkChangeStatus,
  bulkImport,
  uploadImage,
  exportSites,
  getSiteHistory,
  getSummary,
};
