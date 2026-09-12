const asyncHandler = require('express-async-handler');
const Site = require('../models/Site');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const buildFilter = (query) => {
  const filter = {};
  if (query.search) {
    filter.$or = [
      { mediaId: new RegExp(query.search, 'i') },
      { mediaName: new RegExp(query.search, 'i') },
      { mediaType: new RegExp(query.search, 'i') },
      { city: new RegExp(query.search, 'i') },
      { state: new RegExp(query.search, 'i') },
      { location: new RegExp(query.search, 'i') },
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

  res.json({ items, total, page, pages: Math.ceil(total / limit) });
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
  res.json(site);
});

const createSite = asyncHandler(async (req, res) => {
  const payload = { ...req.body, createdBy: req.user._id };
  if (!payload.mediaStatus) payload.mediaStatus = 'available';
  const site = await Site.create(payload);
  res.status(201).json(site);
});

const updateSite = asyncHandler(async (req, res) => {
  const site = await Site.findById(req.params.id);
  if (!site) {
    res.status(404);
    throw new Error('Site not found');
  }
  Object.assign(site, req.body);
  await site.save();
  res.json(site);
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
  const { mediaStatus, blockReason, blockNotes, bookingInfo } = req.body;

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
    site.bookingInfo = undefined;
  } else if (mediaStatus === 'booked') {
    if (!bookingInfo || !bookingInfo.client) {
      res.status(400);
      throw new Error('Booking information with client is required');
    }
    site.bookingInfo = { ...bookingInfo, bookedBy: req.user._id };
    site.blockInfo = undefined;
  } else {
    site.bookingInfo = undefined;
    site.blockInfo = undefined;
  }

  site.mediaStatus = mediaStatus;
  await site.save();
  res.json(site);
});

const bulkImport = asyncHandler(async (req, res) => {
  const { records } = req.body;
  if (!Array.isArray(records) || records.length === 0) {
    res.status(400);
    throw new Error('No records to import');
  }
  const docs = records.map((r) => ({ ...r, mediaStatus: r.mediaStatus || 'available', createdBy: req.user._id }));
  const created = await Site.insertMany(docs, { ordered: false });
  res.status(201).json({ imported: created.length });
});

module.exports = {
  getSites,
  getAvailableSites,
  getSite,
  createSite,
  updateSite,
  deleteSite,
  changeStatus,
  bulkImport,
};
