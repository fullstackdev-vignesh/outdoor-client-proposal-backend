const asyncHandler = require('express-async-handler');
const MediaRate = require('../models/MediaRate');
const { MOUNTING_TYPES } = require('../models/MediaRate');

const escapeRegex = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// Validates and normalises the body. Returns { data } or { error }.
function readBody(body) {
  const mediaType = String(body.mediaType || '').trim();
  if (!mediaType) return { error: 'Media Type is required' };

  const num = (v) => (v === '' || v === null || v === undefined ? NaN : Number(v));
  const printingBackLit = num(body.printingBackLit);
  const printingOther = num(body.printingOther);
  const mountingAmount = num(body.mountingAmount);
  if (!Number.isFinite(printingBackLit) || printingBackLit < 0) return { error: 'Back Lit printing rate must be a number of 0 or more' };
  if (!Number.isFinite(printingOther) || printingOther < 0) return { error: 'Front / Non Lit printing rate must be a number of 0 or more' };
  if (!MOUNTING_TYPES.includes(body.mountingType)) return { error: 'Mounting type is invalid' };
  if (!Number.isFinite(mountingAmount) || mountingAmount < 0) return { error: 'Mounting amount must be a number of 0 or more' };

  return { data: { mediaType, printingBackLit, printingOther, mountingType: body.mountingType, mountingAmount } };
}

// Media types are unique regardless of case ("Hoarding" vs "hoarding").
async function duplicateExists(mediaType, exceptId) {
  const filter = { mediaType: new RegExp(`^${escapeRegex(mediaType)}$`, 'i') };
  if (exceptId) filter._id = { $ne: exceptId };
  return !!(await MediaRate.exists(filter));
}

// @route GET /api/media-rates  — full list, latest updated first (small table; the site form loads it all)
const getMediaRates = asyncHandler(async (req, res) => {
  const filter = {};
  const search = String(req.query.search || '').trim();
  if (search) filter.mediaType = new RegExp(escapeRegex(search), 'i');
  // Most recently added / updated first.
  const rates = await MediaRate.find(filter).sort({ updatedAt: -1, mediaType: 1 }).populate('updatedBy', 'name').lean();
  res.json(rates);
});

// @route GET /api/media-rates/:id
const getMediaRate = asyncHandler(async (req, res) => {
  const rate = await MediaRate.findById(req.params.id).lean();
  if (!rate) {
    res.status(404);
    throw new Error('Rate not found');
  }
  res.json(rate);
});

// @route POST /api/media-rates
const createMediaRate = asyncHandler(async (req, res) => {
  const { data, error } = readBody(req.body);
  if (error) {
    res.status(400);
    throw new Error(error);
  }
  if (await duplicateExists(data.mediaType)) {
    res.status(400);
    throw new Error(`A rate for "${data.mediaType}" already exists`);
  }
  const rate = await MediaRate.create({ ...data, createdBy: req.user._id, updatedBy: req.user._id });
  res.status(201).json(rate);
});

// @route PUT /api/media-rates/:id
const updateMediaRate = asyncHandler(async (req, res) => {
  const rate = await MediaRate.findById(req.params.id);
  if (!rate) {
    res.status(404);
    throw new Error('Rate not found');
  }
  const { data, error } = readBody(req.body);
  if (error) {
    res.status(400);
    throw new Error(error);
  }
  if (await duplicateExists(data.mediaType, rate._id)) {
    res.status(400);
    throw new Error(`A rate for "${data.mediaType}" already exists`);
  }
  Object.assign(rate, data, { updatedBy: req.user._id });
  await rate.save();
  res.json(rate);
});

// @route DELETE /api/media-rates/:id
const deleteMediaRate = asyncHandler(async (req, res) => {
  const rate = await MediaRate.findById(req.params.id);
  if (!rate) {
    res.status(404);
    throw new Error('Rate not found');
  }
  await rate.deleteOne();
  res.json({ message: 'Rate deleted' });
});

module.exports = { getMediaRates, getMediaRate, createMediaRate, updateMediaRate, deleteMediaRate };
