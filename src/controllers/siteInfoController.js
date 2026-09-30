const asyncHandler = require('express-async-handler');
const SiteInfo = require('../models/SiteInfo');

// With ?page (and optional ?limit, default 20, max 100) returns one page as
// { items, total, page, pages } — same shape as the other list endpoints. Without ?page it
// still returns the full array, which the Edit Site form's Site Information dropdown relies on.
const getSiteInfos = asyncHandler(async (req, res) => {
  const filter = {};
  if (req.query.search) {
    filter.title = new RegExp(String(req.query.search).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
  }
  const sort = { updatedAt: -1, _id: -1 };
  if (req.query.page === undefined) {
    const items = await SiteInfo.find(filter).sort(sort);
    return res.json(items);
  }

  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  const [items, total] = await Promise.all([
    SiteInfo.find(filter).sort(sort).skip((page - 1) * limit).limit(limit),
    SiteInfo.countDocuments(filter),
  ]);
  res.json({ items, total, page, pages: Math.ceil(total / limit) });
});

const getSiteInfo = asyncHandler(async (req, res) => {
  const item = await SiteInfo.findById(req.params.id);
  if (!item) {
    res.status(404);
    throw new Error('Site Quote not found');
  }
  res.json(item);
});

const createSiteInfo = asyncHandler(async (req, res) => {
  const { title, description } = req.body;
  if (!title || !String(title).trim()) {
    res.status(400);
    throw new Error('Title is required');
  }
  if (!description || !String(description).trim()) {
    res.status(400);
    throw new Error('Description is required');
  }
  const item = await SiteInfo.create({
    title: String(title).trim(),
    description: String(description).trim(),
    createdBy: req.user._id,
  });
  res.status(201).json(item);
});

const updateSiteInfo = asyncHandler(async (req, res) => {
  const item = await SiteInfo.findById(req.params.id);
  if (!item) {
    res.status(404);
    throw new Error('Site Quote not found');
  }
  if (req.body.title !== undefined) {
    if (!String(req.body.title).trim()) {
      res.status(400);
      throw new Error('Title is required');
    }
    item.title = String(req.body.title).trim();
  }
  if (req.body.description !== undefined) {
    if (!String(req.body.description).trim()) {
      res.status(400);
      throw new Error('Description is required');
    }
    item.description = String(req.body.description).trim();
  }
  await item.save();
  res.json(item);
});

const deleteSiteInfo = asyncHandler(async (req, res) => {
  const item = await SiteInfo.findById(req.params.id);
  if (!item) {
    res.status(404);
    throw new Error('Site Quote not found');
  }
  await item.deleteOne();
  res.json({ message: 'Site Quote deleted' });
});

module.exports = { getSiteInfos, getSiteInfo, createSiteInfo, updateSiteInfo, deleteSiteInfo };
