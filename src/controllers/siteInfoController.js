const asyncHandler = require('express-async-handler');
const SiteInfo = require('../models/SiteInfo');

const getSiteInfos = asyncHandler(async (req, res) => {
  const filter = {};
  if (req.query.search) {
    filter.title = new RegExp(req.query.search, 'i');
  }
  const items = await SiteInfo.find(filter).sort({ updatedAt: -1, _id: -1 });
  res.json(items);
});

const getSiteInfo = asyncHandler(async (req, res) => {
  const item = await SiteInfo.findById(req.params.id);
  if (!item) {
    res.status(404);
    throw new Error('Site information not found');
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
    throw new Error('Site information not found');
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
    throw new Error('Site information not found');
  }
  await item.deleteOne();
  res.json({ message: 'Site information deleted' });
});

module.exports = { getSiteInfos, getSiteInfo, createSiteInfo, updateSiteInfo, deleteSiteInfo };
