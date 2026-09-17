const asyncHandler = require('express-async-handler');
const { saveTemplateFile } = require('../utils/templateStorage');

function makeTemplateController(Model) {
  const getAll = asyncHandler(async (req, res) => {
    const filter = {};
    if (req.query.status) filter.status = req.query.status;
    if (req.query.search) filter.name = new RegExp(req.query.search, 'i');
    const items = await Model.find(filter).sort({ createdAt: -1 });
    res.json(items);
  });

  const getOne = asyncHandler(async (req, res) => {
    const item = await Model.findById(req.params.id);
    if (!item) {
      res.status(404);
      throw new Error('Template not found');
    }
    res.json(item);
  });

  const create = asyncHandler(async (req, res) => {
    const payload = { ...req.body, createdBy: req.user._id };
    if (req.file) {
      payload.fileUrl = await saveTemplateFile(req.file);
    }
    const item = await Model.create(payload);
    res.status(201).json(item);
  });

  const update = asyncHandler(async (req, res) => {
    const item = await Model.findById(req.params.id);
    if (!item) {
      res.status(404);
      throw new Error('Template not found');
    }
    if (req.file) {
      req.body.fileUrl = await saveTemplateFile(req.file);
    }
    Object.assign(item, req.body);
    await item.save();
    res.json(item);
  });

  const remove = asyncHandler(async (req, res) => {
    const item = await Model.findById(req.params.id);
    if (!item) {
      res.status(404);
      throw new Error('Template not found');
    }
    await item.deleteOne();
    res.json({ message: 'Template deleted' });
  });

  const setStatus = asyncHandler(async (req, res) => {
    const item = await Model.findById(req.params.id);
    if (!item) {
      res.status(404);
      throw new Error('Template not found');
    }
    item.status = req.body.status;
    await item.save();
    res.json(item);
  });

  return { getAll, getOne, create, update, remove, setStatus };
}

module.exports = makeTemplateController;
