const asyncHandler = require('express-async-handler');
const { saveTemplateFile } = require('../utils/templateStorage');

function makeTemplateController(Model, options = {}) {
  const fileHandler = options.fileHandler || saveTemplateFile;
  const getAll = asyncHandler(async (req, res) => {
    const filter = {};
    if (req.query.status) filter.status = req.query.status;
    if (req.query.search) filter.name = new RegExp(req.query.search, 'i');
    const items = await Model.find(filter).sort({ updatedAt: -1, _id: -1 });
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
      try {
        payload.fileUrl = await fileHandler(req.file, payload);
      } catch (err) {
        res.status(400);
        throw err;
      }
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
      try {
        req.body.fileUrl = await fileHandler(req.file, { ...item.toObject(), ...req.body });
      } catch (err) {
        res.status(400);
        throw err;
      }
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

module.exports = makeTemplateFactory = makeTemplateController;
module.exports = makeTemplateController;
