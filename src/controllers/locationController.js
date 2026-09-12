const asyncHandler = require('express-async-handler');
const Location = require('../models/Location');

const IGNORED_KEYS = new Set(['_id', '__v']);

const getStates = asyncHandler(async (req, res) => {
  const docs = await Location.find({}).lean();
  const states = new Set();
  for (const doc of docs) {
    for (const key of Object.keys(doc)) {
      if (!IGNORED_KEYS.has(key) && Array.isArray(doc[key])) states.add(key);
    }
  }
  res.json([...states].sort());
});

const getCities = asyncHandler(async (req, res) => {
  const { state } = req.params;
  const docs = await Location.find({ [state]: { $exists: true } }).lean();
  const cities = new Set();
  for (const doc of docs) {
    if (Array.isArray(doc[state])) doc[state].forEach((c) => cities.add(c));
  }
  res.json([...cities].sort());
});

module.exports = { getStates, getCities };
