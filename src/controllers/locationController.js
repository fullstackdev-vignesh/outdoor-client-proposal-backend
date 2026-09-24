const asyncHandler = require('express-async-handler');
const Site = require('../models/Site');

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

module.exports = { getStates, getCities, STATIC_STATES };
