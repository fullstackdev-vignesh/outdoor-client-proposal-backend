const mongoose = require('mongoose');

const locationSchema = new mongoose.Schema({}, { collection: 'location', strict: false });

module.exports = mongoose.model('Location', locationSchema);
