const asyncHandler = require('express-async-handler');
const Client = require('../models/Client');
const Site = require('../models/Site');
const Booking = require('../models/Booking');
const Proposal = require('../models/Proposal');

const getClients = asyncHandler(async (req, res) => {
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  const filter = {};
  if (req.query.search) {
    filter.$or = [
      { name: new RegExp(req.query.search, 'i') },
      { phone: new RegExp(req.query.search, 'i') },
      { email: new RegExp(req.query.search, 'i') },
    ];
  }
  if (req.query.customerType) filter.customerType = req.query.customerType;
  const [items, total] = await Promise.all([
    Client.find(filter).sort({ createdAt: -1 }).skip((page - 1) * limit).limit(limit),
    Client.countDocuments(filter),
  ]);
  res.json({ items, total, page, pages: Math.ceil(total / limit) });
});

const getClient = asyncHandler(async (req, res) => {
  const client = await Client.findById(req.params.id);
  if (!client) {
    res.status(404);
    throw new Error('Client not found');
  }
  const [bookings, proposals, siteCount] = await Promise.all([
    Booking.find({ client: client._id }).sort({ createdAt: -1 }),
    Proposal.find({ client: client._id }).sort({ createdAt: -1 }),
    Site.countDocuments({ 'bookingInfo.client': client._id }),
  ]);
  res.json({ client, bookings, proposals, siteCount });
});

function validateClientPayload(body, { requireGeo = false } = {}) {
  const errors = [];
  const isBlank = (v) => v === '' || v === undefined || v === null;
  const num = (v) => (isBlank(v) ? undefined : Number(v));

  if (!body.name || !String(body.name).trim()) {
    errors.push(body.customerType === 'agency' ? 'Agency name is required' : 'Client name is required');
  }
  if (body.customerType && !['client', 'agency'].includes(body.customerType)) {
    errors.push('Invalid customer type');
  }
  if (requireGeo && isBlank(body.latitude)) errors.push('Latitude is required');
  if (requireGeo && isBlank(body.longitude)) errors.push('Longitude is required');
  if (!isBlank(body.latitude) && (isNaN(num(body.latitude)) || num(body.latitude) < -90 || num(body.latitude) > 90)) {
    errors.push('Latitude must be between -90 and 90');
  }
  if (!isBlank(body.longitude) && (isNaN(num(body.longitude)) || num(body.longitude) < -180 || num(body.longitude) > 180)) {
    errors.push('Longitude must be between -180 and 180');
  }
  if (!isBlank(body.agencyComm) && (isNaN(num(body.agencyComm)) || num(body.agencyComm) < 0)) {
    errors.push('Agency Comm must be a number greater than or equal to 0');
  }
  if (!isBlank(body.gst) && (isNaN(num(body.gst)) || num(body.gst) < 0 || num(body.gst) > 100)) {
    errors.push('GST must be a percentage between 0 and 100');
  }
  if (!isBlank(body.vendorCost) && (isNaN(num(body.vendorCost)) || num(body.vendorCost) < 0)) {
    errors.push('Vendor Cost must be a number greater than or equal to 0');
  }
  return errors;
}

const createClient = asyncHandler(async (req, res) => {
  const errors = validateClientPayload(req.body, { requireGeo: true });
  if (errors.length) {
    res.status(400);
    throw new Error(errors.join('; '));
  }
  const client = await Client.create({ ...req.body, createdBy: req.user._id });
  res.status(201).json(client);
});

const updateClient = asyncHandler(async (req, res) => {
  const client = await Client.findById(req.params.id);
  if (!client) {
    res.status(404);
    throw new Error('Client not found');
  }
  const errors = validateClientPayload({ ...client.toObject(), ...req.body });
  if (errors.length) {
    res.status(400);
    throw new Error(errors.join('; '));
  }
  Object.assign(client, req.body);
  await client.save();
  res.json(client);
});

const deleteClient = asyncHandler(async (req, res) => {
  const client = await Client.findById(req.params.id);
  if (!client) {
    res.status(404);
    throw new Error('Client not found');
  }
  await client.deleteOne();
  res.json({ message: 'Client deleted' });
});

module.exports = { getClients, getClient, createClient, updateClient, deleteClient };
