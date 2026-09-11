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

const createClient = asyncHandler(async (req, res) => {
  const client = await Client.create({ ...req.body, createdBy: req.user._id });
  res.status(201).json(client);
});

const updateClient = asyncHandler(async (req, res) => {
  const client = await Client.findById(req.params.id);
  if (!client) {
    res.status(404);
    throw new Error('Client not found');
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
