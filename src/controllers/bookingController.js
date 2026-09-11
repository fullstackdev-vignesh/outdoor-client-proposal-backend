const asyncHandler = require('express-async-handler');
const Booking = require('../models/Booking');
const Site = require('../models/Site');

const genBookingId = () => `BK-${Date.now().toString(36).toUpperCase()}`;

const getBookings = asyncHandler(async (req, res) => {
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  if (req.query.bookingId) filter.bookingId = new RegExp(req.query.bookingId, 'i');
  if (req.query.startDate || req.query.endDate) {
    filter.startDate = {};
    if (req.query.startDate) filter.startDate.$gte = new Date(req.query.startDate);
    if (req.query.endDate) filter.startDate.$lte = new Date(req.query.endDate);
  }

  const [items, total] = await Promise.all([
    Booking.find(filter)
      .populate('client', 'name')
      .populate('sites', 'mediaName mediaId city state')
      .sort({ createdAt: -1 })
      .skip((page - 1) * limit)
      .limit(limit),
    Booking.countDocuments(filter),
  ]);
  res.json({ items, total, page, pages: Math.ceil(total / limit) });
});

const getBooking = asyncHandler(async (req, res) => {
  const booking = await Booking.findById(req.params.id)
    .populate('client')
    .populate('sites')
    .populate('createdBy', 'name');
  if (!booking) {
    res.status(404);
    throw new Error('Booking not found');
  }
  res.json(booking);
});

const createBooking = asyncHandler(async (req, res) => {
  const { client, sites, startDate, endDate, amount, gstAmount, totalAmount, notes } = req.body;
  if (!Array.isArray(sites) || sites.length === 0) {
    res.status(400);
    throw new Error('At least one media/site must be selected');
  }

  const siteDocs = await Site.find({ _id: { $in: sites } });
  const unavailable = siteDocs.filter((s) => s.mediaStatus !== 'available');
  if (unavailable.length > 0) {
    res.status(400);
    throw new Error(
      `The following media are not available: ${unavailable.map((s) => s.mediaName).join(', ')}`
    );
  }

  const booking = await Booking.create({
    bookingId: genBookingId(),
    client,
    sites,
    startDate,
    endDate,
    amount,
    gstAmount,
    totalAmount,
    notes,
    status: 'active',
    createdBy: req.user._id,
  });

  await Site.updateMany(
    { _id: { $in: sites } },
    {
      $set: {
        mediaStatus: 'booked',
        bookingInfo: {
          client,
          booking: booking._id,
          bookingRef: booking.bookingId,
          startDate,
          endDate,
          amount: totalAmount,
          bookedBy: req.user._id,
        },
      },
    }
  );

  res.status(201).json(booking);
});

const cancelBooking = asyncHandler(async (req, res) => {
  const booking = await Booking.findById(req.params.id);
  if (!booking) {
    res.status(404);
    throw new Error('Booking not found');
  }
  if (booking.status === 'cancelled') {
    res.status(400);
    throw new Error('Booking already cancelled');
  }

  booking.status = 'cancelled';
  await booking.save();

  await Site.updateMany(
    { _id: { $in: booking.sites }, 'bookingInfo.booking': booking._id },
    { $set: { mediaStatus: 'available' }, $unset: { bookingInfo: '' } }
  );

  res.json(booking);
});

const updateBooking = asyncHandler(async (req, res) => {
  const booking = await Booking.findById(req.params.id);
  if (!booking) {
    res.status(404);
    throw new Error('Booking not found');
  }
  const { startDate, endDate, amount, gstAmount, totalAmount, notes } = req.body;
  Object.assign(booking, { startDate, endDate, amount, gstAmount, totalAmount, notes });
  await booking.save();
  res.json(booking);
});

module.exports = { getBookings, getBooking, createBooking, cancelBooking, updateBooking };
