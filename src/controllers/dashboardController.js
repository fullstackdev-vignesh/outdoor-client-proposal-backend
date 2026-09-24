const asyncHandler = require('express-async-handler');
const Site = require('../models/Site');
const Client = require('../models/Client');
const Booking = require('../models/Booking');
const Proposal = require('../models/Proposal');
const User = require('../models/User');
const InventoryHistory = require('../models/InventoryHistory');

async function getRecentBookings(limit = 5) {
  const allBookings = [];
  const existingBookingIds = new Set();

  // 1. Fetch site bookings from Site collection
  const sitesWithBookings = await Site.find({
    'bookings.0': { $exists: true },
  })
    .select('mediaId mediaCode mediaType city state bookings')
    .lean();

  for (const site of sitesWithBookings) {
    if (Array.isArray(site.bookings)) {
      for (const b of site.bookings) {
        const key = b.bookingId || site._id.toString();
        if (!existingBookingIds.has(key)) {
          allBookings.push({
            id: key,
            bookingId: b.bookingId || 'BOOKING',
            mediaCode: site.mediaCode || site.mediaId,
            mediaId: site.mediaId,
            mediaType: site.mediaType || '',
            clientName: b.customerName || (typeof b.client === 'object' ? b.client?.name : '') || '',
            startDate: b.startDate,
            endDate: b.endDate,
            status: b.status || 'booked',
            amount: b.amount || 0,
            createdAt: b.createdAt || b.updatedAt || new Date(0),
          });
          existingBookingIds.add(key);
        }
      }
    }
  }

  // 2. Fetch from InventoryHistory for booked/cancelled entries
  const historyBookings = await InventoryHistory.find({
    status: { $in: ['booked', 'cancelled'] },
  })
    .sort({ changedAt: -1 })
    .limit(10)
    .lean();

  for (const h of historyBookings) {
    const key = h.bookingId || h._id.toString();
    if (!existingBookingIds.has(key)) {
      allBookings.push({
        id: key,
        bookingId: h.bookingId || 'BOOKING',
        mediaCode: h.mediaId,
        mediaId: h.mediaId,
        mediaType: h.mediaType || '',
        clientName: h.bookingSnapshot?.customerName || '',
        startDate: h.bookingSnapshot?.startDate,
        endDate: h.bookingSnapshot?.endDate,
        status: h.status,
        amount: h.bookingSnapshot?.amount || 0,
        createdAt: h.changedAt || new Date(0),
      });
      existingBookingIds.add(key);
    }
  }

  // 3. Fetch from Booking collection if legacy bookings exist
  const legacyBookings = await Booking.find()
    .populate('client', 'name')
    .sort({ createdAt: -1 })
    .limit(limit)
    .lean();

  for (const lb of legacyBookings) {
    const key = lb.bookingId || lb._id.toString();
    if (!existingBookingIds.has(key)) {
      allBookings.push({
        id: key,
        bookingId: lb.bookingId || 'BOOKING',
        mediaCode: 'BOOKING',
        mediaId: 'BOOKING',
        mediaType: '',
        clientName: typeof lb.client === 'object' ? lb.client?.name : '',
        startDate: lb.startDate,
        endDate: lb.endDate,
        status: lb.status || 'active',
        amount: lb.totalAmount || lb.amount || 0,
        createdAt: lb.createdAt || new Date(0),
      });
      existingBookingIds.add(key);
    }
  }

  // Sort by createdAt descending (newest / most recent first)
  allBookings.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

  return allBookings.slice(0, limit);
}

const getDashboardStats = asyncHandler(async (req, res) => {
  const [
    totalUsers,
    totalTLs,
    totalBDs,
    totalSites,
    activeSites,
    inactiveSites,
    availableMedia,
    bookedMedia,
    blockedMedia,
    totalClients,
    totalProposals,
    recentSites,
    recentClients,
    recentBookings,
    recentProposals,
    recentUsers,
  ] = await Promise.all([
    User.countDocuments({ role: 'user' }),
    User.countDocuments({ role: 'tl' }),
    User.countDocuments({ role: 'bd' }),
    Site.countDocuments(),
    Site.countDocuments({ isActive: true }),
    Site.countDocuments({ isActive: false }),
    Site.countDocuments({ mediaStatus: 'available' }),
    Site.countDocuments({ mediaStatus: 'booked' }),
    Site.countDocuments({ mediaStatus: 'blocked' }),
    Client.countDocuments(),
    Proposal.countDocuments(),
    Site.find().sort({ createdAt: -1 }).limit(5),
    Client.find().sort({ createdAt: -1 }).limit(5),
    getRecentBookings(5),
    Proposal.find().populate('client', 'name').sort({ createdAt: -1 }).limit(5),
    User.find().sort({ createdAt: -1 }).limit(5),
  ]);

  const siteBookingsCountAgg = await Site.aggregate([
    { $unwind: '$bookings' },
    { $count: 'count' },
  ]);
  const totalSiteBookingsCount = siteBookingsCountAgg[0]?.count || 0;
  const legacyBookingsCount = await Booking.countDocuments();
  const totalBookings = totalSiteBookingsCount + legacyBookingsCount;

  res.json({
    cards: {
      totalUsers,
      totalTLs,
      totalBDs,
      totalSites,
      activeSites,
      inactiveSites,
      availableMedia,
      bookedMedia,
      blockedMedia,
      totalClients,
      totalProposals,
      totalBookings,
    },
    mediaStatusSummary: { available: availableMedia, booked: bookedMedia, blocked: blockedMedia },
    recent: { sites: recentSites, clients: recentClients, bookings: recentBookings, proposals: recentProposals, users: recentUsers },
  });
});

const getReports = asyncHandler(async (req, res) => {
  const siteBookingsCountAgg = await Site.aggregate([
    { $unwind: '$bookings' },
    { $count: 'count' },
  ]);
  const totalSiteBookings = siteBookingsCountAgg[0]?.count || 0;
  const legacyBookingsTotal = await Booking.countDocuments();
  const totalBookingsAll = totalSiteBookings + legacyBookingsTotal;

  const [siteReport, bookingReport, clientReport, proposalReport] = await Promise.all([
    (async () => ({
      total: await Site.countDocuments(),
      active: await Site.countDocuments({ isActive: true }),
      inactive: await Site.countDocuments({ isActive: false }),
      available: await Site.countDocuments({ mediaStatus: 'available' }),
      booked: await Site.countDocuments({ mediaStatus: 'booked' }),
      blocked: await Site.countDocuments({ mediaStatus: 'blocked' }),
    }))(),
    (async () => ({
      total: totalBookingsAll,
      active: await Site.countDocuments({ mediaStatus: 'booked' }),
      completed: 0,
      cancelled: (await InventoryHistory.countDocuments({ status: 'cancelled' })),
    }))(),
    (async () => ({
      total: await Client.countDocuments(),
      withBookings: (await Site.distinct('bookings.client')).length,
      withProposals: (await Proposal.distinct('client')).length,
    }))(),
    (async () => ({
      total: await Proposal.countDocuments(),
      draft: await Proposal.countDocuments({ status: 'draft' }),
      generated: await Proposal.countDocuments({ status: 'generated' }),
      completed: await Proposal.countDocuments({ status: 'completed' }),
    }))(),
  ]);

  res.json({ siteReport, bookingReport, clientReport, proposalReport });
});

module.exports = { getDashboardStats, getReports };
