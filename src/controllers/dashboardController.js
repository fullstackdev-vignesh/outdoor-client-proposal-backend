const asyncHandler = require('express-async-handler');
const Site = require('../models/Site');
const Client = require('../models/Client');
const Booking = require('../models/Booking');
const Proposal = require('../models/Proposal');
const User = require('../models/User');

const getDashboardStats = asyncHandler(async (req, res) => {
  const [
    totalUsers,
    totalTLs,
    totalSites,
    activeSites,
    inactiveSites,
    availableMedia,
    bookedMedia,
    blockedMedia,
    totalClients,
    totalProposals,
    totalBookings,
    recentSites,
    recentClients,
    recentBookings,
    recentProposals,
    recentUsers,
  ] = await Promise.all([
    User.countDocuments({ role: 'user' }),
    User.countDocuments({ role: 'tl' }),
    Site.countDocuments(),
    Site.countDocuments({ isActive: true }),
    Site.countDocuments({ isActive: false }),
    Site.countDocuments({ mediaStatus: 'available' }),
    Site.countDocuments({ mediaStatus: 'booked' }),
    Site.countDocuments({ mediaStatus: 'blocked' }),
    Client.countDocuments(),
    Proposal.countDocuments(),
    Booking.countDocuments(),
    Site.find().sort({ createdAt: -1 }).limit(5),
    Client.find().sort({ createdAt: -1 }).limit(5),
    Booking.find().populate('client', 'name').sort({ createdAt: -1 }).limit(5),
    Proposal.find().populate('client', 'name').sort({ createdAt: -1 }).limit(5),
    User.find().sort({ createdAt: -1 }).limit(5),
  ]);

  res.json({
    cards: {
      totalUsers,
      totalTLs,
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
      total: await Booking.countDocuments(),
      active: await Booking.countDocuments({ status: 'active' }),
      completed: await Booking.countDocuments({ status: 'completed' }),
      cancelled: await Booking.countDocuments({ status: 'cancelled' }),
    }))(),
    (async () => ({
      total: await Client.countDocuments(),
      withBookings: (await Booking.distinct('client')).length,
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
