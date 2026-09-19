const asyncHandler = require('express-async-handler');
const Proposal = require('../models/Proposal');
const Site = require('../models/Site');
const Client = require('../models/Client');
const { generateProposalPpt, generateProposalExcel } = require('../utils/proposalFileGenerator');

const genProposalId = () => `PR-${Date.now().toString(36).toUpperCase()}`;

const getProposals = asyncHandler(async (req, res) => {
  const page = Math.max(1, Number(req.query.page) || 1);
  const limit = Math.min(100, Number(req.query.limit) || 20);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  if (req.query.proposalId) filter.proposalId = new RegExp(req.query.proposalId, 'i');

  const [items, total] = await Promise.all([
    Proposal.find(filter)
      .populate('client', 'name')
      .populate('pptTemplate', 'name')
      .populate('excelTemplate', 'name')
      .sort({ createdAt: -1 })
      .skip((page - 1) * limit)
      .limit(limit),
    Proposal.countDocuments(filter),
  ]);
  res.json({ items, total, page, pages: Math.ceil(total / limit) });
});

const getProposal = asyncHandler(async (req, res) => {
  const proposal = await Proposal.findById(req.params.id)
    .populate('client')
    .populate('sites')
    .populate('pptTemplate')
    .populate('excelTemplate');
  if (!proposal) {
    res.status(404);
    throw new Error('Proposal not found');
  }
  res.json(proposal);
});

const createProposal = asyncHandler(async (req, res) => {
  const { client, sites, pptTemplate, excelTemplate, variant, totalAmount, gstAmount, monthlyAmount } = req.body;
  if (!client) {
    res.status(400);
    throw new Error('A client must be selected');
  }
  const clientDoc = await Client.findById(client);
  if (!clientDoc) {
    res.status(400);
    throw new Error('Selected client could not be found');
  }
  if (!Array.isArray(sites) || sites.length === 0) {
    res.status(400);
    throw new Error('At least one media/site must be selected');
  }

  const siteDocs = await Site.find({ _id: { $in: sites } });
  if (siteDocs.length !== sites.length) {
    res.status(400);
    throw new Error('One or more selected sites could not be found');
  }
  const unavailable = siteDocs.filter((s) => s.mediaStatus !== 'available');
  if (unavailable.length > 0) {
    res.status(400);
    throw new Error(
      `The following media cannot be added to a proposal: ${unavailable.map((s) => s.mediaName).join(', ')}`
    );
  }

  const proposal = await Proposal.create({
    proposalId: genProposalId(),
    client,
    sites,
    pptTemplate,
    excelTemplate,
    variant,
    totalAmount,
    gstAmount,
    monthlyAmount,
    status: 'draft',
    createdBy: req.user._id,
  });

  res.status(201).json(proposal);
});

const updateProposal = asyncHandler(async (req, res) => {
  const proposal = await Proposal.findById(req.params.id);
  if (!proposal) {
    res.status(404);
    throw new Error('Proposal not found');
  }
  Object.assign(proposal, req.body);
  await proposal.save();
  res.json(proposal);
});

const deleteProposal = asyncHandler(async (req, res) => {
  const proposal = await Proposal.findById(req.params.id);
  if (!proposal) {
    res.status(404);
    throw new Error('Proposal not found');
  }
  await proposal.deleteOne();
  res.json({ message: 'Proposal deleted' });
});

const duplicateProposal = asyncHandler(async (req, res) => {
  const source = await Proposal.findById(req.params.id).lean();
  if (!source) {
    res.status(404);
    throw new Error('Proposal not found');
  }
  delete source._id;
  const copy = await Proposal.create({
    ...source,
    proposalId: genProposalId(),
    status: 'draft',
    generatedPptUrl: undefined,
    generatedExcelUrl: undefined,
    createdBy: req.user._id,
  });
  res.status(201).json(copy);
});

const generatePpt = asyncHandler(async (req, res) => {
  const proposal = await Proposal.findById(req.params.id)
    .populate('client')
    .populate({ path: 'sites', populate: { path: 'siteInfoId' } })
    .populate('pptTemplate');
  if (!proposal) {
    res.status(404);
    throw new Error('Proposal not found');
  }
  try {
    proposal.generatedPptUrl = await generateProposalPpt(proposal);
    proposal.status = proposal.generatedExcelUrl ? 'completed' : 'generated';
    await proposal.save();
    res.json(proposal);
  } catch (err) {
    console.error('generatePpt error:', err);
    res.status(500);
    throw new Error(`Failed to generate PPT: ${err.message}`);
  }
});

const generateExcel = asyncHandler(async (req, res) => {
  const proposal = await Proposal.findById(req.params.id)
    .populate('client')
    .populate('sites')
    .populate('excelTemplate');
  if (!proposal) {
    res.status(404);
    throw new Error('Proposal not found');
  }
  try {
    proposal.generatedExcelUrl = await generateProposalExcel(proposal);
    proposal.status = proposal.generatedPptUrl ? 'completed' : 'generated';
    await proposal.save();
    res.json(proposal);
  } catch (err) {
    console.error('generateExcel error:', err);
    res.status(500);
    throw new Error(`Failed to generate Excel: ${err.message}`);
  }
});

module.exports = {
  getProposals,
  getProposal,
  createProposal,
  updateProposal,
  deleteProposal,
  duplicateProposal,
  generatePpt,
  generateExcel,
};
