const express = require('express');
const { protect } = require('../middleware/auth');
const {
  getProposals,
  getProposal,
  createProposal,
  updateProposal,
  deleteProposal,
  duplicateProposal,
  generatePpt,
  generateExcel,
} = require('../controllers/proposalController');

const router = express.Router();

router.use(protect);

router.get('/', getProposals);
router.get('/:id', getProposal);
router.post('/', createProposal);
router.put('/:id', updateProposal);
router.delete('/:id', deleteProposal);
router.post('/:id/duplicate', duplicateProposal);
router.post('/:id/generate-ppt', generatePpt);
router.post('/:id/generate-excel', generateExcel);

module.exports = router;
