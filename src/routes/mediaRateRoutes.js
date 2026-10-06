const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const {
  getMediaRates,
  getMediaRate,
  createMediaRate,
  updateMediaRate,
  deleteMediaRate,
} = require('../controllers/mediaRateController');

const router = express.Router();

router.use(protect);

// Everyone signed in reads rates (the site form uses them); only admins change them.
router.get('/', getMediaRates);
router.get('/:id', getMediaRate);
router.post('/', authorize('admin'), createMediaRate);
router.put('/:id', authorize('admin'), updateMediaRate);
router.delete('/:id', authorize('admin'), deleteMediaRate);

module.exports = router;
