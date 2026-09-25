const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const {
  getBookings,
  getBooking,
  createBooking,
  cancelBooking,
  updateBooking,
} = require('../controllers/bookingController');

const router = express.Router();

router.use(protect);

router.get('/', getBookings);
router.get('/:id', getBooking);
router.post('/', authorize('admin', 'tl', 'user', 'bd'), createBooking);
router.put('/:id', authorize('admin', 'tl', 'user', 'bd'), updateBooking);
router.patch('/:id/cancel', authorize('admin', 'tl', 'user', 'bd'), cancelBooking);

module.exports = router;
