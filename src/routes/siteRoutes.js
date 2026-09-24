const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const upload = require('../middleware/upload');
const {
  getSites,
  getAvailableSites,
  getSite,
  createSite,
  updateSite,
  deleteSite,
  changeStatus,
  cancelBooking,
  bulkChangeStatus,
  bulkImport,
  uploadImage,
  exportSites,
  getSiteHistory,
  getSummary,
  getSiteOwners,
  getSiteTimeline,
  getTimeline,
  getTimelineSummary,
  exportTimeline,
  getStates,
  getCities,
} = require('../controllers/siteController');

const router = express.Router();

router.use(protect);

router.get('/states', getStates);
router.get('/states/:state/cities', getCities);
router.get('/cities', getCities);
router.get('/available', getAvailableSites);
router.get('/summary', getSummary);
router.get('/owners', getSiteOwners);
router.get('/export', exportSites);
router.get('/timeline', getTimeline);
router.get('/timeline/summary', getTimelineSummary);
router.get('/timeline/export', exportTimeline);
router.get('/', getSites);
router.get('/:id/history', getSiteHistory);
router.get('/:id/timeline', getSiteTimeline);
router.get('/:id', getSite);
const uploadFields = upload.fields([
  { name: 'mediaImage', maxCount: 1 },
  { name: 'image', maxCount: 1 },
  { name: 'file', maxCount: 1 },
]);

router.post('/', authorize('admin', 'tl'), uploadFields, createSite);
router.post('/upload-image', authorize('admin', 'tl'), uploadFields, uploadImage);
router.post('/bulk-import', authorize('admin', 'tl'), bulkImport);
router.patch('/bulk-status', authorize('admin', 'tl'), bulkChangeStatus);
router.put('/:id', authorize('admin', 'tl'), uploadFields, updateSite);
router.patch('/:id/status', authorize('admin', 'tl'), changeStatus);
router.patch('/:id/bookings/:bookingId/cancel', authorize('admin', 'tl'), cancelBooking);
// Deleting a site is admin-only — TL, BD and User roles cannot delete.
router.delete('/:id', authorize('admin'), deleteSite);

module.exports = router;
