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
  bulkChangeStatus,
  bulkImport,
  uploadImage,
  exportSites,
  getSiteHistory,
  getSummary,
} = require('../controllers/siteController');

const router = express.Router();

router.use(protect);

router.get('/available', getAvailableSites);
router.get('/summary', getSummary);
router.get('/export', exportSites);
router.get('/', getSites);
router.get('/:id/history', getSiteHistory);
router.get('/:id', getSite);
router.post('/', authorize('admin', 'tl'), createSite);
router.post('/upload-image', authorize('admin', 'tl'), upload.single('image'), uploadImage);
router.post('/bulk-import', authorize('admin', 'tl'), bulkImport);
router.patch('/bulk-status', authorize('admin', 'tl'), bulkChangeStatus);
router.put('/:id', authorize('admin', 'tl'), updateSite);
router.patch('/:id/status', authorize('admin', 'tl'), changeStatus);
router.delete('/:id', authorize('admin', 'tl'), deleteSite);

module.exports = router;
