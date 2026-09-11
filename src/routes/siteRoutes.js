const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const {
  getSites,
  getAvailableSites,
  getSite,
  createSite,
  updateSite,
  deleteSite,
  changeStatus,
  bulkImport,
} = require('../controllers/siteController');

const router = express.Router();

router.use(protect);

router.get('/available', getAvailableSites);
router.get('/', getSites);
router.get('/:id', getSite);
router.post('/', authorize('admin', 'tl'), createSite);
router.post('/bulk-import', authorize('admin', 'tl'), bulkImport);
router.put('/:id', authorize('admin', 'tl'), updateSite);
router.patch('/:id/status', authorize('admin', 'tl'), changeStatus);
router.delete('/:id', authorize('admin', 'tl'), deleteSite);

module.exports = router;
