const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const {
  getSiteInfos,
  getSiteInfo,
  createSiteInfo,
  updateSiteInfo,
  deleteSiteInfo,
} = require('../controllers/siteInfoController');

const router = express.Router();

router.use(protect);

router.get('/', getSiteInfos);
router.get('/:id', getSiteInfo);
router.post('/', createSiteInfo);
router.put('/:id', updateSiteInfo);
router.delete('/:id', authorize('admin', 'tl'), deleteSiteInfo);

module.exports = router;
