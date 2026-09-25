const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const upload = require('../middleware/upload');
const {
  getClients,
  getClient,
  createClient,
  updateClient,
  deleteClient,
} = require('../controllers/clientController');

const router = express.Router();

router.use(protect);

const uploadFields = upload.fields([{ name: 'clientLocationPinImage', maxCount: 1 }]);

router.get('/', getClients);
router.get('/:id', getClient);
router.post('/', uploadFields, createClient);
router.put('/:id', uploadFields, updateClient);
router.delete('/:id', authorize('admin', 'tl','user'), deleteClient);

module.exports = router;
