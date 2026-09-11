const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const {
  getClients,
  getClient,
  createClient,
  updateClient,
  deleteClient,
} = require('../controllers/clientController');

const router = express.Router();

router.use(protect);

router.get('/', getClients);
router.get('/:id', getClient);
router.post('/', createClient);
router.put('/:id', updateClient);
router.delete('/:id', authorize('admin', 'tl'), deleteClient);

module.exports = router;
