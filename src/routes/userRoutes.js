const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const { getUsers, getUser, createUser, updateUser, deleteUser } = require('../controllers/userController');

const router = express.Router();

router.use(protect, authorize('admin', 'tl'));

router.get('/', getUsers);
router.get('/:id', getUser);
router.post('/', createUser);
router.put('/:id', updateUser);
router.delete('/:id', deleteUser);

module.exports = router;
