const express = require('express');
const { protect } = require('../middleware/auth');
const { getStates, getCities } = require('../controllers/locationController');

const router = express.Router();

router.use(protect);

router.get('/states', getStates);
router.get('/states/:state/cities', getCities);

module.exports = router;
