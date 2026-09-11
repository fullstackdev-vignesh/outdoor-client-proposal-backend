const express = require('express');
const { protect } = require('../middleware/auth');
const { getDashboardStats, getReports } = require('../controllers/dashboardController');

const router = express.Router();

router.use(protect);

router.get('/stats', getDashboardStats);
router.get('/reports', getReports);

module.exports = router;
