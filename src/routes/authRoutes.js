const express = require('express');
const { login, register, me, forgotPin, forgotPinVerify, resetPin } = require('../controllers/authController');
const { protect } = require('../middleware/auth');

const router = express.Router();

router.post('/login', login);
router.post('/register', register);
router.post('/forgot-pin', forgotPin);
router.post('/forgot-pin-verify', forgotPinVerify);
router.post('/reset-pin', resetPin);
router.get('/me', protect, me);

module.exports = router;
