const jwt = require('jsonwebtoken');
const asyncHandler = require('express-async-handler');
const User = require('../models/User');

const protect = asyncHandler(async (req, res, next) => {
  const header = req.headers.authorization;
  if (!header || !header.startsWith('Bearer ')) {
    res.status(401);
    throw new Error('Not authorized, no token');
  }
  const token = header.split(' ')[1];
  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    req.user = await User.findById(decoded.id);
    if (!req.user || !req.user.isActive) {
      res.status(401);
      throw new Error('Not authorized, user not found');
    }
    next();
  } catch (err) {
    res.status(401);
    throw new Error('Not authorized, invalid token');
  }
});

function authorize(...roles) {
  return (req, res, next) => {
    if (!req.user || !roles.includes(req.user.role)) {
      res.status(403);
      throw new Error('Forbidden: insufficient role');
    }
    next();
  };
}

module.exports = { protect, authorize };
