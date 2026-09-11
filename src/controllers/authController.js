const asyncHandler = require('express-async-handler');
const User = require('../models/User');
const generateToken = require('../utils/generateToken');

const login = asyncHandler(async (req, res) => {
  const { email, password, role } = req.body;
  const user = await User.findOne({ email }).select('+password');
  if (!user || !(await user.comparePassword(password))) {
    res.status(401);
    throw new Error('Invalid email or password');
  }
  if (role && user.role !== role) {
    res.status(401);
    throw new Error('Invalid role for this account');
  }
  if (!user.isActive) {
    res.status(403);
    throw new Error('Account is inactive');
  }
  res.json({
    token: generateToken(user),
    user: {
      id: user._id,
      name: user.name,
      email: user.email,
      role: user.role,
    },
  });
});

const me = asyncHandler(async (req, res) => {
  res.json({ user: req.user });
});

module.exports = { login, me };
