const asyncHandler = require('express-async-handler');
const User = require('../models/User');
const generateToken = require('../utils/generateToken');

function mapUserTypeToRole(userType) {
  const typeNum = Number(userType);
  if (typeNum === 4) return 'bd';
  if (typeNum === 3) return 'admin';
  if (typeNum === 2) return 'tl';
  return 'user';
}

function mapRoleToUserType(role) {
  if (role === 'bd') return 4;
  if (role === 'admin') return 3;
  if (role === 'tl') return 2;
  return 1;
}

const register = asyncHandler(async (req, res) => {
  const userName = req.body.userName || req.body.name;
  let userEmail = req.body.userEmail || req.body.email;
  const userPhone = req.body.userPhone || req.body.phone;
  const password = req.body.password || req.body.registerPassword;
  const confirmPassword = req.body.confirmPassword;
  const userType = req.body.userType !== undefined ? req.body.userType : 1;

  if (!userName || !userPhone || !password) {
    res.status(400);
    throw new Error('userName, userPhone, and password are required');
  }

  if (confirmPassword !== undefined && password !== confirmPassword) {
    res.status(400);
    throw new Error('Password and confirmPassword do not match');
  }

  if (password.length < 6) {
    res.status(400);
    throw new Error('Password must be at least 6 characters');
  }

  if (userEmail && typeof userEmail === 'string') {
    userEmail = userEmail.trim().toLowerCase();
    if (userEmail === '') userEmail = undefined;
  } else {
    userEmail = undefined;
  }

  const trimmedPhone = userPhone.trim();

  const existingPhoneUser = await User.findOne({ phone: trimmedPhone });
  if (existingPhoneUser) {
    res.status(400);
    throw new Error('Phone number is already registered');
  }

  if (userEmail) {
    const existingEmailUser = await User.findOne({ email: userEmail });
    if (existingEmailUser) {
      res.status(400);
      throw new Error('Email is already registered');
    }
  }

  const role = mapUserTypeToRole(userType);

  const user = await User.create({
    name: userName.trim(),
    email: userEmail,
    phone: trimmedPhone,
    password,
    userType: Number(userType) || 1,
    role,
    isPhoneVerified: true,
    isActive: true,
  });

  const token = generateToken(user);

  res.status(201).json({
    status: true,
    message: 'User registered successfully',
    token,
    user: {
      id: user._id,
      userName: user.name,
      userEmail: user.email || '',
      userPhone: user.phone,
      userType: user.userType,
      role: user.role,
    },
  });
});

const login = asyncHandler(async (req, res) => {
  const { userPhone, phone, email, identifier, password, role } = req.body;

  const loginInput = email || identifier || phone || userPhone;

  if (!loginInput) {
    res.status(400);
    throw new Error('Email or phone number is required');
  }

  const trimmedInput = loginInput.trim();

  let user = await User.findOne({
    $or: [
      { email: trimmedInput.toLowerCase() },
      { phone: trimmedInput }
    ]
  }).select('+password');

  if (!user) {
    res.status(401);
    throw new Error('Invalid email/phone or password');
  }

  if (password && !(await user.comparePassword(password))) {
    res.status(401);
    throw new Error('Invalid email/phone or password');
  }

  if (role && user.role !== role) {
    res.status(401);
    throw new Error('Invalid role for this account');
  }

  if (!user.isActive) {
    res.status(403);
    throw new Error('Account is inactive');
  }

  const token = generateToken(user);
  res.json({
    status: true,
    token,
    user: {
      id: user._id,
      name: user.name,
      email: user.email || '',
      role: user.role,
      userName: user.name,
      userEmail: user.email || '',
      userPhone: user.phone || '',
      userType: user.userType || mapRoleToUserType(user.role),
    },
  });
});

const me = asyncHandler(async (req, res) => {
  res.json({
    status: true,
    user: {
      id: req.user._id,
      name: req.user.name,
      email: req.user.email || '',
      role: req.user.role,
      userName: req.user.name,
      userEmail: req.user.email || '',
      userPhone: req.user.phone || '',
      userType: req.user.userType || mapRoleToUserType(req.user.role),
    },
  });
});

module.exports = {
  register,
  login,
  me,
};
