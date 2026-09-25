const asyncHandler = require('express-async-handler');
const bcrypt = require('bcryptjs');
const User = require('../models/User');
const generateToken = require('../utils/generateToken');

const STAFF_REGISTER_PASSWORD = process.env.STAFF_REGISTER_PASSWORD || 'Adinn@123';
const TEAMHEAD_REGISTER_PASSWORD = process.env.TEAMHEAD_REGISTER_PASSWORD || 'Adinn@1234';
const BD_REGISTER_PASSWORD = process.env.BD_REGISTER_PASSWORD  || 'Adinn@12345';
const ADMIN_REGISTER_PASSWORD = process.env.ADMIN_REGISTER_PASSWORD || 'Adinn@123456';

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
  const userName = (req.body.userName || req.body.name || '').trim();
  let userEmail = (req.body.userEmail || req.body.email || '').trim().toLowerCase();
  const userPhone = (req.body.userPhone || req.body.phone || '').trim();
  const pin = String(req.body.password || req.body.registerPassword || req.body.pin || req.body.userPin || '').trim();
  const confirmPin = String(req.body.confirmPassword || req.body.confirmPin || '').trim();
  const registerPassword = (req.body.registerPassword || req.body.adminKey || req.body.rolePassword || '').trim();

  let role = req.body.role;
  let userType = req.body.userType !== undefined ? Number(req.body.userType) : undefined;

  if (role) {
    userType = mapRoleToUserType(role);
  } else if (userType !== undefined) {
    role = mapUserTypeToRole(userType);
  } else {
    role = 'user';
    userType = 1;
  }

  if (!userName) {
    res.status(400);
    throw new Error('Full Name is required');
  }

  if (!userPhone || !/^\d{10}$/.test(userPhone)) {
    res.status(400);
    throw new Error('Valid 10-digit Phone Number is required');
  }

  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  if (!userEmail || !emailRegex.test(userEmail)) {
    res.status(400);
    throw new Error('Valid Email address is required');
  }

  if (!pin || !/^\d{4}$/.test(pin)) {
    res.status(400);
    throw new Error('4-digit PIN is required');
  }

  if (confirmPin && pin !== confirmPin) {
    res.status(400);
    throw new Error('PIN and Confirm PIN do not match');
  }

  if (!registerPassword) {
    res.status(400);
    throw new Error(`Registration password is required for ${role.toUpperCase()}`);
  }

  if (role === 'user' && registerPassword !== STAFF_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid User registration password');
  } else if (role === 'tl' && registerPassword !== TEAMHEAD_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid TL registration password');
  } else if (role === 'bd' && registerPassword !== BD_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid BD registration password');
  } else if (role === 'admin' && registerPassword !== ADMIN_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid Admin registration password');
  }

  const existingPhoneUser = await User.findOne({ phone: userPhone });
  if (existingPhoneUser) {
    res.status(400);
    throw new Error('Phone number is already registered');
  }

  const existingEmailUser = await User.findOne({ email: userEmail });
  if (existingEmailUser) {
    res.status(400);
    throw new Error('Email address is already registered');
  }

  const user = await User.create({
    name: userName,
    email: userEmail,
    phone: userPhone,
    password: pin,
    userType,
    role,
    registerPassword,
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
  const pin = String(password || req.body.pin || req.body.userPin || req.body.loginPin || '').trim();

  const loginInput = email || identifier || phone || userPhone;

  if (!loginInput) {
    res.status(400);
    throw new Error('Email or phone number is required');
  }

  if (!pin) {
    res.status(400);
    throw new Error('4-digit PIN is required');
  }

  const trimmedInput = String(loginInput).trim();

  let user = await User.findOne({
    $or: [{ email: trimmedInput.toLowerCase() }, { phone: trimmedInput }],
  }).select('+password');

  if (!user) {
    res.status(401);
    throw new Error('Invalid email/phone or PIN');
  }

  if (!(await user.comparePassword(pin))) {
    res.status(401);
    throw new Error('Invalid email/phone or PIN');
  }

  if (role && user.role !== role) {
    res.status(401);
    throw new Error(`This account is registered as ${user.role.toUpperCase()}, not ${role.toUpperCase()}`);
  }

  if (!user.isActive) {
    res.status(403);
    throw new Error('Account is inactive');
  }

  user.lastLogin = new Date();
  await user.save();

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

// FORGOT PIN STEP 1: Generates temporary 4-digit PIN
// Checks MAIL_MODE: 'development' (or 'dev') returns pin in API response; 'production' (or 'live') sends email via PHP Mail API.
const forgotPin = asyncHandler(async (req, res) => {
  const input = (req.body.userEmail || req.body.email || req.body.userPhone || req.body.phone || req.body.identifier || '').trim();

  if (!input) {
    res.status(400);
    throw new Error('Registered phone number or email is required');
  }

  const user = await User.findOne({
    $or: [{ email: input.toLowerCase() }, { phone: input }],
  });

  if (!user) {
    res.status(404);
    throw new Error('No user account found with this phone number or email');
  }

  const targetEmail = user.email || (input.includes('@') ? input : null);

  // Generate random 4-digit temporary PIN (1000 - 9999)
  const generatedPin = String(Math.floor(1000 + Math.random() * 9000));

  user.password = generatedPin;
  await user.save();

  const mailMode = (process.env.MAIL_MODE || 'development').toLowerCase();

  // In Development mode (or 'dev'): Skip PHP Mail API call & return temporary PIN directly in response
  if (mailMode === 'development' || mailMode === 'dev') {
    return res.json({
      status: true,
      message: `Temporary PIN generated successfully.`,
      pin: generatedPin,
      email: targetEmail || '',
      phone: user.phone,
    });
  }

  // In Production mode (or 'live'): Send email via PHP Mail API
  if (!targetEmail) {
    res.status(400);
    throw new Error('No registered email address found for this user account');
  }

  const phpMailUrl = process.env.PHP_MAIL_URL || 'https://adinndigital.com/api/outdoorproposal/forgotPinMail.php';
  const mailPayload = {
    mailtype: 'forgotpin',
    to: [targetEmail],
    data: {
      userName: user.name || 'User',
      email: targetEmail,
      pin: generatedPin,
    },
  };

  try {
    const response = await fetch(phpMailUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(mailPayload),
    });

    const result = await response.json().catch(() => ({}));
    if (response.ok && (result.status === 'success' || result.status === true)) {
      res.json({
        status: true,
        message: 'A temporary 4-digit PIN has been sent to your registered email address.',
        email: targetEmail,
        phone: user.phone,
      });
    } else {
      res.json({
        status: true,
        message: 'A temporary 4-digit PIN has been set for your account.',
        email: targetEmail,
        phone: user.phone,
      });
    }
  } catch (err) {
    console.error('PHP Mail Request Exception:', err.message);
    res.json({
      status: true,
      message: 'A temporary 4-digit PIN has been generated for your account.',
      email: targetEmail,
      phone: user.phone,
    });
  }
});

// FORGOT PIN STEP 2: Verifies temporary PIN & resets to new MPIN
const resetPin = asyncHandler(async (req, res) => {
  const input = (req.body.userEmail || req.body.email || req.body.userPhone || req.body.phone || req.body.identifier || '').trim();
  const tempPin = String(req.body.temporaryPin || req.body.currentPin || req.body.tempPin || req.body.oldPin || '').trim();
  const newPin = String(req.body.newPin || req.body.pin || req.body.password || '').trim();
  const confirmPin = String(req.body.confirmPin || req.body.confirmPassword || '').trim();

  if (!input) {
    res.status(400);
    throw new Error('Phone number or email is required');
  }

  if (!newPin || !/^\d{4}$/.test(newPin)) {
    res.status(400);
    throw new Error('New MPIN must be exactly 4 numeric digits');
  }

  if (confirmPin && newPin !== confirmPin) {
    res.status(400);
    throw new Error('New MPIN and Confirm MPIN do not match');
  }

  const user = await User.findOne({
    $or: [{ email: input.toLowerCase() }, { phone: input }],
  }).select('+password');

  if (!user) {
    res.status(404);
    throw new Error('User account not found');
  }

  // If temporary pin was provided, verify it first
  if (tempPin) {
    const isTempValid = await user.comparePassword(tempPin);
    if (!isTempValid) {
      res.status(400);
      throw new Error('Invalid temporary 4-digit PIN');
    }
  }

  user.password = newPin;
  await user.save();

  res.json({
    status: true,
    message: 'MPIN updated successfully. Please log in with your new MPIN.',
  });
});

const forgotPinVerify = asyncHandler(async (req, res) => {
  const { userPhone, phone, email, registerPassword, role } = req.body;
  const identifier = (phone || userPhone || email || '').trim();

  if (!identifier) {
    res.status(400);
    throw new Error('Phone number or email is required');
  }

  if (!registerPassword) {
    res.status(400);
    throw new Error('Registration password is required');
  }

  const user = await User.findOne({
    $or: [{ email: identifier.toLowerCase() }, { phone: identifier }],
  });

  if (!user) {
    res.status(404);
    throw new Error('User not found with provided phone number / email');
  }

  const targetRole = role || user.role;

  if (targetRole === 'user' && registerPassword !== STAFF_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid User registration password');
  } else if (targetRole === 'tl' && registerPassword !== TEAMHEAD_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid TL registration password');
  } else if (targetRole === 'bd' && registerPassword !== BD_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid BD registration password');
  } else if (targetRole === 'admin' && registerPassword !== ADMIN_REGISTER_PASSWORD) {
    res.status(400);
    throw new Error('Invalid Admin registration password');
  }

  res.json({
    status: true,
    message: 'Details verified successfully. You can now reset your MPIN.',
    userPhone: user.phone,
    userEmail: user.email,
    role: user.role,
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
  forgotPin,
  forgotPinVerify,
  resetPin,
  me,
};
