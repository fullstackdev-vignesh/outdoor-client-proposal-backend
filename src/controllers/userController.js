const asyncHandler = require('express-async-handler');
const User = require('../models/User');

const mapRoleToUserType = (r) => {
  if (r === 'bd') return 4;
  if (r === 'admin') return 3;
  if (r === 'tl') return 2;
  return 1;
};

const getUsers = asyncHandler(async (req, res) => {
  const filter = {};
  if (req.query.role) filter.role = req.query.role;
  if (req.query.search) filter.name = new RegExp(req.query.search, 'i');

  if (req.user.role === 'tl') {
    filter.role = 'user';
    filter.assignedTL = req.user._id;
  }

  const users = await User.find(filter).sort({ updatedAt: -1, _id: -1 });
  res.json(users);
});

const getUser = asyncHandler(async (req, res) => {
  const user = await User.findById(req.params.id);
  if (!user) {
    res.status(404);
    throw new Error('User not found');
  }

  if (req.user.role === 'tl' && (user.role !== 'user' || user.assignedTL?.toString() !== req.user._id.toString())) {
    res.status(403);
    throw new Error('Not authorized to view this user');
  }

  res.json(user);
});

const createUser = asyncHandler(async (req, res) => {
  const { email, phone, role, password, confirmPassword, name, ...rest } = req.body;
  const creatorRole = req.user.role;
  const targetRole = role || 'user';

  if (confirmPassword !== undefined && password !== confirmPassword) {
    res.status(400);
    throw new Error('Password and Confirm Password do not match');
  }

  if (creatorRole === 'tl') {
    if (targetRole !== 'user') {
      res.status(403);
      throw new Error('TL can only create User accounts (cannot create TL or BD)');
    }
  } else if (creatorRole !== 'admin') {
    res.status(403);
    throw new Error('You do not have permission to create accounts');
  }

  if (email) {
    const exists = await User.findOne({ email: email.trim().toLowerCase() });
    if (exists) {
      res.status(400);
      throw new Error('Email already in use');
    }
  }

  if (phone) {
    const existsPhone = await User.findOne({ phone: phone.trim() });
    if (existsPhone) {
      res.status(400);
      throw new Error('Phone number already in use');
    }
  }

  const userType = mapRoleToUserType(targetRole);

  const userData = {
    name,
    email: email ? email.trim().toLowerCase() : undefined,
    phone: phone ? phone.trim() : undefined,
    password,
    role: targetRole,
    userType,
    assignedTL: creatorRole === 'tl' ? req.user._id : req.body.assignedTL || null,
    ...rest,
  };

  const user = await User.create(userData);
  res.status(201).json(user);
});

const updateUser = asyncHandler(async (req, res) => {
  const user = await User.findById(req.params.id);
  if (!user) {
    res.status(404);
    throw new Error('User not found');
  }

  if (req.user.role === 'tl') {
    if (user.role !== 'user' || user.assignedTL?.toString() !== req.user._id.toString()) {
      res.status(403);
      throw new Error('Not authorized to update this user');
    }
    if (req.body.role && req.body.role !== 'user') {
      res.status(403);
      throw new Error('TL cannot change user role');
    }
  }

  const { password, confirmPassword, ...rest } = req.body;
  if (password) {
    if (confirmPassword !== undefined && password !== confirmPassword) {
      res.status(400);
      throw new Error('Password and Confirm Password do not match');
    }
    user.password = password;
  }

  Object.assign(user, rest);

  if (rest.role) {
    user.userType = mapRoleToUserType(rest.role);
  }

  await user.save();
  res.json(user);
});

const deleteUser = asyncHandler(async (req, res) => {
  const user = await User.findById(req.params.id);
  if (!user) {
    res.status(404);
    throw new Error('User not found');
  }

  if (req.user.role === 'tl') {
    if (user.role !== 'user' || user.assignedTL?.toString() !== req.user._id.toString()) {
      res.status(403);
      throw new Error('Not authorized to delete this user');
    }
  }

  await user.deleteOne();
  res.json({ message: 'User deleted' });
});

module.exports = { getUsers, getUser, createUser, updateUser, deleteUser };
