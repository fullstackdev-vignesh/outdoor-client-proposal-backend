const asyncHandler = require('express-async-handler');
const User = require('../models/User');

const getUsers = asyncHandler(async (req, res) => {
  const filter = {};
  if (req.query.role) filter.role = req.query.role;
  if (req.query.search) filter.name = new RegExp(req.query.search, 'i');
  const users = await User.find(filter).sort({ createdAt: -1 });
  res.json(users);
});

const getUser = asyncHandler(async (req, res) => {
  const user = await User.findById(req.params.id);
  if (!user) {
    res.status(404);
    throw new Error('User not found');
  }
  res.json(user);
});

const createUser = asyncHandler(async (req, res) => {
  const exists = await User.findOne({ email: req.body.email });
  if (exists) {
    res.status(400);
    throw new Error('Email already in use');
  }
  const user = await User.create(req.body);
  res.status(201).json(user);
});

const updateUser = asyncHandler(async (req, res) => {
  const user = await User.findById(req.params.id);
  if (!user) {
    res.status(404);
    throw new Error('User not found');
  }
  const { password, ...rest } = req.body;
  Object.assign(user, rest);
  if (password) user.password = password;
  await user.save();
  res.json(user);
});

const deleteUser = asyncHandler(async (req, res) => {
  const user = await User.findById(req.params.id);
  if (!user) {
    res.status(404);
    throw new Error('User not found');
  }
  await user.deleteOne();
  res.json({ message: 'User deleted' });
});

module.exports = { getUsers, getUser, createUser, updateUser, deleteUser };
