const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const userSchema = new mongoose.Schema(
  {
    name: { type: String, required: true, trim: true },
    email: { type: String, lowercase: true, trim: true, sparse: true },
    password: { type: String, minlength: 6, select: false },
    role: { type: String, enum: ['admin', 'tl', 'user', 'bd'], default: 'user' },
    userType: { type: Number, default: 1 },
    assignedTL: { type: mongoose.Schema.Types.ObjectId, ref: 'User', default: null },
    assignedSites: [{ type: mongoose.Schema.Types.ObjectId, ref: 'Site' }],
    isActive: { type: Boolean, default: true },
    phone: { type: String, trim: true, sparse: true },
    otp: { type: String, select: false },
    otpExpires: { type: Date, select: false },
    isPhoneVerified: { type: Boolean, default: false },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

userSchema.pre('save', async function hashPassword(next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  this.updatedAt = now;

  if (!this.isModified('password') || !this.password) return next();
  this.password = await bcrypt.hash(this.password, 10);
  next();
});

userSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

userSchema.methods.comparePassword = function comparePassword(candidate) {
  if (!this.password) return false;
  return bcrypt.compare(candidate, this.password);
};

module.exports = mongoose.model('User', userSchema);
