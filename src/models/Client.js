const mongoose = require('mongoose');

const clientSchema = new mongoose.Schema(
  {
    name: { type: String, required: true, trim: true },
    phone: { type: String, trim: true },
    email: { type: String, trim: true, lowercase: true },
    location: { type: String, trim: true },
    latitude: Number,
    longitude: Number,
    notes: String,
    isActive: { type: Boolean, default: true },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { timestamps: true }
);

clientSchema.index({ name: 'text', phone: 'text', email: 'text' });

module.exports = mongoose.model('Client', clientSchema);
