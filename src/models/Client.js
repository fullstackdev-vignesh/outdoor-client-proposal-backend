const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const clientSchema = new mongoose.Schema(
  {
    customerType: { type: String, enum: ['client', 'agency'], default: 'client' },
    name: { type: String, required: true, trim: true },
    phone: { type: String, trim: true },
    email: { type: String, trim: true, lowercase: true },
    location: { type: String, trim: true },
    latitude: { type: Number, min: -90, max: 90 },
    longitude: { type: Number, min: -180, max: 180 },
    agencyComm: { type: Number, min: 0, max: 100 },
    // Percentage (e.g. 18 => 18%), not a GST registration number — drives the optional GST
    // column in generated proposal Excel exports (see excelTemplateConfigs.js).
    gst: { type: Number, min: 0, max: 100 },
    notes: String,
    isActive: { type: Boolean, default: true },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

clientSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  this.updatedAt = now;
  next();
});

clientSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

clientSchema.index({ name: 'text', phone: 'text', email: 'text' });

module.exports = mongoose.model('Client', clientSchema);
