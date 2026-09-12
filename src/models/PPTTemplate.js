const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const pptTemplateSchema = new mongoose.Schema(
  {
    name: { type: String, required: true, trim: true },
    description: String,
    version: { type: String, default: '1.0' },
    variant: { type: String, default: 'Standard' },
    fileUrl: String,
    status: { type: String, enum: ['active', 'inactive'], default: 'active' },
    usedCount: { type: Number, default: 0 },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

pptTemplateSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  this.updatedAt = now;
  next();
});

pptTemplateSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

module.exports = mongoose.model('PPTTemplate', pptTemplateSchema);
