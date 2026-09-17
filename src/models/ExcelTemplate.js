const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const excelTemplateSchema = new mongoose.Schema(
  {
    name: { type: String, required: true, trim: true },
    description: String,
    version: { type: String, default: '1.0' },
    fileUrl: String,
    // Which entry in config/excelTemplateConfigs.js drives generation for this uploaded
    // file's column layout — empty/unset falls back to the 'generic' (Adinn) mapping.
    formatKey: { type: String, default: '', trim: true },
    status: { type: String, enum: ['active', 'inactive'], default: 'active' },
    usedCount: { type: Number, default: 0 },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

excelTemplateSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  this.updatedAt = now;
  next();
});

excelTemplateSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

module.exports = mongoose.model('ExcelTemplate', excelTemplateSchema);
