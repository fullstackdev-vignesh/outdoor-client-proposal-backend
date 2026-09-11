const mongoose = require('mongoose');

const excelTemplateSchema = new mongoose.Schema(
  {
    name: { type: String, required: true, trim: true },
    description: String,
    version: { type: String, default: '1.0' },
    fileUrl: String,
    status: { type: String, enum: ['active', 'inactive'], default: 'active' },
    usedCount: { type: Number, default: 0 },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { timestamps: true }
);

module.exports = mongoose.model('ExcelTemplate', excelTemplateSchema);
