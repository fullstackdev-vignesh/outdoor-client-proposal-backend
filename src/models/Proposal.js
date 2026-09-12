const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

const proposalSchema = new mongoose.Schema(
  {
    proposalId: { type: String, required: true, unique: true },
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client', required: true },
    sites: [{ type: mongoose.Schema.Types.ObjectId, ref: 'Site', required: true }],
    pptTemplate: { type: mongoose.Schema.Types.ObjectId, ref: 'PPTTemplate' },
    excelTemplate: { type: mongoose.Schema.Types.ObjectId, ref: 'ExcelTemplate' },
    variant: { type: String, default: 'Variant 1' },
    totalAmount: Number,
    gstAmount: Number,
    monthlyAmount: Number,
    status: {
      type: String,
      enum: ['draft', 'generated', 'completed'],
      default: 'draft',
    },
    generatedPptUrl: String,
    generatedExcelUrl: String,
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

proposalSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  this.updatedAt = now;
  next();
});

proposalSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

module.exports = mongoose.model('Proposal', proposalSchema);
