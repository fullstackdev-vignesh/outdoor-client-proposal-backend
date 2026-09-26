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
    // generatedPptUrl always holds the MOST RECENT generation regardless of mode (kept for
    // backward compatibility with any existing caller that only reads this one field); the two
    // mode-specific fields below let the UI offer separate persistent "Download" links for each
    // mode without one regenerate overwriting the other's URL.
    generatedPptUrl: String,
    generatedPptWithLocationUrl: String,
    generatedPptWithoutLocationUrl: String,
    generatedExcelUrl: String,
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedBy: { type: String, default: 'System' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

proposalSchema.pre('save', function (next) {
  const now = nowIST();
  if (!this.createdAt) this.createdAt = now;
  this.updatedAt = now;
  if (this.$locals.currentUserName) this.updatedBy = this.$locals.currentUserName;
  next();
});

proposalSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

module.exports = mongoose.model('Proposal', proposalSchema);
