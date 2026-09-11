const mongoose = require('mongoose');

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
  },
  { timestamps: true }
);

module.exports = mongoose.model('Proposal', proposalSchema);
