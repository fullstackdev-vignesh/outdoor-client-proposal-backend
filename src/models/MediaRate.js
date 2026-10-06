const mongoose = require('mongoose');

const IST_OFFSET_MS = 330 * 60000;
const nowIST = () => new Date(Date.now() + IST_OFFSET_MS);

// How a media type's Mounting Cost is worked out (always multiplied by the site's Quantity):
//   perSqFt     -> mountingAmount × Total Sq.ft
//   fixed       -> mountingAmount (flat per unit)
//   perQuantity -> mountingAmount per unit (same maths as fixed, kept separate to match the rate sheet)
const MOUNTING_TYPES = ['perSqFt', 'fixed', 'perQuantity'];

// Rate Master — Printing / Mounting rates per Media Type. The site form uses these to fill
// Printing Cost and Mounting Cost.
const mediaRateSchema = new mongoose.Schema(
  {
    mediaType: { type: String, required: true, trim: true, unique: true },
    // Printing rate per sq.ft for Back Lit sites.
    printingBackLit: { type: Number, required: true, min: 0, default: 0 },
    // Printing rate per sq.ft for Front Lit / Non Lit sites.
    printingOther: { type: Number, required: true, min: 0, default: 0 },
    mountingType: { type: String, enum: MOUNTING_TYPES, required: true, default: 'perSqFt' },
    mountingAmount: { type: Number, required: true, min: 0, default: 0 },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: nowIST },
    updatedAt: { type: Date, default: nowIST },
  },
  { timestamps: false }
);

mediaRateSchema.pre('save', function (next) {
  this.updatedAt = nowIST();
  next();
});

mediaRateSchema.pre(['updateOne', 'findOneAndUpdate', 'updateMany'], function (next) {
  this.set({ updatedAt: nowIST() });
  next();
});

module.exports = mongoose.model('MediaRate', mediaRateSchema);
module.exports.MOUNTING_TYPES = MOUNTING_TYPES;
