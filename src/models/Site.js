const mongoose = require('mongoose');

const bookingInfoSchema = new mongoose.Schema(
  {
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client' },
    booking: { type: mongoose.Schema.Types.ObjectId, ref: 'Booking' },
    bookingRef: String,
    startDate: Date,
    endDate: Date,
    amount: Number,
    bookedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

const blockInfoSchema = new mongoose.Schema(
  {
    reason: String,
    notes: String,
    blockedDate: { type: Date, default: Date.now },
    blockedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { _id: false }
);

const siteSchema = new mongoose.Schema(
  {
    mediaId: { type: String, required: true, unique: true, trim: true },
    mediaName: { type: String, required: true, trim: true },
    mediaType: { type: String, required: true, trim: true },
    state: { type: String, required: true, trim: true },
    city: { type: String, required: true, trim: true },
    location: { type: String, trim: true },
    latitude: Number,
    longitude: Number,
    width: Number,
    height: Number,
    sizeUnit: { type: String, default: 'ft' },
    amount: Number,
    gstAmount: Number,
    monthlyAmount: Number,
    image: String,
    isActive: { type: Boolean, default: true },
    mediaStatus: {
      type: String,
      enum: ['available', 'booked', 'blocked'],
      default: 'available',
      index: true,
    },
    bookingInfo: bookingInfoSchema,
    blockInfo: blockInfoSchema,
    assignedTL: { type: mongoose.Schema.Types.ObjectId, ref: 'User', default: null },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { timestamps: true }
);

siteSchema.index({ mediaName: 'text', location: 'text', city: 'text', state: 'text' });

module.exports = mongoose.model('Site', siteSchema);
