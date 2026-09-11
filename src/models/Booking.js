const mongoose = require('mongoose');

const bookingSchema = new mongoose.Schema(
  {
    bookingId: { type: String, required: true, unique: true },
    client: { type: mongoose.Schema.Types.ObjectId, ref: 'Client', required: true },
    sites: [{ type: mongoose.Schema.Types.ObjectId, ref: 'Site', required: true }],
    startDate: { type: Date, required: true },
    endDate: { type: Date, required: true },
    amount: Number,
    gstAmount: Number,
    totalAmount: Number,
    notes: String,
    status: {
      type: String,
      enum: ['active', 'completed', 'cancelled'],
      default: 'active',
    },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
  },
  { timestamps: true }
);

module.exports = mongoose.model('Booking', bookingSchema);
