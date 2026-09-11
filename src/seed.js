require('dotenv').config();
const connectDB = require('./config/db');
const User = require('./models/User');
const Site = require('./models/Site');
const Client = require('./models/Client');

const CITIES = [
  ['Maharashtra', 'Mumbai'],
  ['Delhi', 'New Delhi'],
  ['Karnataka', 'Bengaluru'],
  ['Tamil Nadu', 'Chennai'],
  ['West Bengal', 'Kolkata'],
];
const MEDIA_TYPES = ['Hoarding', 'Digital Hoarding', 'Unipole', 'Gantry', 'Bus Shelter'];

async function run() {
  await connectDB();

  await Promise.all([User.deleteMany({}), Site.deleteMany({}), Client.deleteMany({})]);

  const admin = await User.create({
    name: 'Admin User',
    email: 'admin@outdoor.com',
    password: 'Admin@123',
    role: 'admin',
  });
  const tl = await User.create({
    name: 'Team Leader',
    email: 'tl@outdoor.com',
    password: 'Tl@12345',
    role: 'tl',
  });
  const user = await User.create({
    name: 'Regular User',
    email: 'user@outdoor.com',
    password: 'User@123',
    role: 'user',
  });

  const clients = await Client.insertMany(
    Array.from({ length: 8 }).map((_, i) => ({
      name: `Client ${i + 1} Pvt Ltd`,
      phone: `98765432${10 + i}`,
      email: `client${i + 1}@example.com`,
      location: CITIES[i % CITIES.length][1],
      createdBy: admin._id,
    }))
  );

  const sites = [];
  for (let i = 1; i <= 60; i += 1) {
    const [state, city] = CITIES[i % CITIES.length];
    const statusRoll = i % 5;
    const mediaStatus = statusRoll === 0 ? 'booked' : statusRoll === 1 ? 'blocked' : 'available';
    const site = {
      mediaId: `MEDIA-${String(i).padStart(4, '0')}`,
      mediaName: `${MEDIA_TYPES[i % MEDIA_TYPES.length]} ${city} #${i}`,
      mediaType: MEDIA_TYPES[i % MEDIA_TYPES.length],
      state,
      city,
      location: `${city} Highway Junction ${i}`,
      latitude: 19 + Math.random() * 10,
      longitude: 72 + Math.random() * 10,
      width: 20,
      height: 10,
      amount: 50000 + i * 500,
      gstAmount: (50000 + i * 500) * 0.18,
      monthlyAmount: 50000 + i * 500,
      isActive: true,
      mediaStatus,
      createdBy: admin._id,
    };
    if (mediaStatus === 'booked') {
      site.bookingInfo = {
        client: clients[i % clients.length]._id,
        bookingRef: `BK-SEED-${i}`,
        startDate: new Date(),
        endDate: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
        amount: site.monthlyAmount,
        bookedBy: admin._id,
      };
    }
    if (mediaStatus === 'blocked') {
      site.blockInfo = {
        reason: 'Under maintenance',
        notes: 'Structural repair in progress',
        blockedDate: new Date(),
        blockedBy: admin._id,
      };
    }
    sites.push(site);
  }
  await Site.insertMany(sites);

  console.log('Seed complete:');
  console.log('  admin@outdoor.com / Admin@123');
  console.log('  tl@outdoor.com / Tl@12345');
  console.log('  user@outdoor.com / User@123');
  process.exit(0);
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
