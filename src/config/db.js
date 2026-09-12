const mongoose = require('mongoose');
const dns = require('dns');

dns.setServers(['8.8.8.8', '1.1.1.1']);

async function connectDB() {
  const uri = process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/outdoor';
  await mongoose.connect(uri);
  console.log('MongoDB connected:', uri);
}

module.exports = connectDB;
