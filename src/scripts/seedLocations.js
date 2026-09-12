require('dotenv').config();
const mongoose = require('mongoose');
const Location = require('../models/Location');

const DATA = {
  'Andaman and Nicobar Islands': ['Port Blair'],
  'Andhra Pradesh': ['Visakhapatnam', 'Vijayawada', 'Guntur', 'Nellore', 'Tirupati'],
  'Arunachal Pradesh': ['Itanagar'],
  Assam: ['Guwahati', 'Silchar', 'Dibrugarh'],
  Bihar: ['Patna', 'Gaya', 'Bhagalpur', 'Muzaffarpur'],
  Chandigarh: ['Chandigarh'],
  Chhattisgarh: ['Raipur', 'Bhilai', 'Bilaspur'],
  'Dadra and Nagar Haveli and Daman and Diu': ['Daman', 'Silvassa'],
  Delhi: ['New Delhi', 'Dwarka', 'Rohini'],
  Goa: ['Panaji', 'Margao', 'Vasco da Gama'],
  Gujarat: ['Ahmedabad', 'Surat', 'Vadodara', 'Rajkot', 'Gandhinagar'],
  Haryana: ['Faridabad', 'Gurgaon', 'Hisar', 'Panipat', 'Rohtak', 'Karnal'],
  'Himachal Pradesh': ['Shimla', 'Manali', 'Dharamshala'],
  'Jammu and Kashmir': ['Srinagar', 'Jammu'],
  Jharkhand: ['Ranchi', 'Jamshedpur', 'Dhanbad'],
  Karnataka: ['Bengaluru', 'Mysuru', 'Mangaluru', 'Hubballi'],
  Kerala: ['Kochi', 'Thiruvananthapuram', 'Kozhikode', 'Thrissur'],
  Ladakh: ['Leh'],
  'Madhya Pradesh': ['Bhopal', 'Indore', 'Jabalpur', 'Gwalior'],
  Maharashtra: ['Mumbai', 'Pune', 'Nagpur', 'Nashik', 'Thane', 'Aurangabad'],
  Manipur: ['Imphal'],
  Meghalaya: ['Shillong'],
  Mizoram: ['Aizawl'],
  Nagaland: ['Kohima', 'Dimapur'],
  Odisha: ['Bhubaneswar', 'Cuttack', 'Rourkela'],
  Puducherry: ['Puducherry'],
  Punjab: ['Ludhiana', 'Amritsar', 'Jalandhar', 'Patiala', 'Mohali'],
  Rajasthan: ['Jaipur', 'Jodhpur', 'Udaipur', 'Kota', 'Ajmer'],
  Sikkim: ['Gangtok'],
  'Tamil Nadu': ['Chennai', 'Coimbatore', 'Madurai', 'Tiruchirappalli', 'Salem'],
  Telangana: ['Hyderabad', 'Warangal', 'Nizamabad'],
  Tripura: ['Agartala'],
  'Uttar Pradesh': ['Lucknow', 'Kanpur', 'Noida', 'Ghaziabad', 'Agra', 'Varanasi'],
  Uttarakhand: ['Dehradun', 'Haridwar', 'Rishikesh'],
  'West Bengal': ['Kolkata', 'Howrah', 'Durgapur', 'Siliguri'],
};

async function run() {
  await mongoose.connect(process.env.MONGO_URI);
  const existing = await Location.findOne({});
  const update = {};
  for (const [state, cities] of Object.entries(DATA)) {
    update[state] = { $each: cities };
  }
  if (existing) {
    await Location.updateOne({ _id: existing._id }, { $addToSet: update });
  } else {
    await Location.create(DATA);
  }
  console.log('Locations seeded/merged into the location collection.');
  await mongoose.disconnect();
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
