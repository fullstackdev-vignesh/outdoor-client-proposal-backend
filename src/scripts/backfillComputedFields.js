require('dotenv').config();
const mongoose = require('mongoose');
const Site = require('../models/Site');

async function run() {
  await mongoose.connect(process.env.MONGO_URI);
  const sites = await Site.find({
    $or: [{ autoSize: { $exists: false } }, { totalCost: { $exists: false } }],
  });

  let updated = 0;
  for (const site of sites) {
    // save() re-runs the model's pre-save hook, which recomputes autoSize/totalCost.
    await site.save();
    updated += 1;
  }

  console.log(`Backfilled autoSize/totalCost on ${updated} site(s).`);
  await mongoose.disconnect();
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
