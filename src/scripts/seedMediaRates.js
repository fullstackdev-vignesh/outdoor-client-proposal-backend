// Fills the Rate Master with the rate sheet's defaults the first time the server starts.
// Runs only while the collection is empty, so rates edited or deleted in the app are never
// overwritten. Runs automatically on server start (src/server.js).
const MediaRate = require('../models/MediaRate');

const LIT = { printingBackLit: 25, printingOther: 13 };
const PLAIN = { printingBackLit: 13, printingOther: 13 };

const DEFAULT_RATES = [
  { mediaType: 'ROB', ...LIT, mountingType: 'fixed', mountingAmount: 10000 },
  { mediaType: 'Pole Kiosk', ...LIT, mountingType: 'perQuantity', mountingAmount: 300 },
  { mediaType: 'Bus Shelter', ...LIT, mountingType: 'fixed', mountingAmount: 2000 },
  { mediaType: 'Signal Post', ...LIT, mountingType: 'perSqFt', mountingAmount: 20 },
  { mediaType: 'Branding Board', ...LIT, mountingType: 'perSqFt', mountingAmount: 5 },
  { mediaType: 'Gantry', ...LIT, mountingType: 'fixed', mountingAmount: 4000 },
  { mediaType: 'Police booth', ...LIT, mountingType: 'perSqFt', mountingAmount: 15 },
  { mediaType: 'Lamp Post', ...LIT, mountingType: 'perQuantity', mountingAmount: 300 },
  { mediaType: 'Center Median', ...LIT, mountingType: 'perQuantity', mountingAmount: 300 },
  { mediaType: 'Top panels', ...LIT, mountingType: 'perSqFt', mountingAmount: 5 },
  { mediaType: 'Police Umbrella', ...LIT, mountingType: 'perSqFt', mountingAmount: 15 },
  { mediaType: 'Transformer', ...LIT, mountingType: 'perSqFt', mountingAmount: 20 },
  { mediaType: 'Hoarding', ...PLAIN, mountingType: 'perSqFt', mountingAmount: 5 },
  { mediaType: 'Unipole', ...PLAIN, mountingType: 'perSqFt', mountingAmount: 7 },
  { mediaType: 'Wall Graphics', ...PLAIN, mountingType: 'perSqFt', mountingAmount: 5 },
  { mediaType: 'Wall Frame', ...PLAIN, mountingType: 'perSqFt', mountingAmount: 5 },
  { mediaType: 'LED Hoarding', printingBackLit: 0, printingOther: 0, mountingType: 'perSqFt', mountingAmount: 0 },
  { mediaType: 'Led', printingBackLit: 0, printingOther: 0, mountingType: 'perSqFt', mountingAmount: 0 },
  { mediaType: 'Wall Hoarding', ...PLAIN, mountingType: 'perSqFt', mountingAmount: 5 },
  { mediaType: 'Wall Wrap', ...PLAIN, mountingType: 'perSqFt', mountingAmount: 5 },
];

async function seedMediaRates() {
  if (await MediaRate.estimatedDocumentCount()) return;
  await MediaRate.insertMany(DEFAULT_RATES);
  console.log(`[rateMaster] seeded ${DEFAULT_RATES.length} default media rates`);
}

module.exports = { seedMediaRates, DEFAULT_RATES };
