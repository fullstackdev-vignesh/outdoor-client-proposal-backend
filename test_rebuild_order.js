const JSZip = require('jszip');
const fs = require('fs');

async function testRebuild(filename) {
  console.log('\n========================================');
  console.log('Testing template:', filename);
  console.log('========================================');

  const buf = fs.readFileSync(`./assets/proposal-templates/${filename}`);
  const zip = await JSZip.loadAsync(buf);

  // 1. Find all slide files in zip
  const slideFiles = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .map((f) => f.replace('ppt/slides/', '').replace('.xml', ''));

  slideFiles.sort((a, b) => parseInt(a.replace('slide', '')) - parseInt(b.replace('slide', '')));

  console.log('Original slide files in zip:', slideFiles);

  // Read text of each slide to identify roles
  const slideTexts = {};
  for (const sf of slideFiles) {
    const xml = await zip.file(`ppt/slides/${sf}.xml`).async('string');
    const texts = [...xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map((m) => m[1].trim()).filter(Boolean);
    slideTexts[sf] = texts;
    console.log(`[${sf}]`, texts.join(' | '));
  }

  // Identify roles:
  const coverTpl = slideFiles[0]; // slide1
  let cityDividerTpl = null;
  let siteSpecTpl = null;
  let thankYouTpl = null;

  for (const sf of slideFiles) {
    const txts = slideTexts[sf].join(' ').toLowerCase();
    if (!cityDividerTpl && (sf === 'slide2' || txts.includes('chennai') || txts.includes('madurai'))) {
      cityDividerTpl = sf;
    }
    if (!siteSpecTpl && (txts.includes('duration') || txts.includes('media type') || txts.includes('illumination') || txts.includes('flyover') || txts.includes('bridge') || txts.includes('hoarding') || txts.includes('unipole') || txts.includes('40x30'))) {
      siteSpecTpl = sf;
    }
    if (txts.includes('thank you') || txts.includes('thanks') || txts.includes('awaiting your approval')) {
      thankYouTpl = sf;
    }
  }

  console.log('\nRole Mapping:');
  console.log('  Cover:', coverTpl);
  console.log('  City Divider:', cityDividerTpl);
  console.log('  Site Spec:', siteSpecTpl);
  console.log('  Thank You:', thankYouTpl);
}

async function run() {
  await testRebuild('master.pptx');
  await testRebuild('template-1789471038153-fb1d78ae.pptx');
}

run().catch(console.error);
