const JSZip = require('jszip');
const fs = require('fs');
const { PptxTemplate } = require('./src/utils/pptxTemplateEngine');

async function testGenerator(templateFile) {
  console.log('\n========================================');
  console.log('Testing generation with:', templateFile);
  console.log('========================================');

  const buf = fs.readFileSync(`./assets/proposal-templates/${templateFile}`);
  const zip = await JSZip.loadAsync(buf);
  const tpl = new PptxTemplate(zip);

  const slideFiles = await tpl.getSlideFiles();
  console.log('Original slideFiles:', slideFiles);

  // Read slide texts to detect roles
  const slideTexts = {};
  for (const sf of slideFiles) {
    const xml = await tpl.readText(`ppt/slides/${sf}.xml`);
    const texts = [...xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map((m) => m[1].trim()).filter(Boolean);
    slideTexts[sf] = texts;
  }

  const coverTpl = slideFiles[0];
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

  console.log('Detected roles: Cover =', coverTpl, '| CityDivider =', cityDividerTpl, '| SiteSpec =', siteSpecTpl, '| ThankYou =', thankYouTpl);

  // Mock requested sites from req/res proposal
  const mockSites = [
    { _id: 's1', city: 'Chennai', state: 'Tamil Nadu', location: 'Gemini flyover towards Nungambakkam', mediaType: 'Digital', illumination: 'LED', width: 30, height: 20 },
    { _id: 's2', city: 'Chennai', state: 'Tamil Nadu', location: 'Vadapalani Signal', mediaType: 'Digital', illumination: 'LED', width: 15, height: 12 },
    { _id: 's3', city: 'Madurai', state: 'Tamil Nadu', location: 'Goripalayam Junction', mediaType: 'Hoarding', illumination: 'Frontlit', width: 40, height: 20 },
  ];

  // Group sites by city
  const sitesByCity = {};
  for (const site of mockSites) {
    const city = site.city || 'Other';
    if (!sitesByCity[city]) sitesByCity[city] = [];
    sitesByCity[city].push(site);
  }

  const generatedSlideOrder = [coverTpl];

  for (const [city, citySites] of Object.entries(sitesByCity)) {
    // 1. Clone City Divider Slide if cityDividerTpl exists
    if (cityDividerTpl) {
      const stateName = citySites[0]?.state || 'Tamil Nadu';
      const cityDividerBase = await tpl.cloneSlide(
        cityDividerTpl,
        {
          textReplacements: [
            ['Chennai', city],
            ['Madurai', city],
            ['Tamil Nadu', stateName],
          ],
        },
        { city, state: stateName }
      );
      generatedSlideOrder.push(cityDividerBase);
    }

    // 2. Clone Site Spec Slide for each requested site
    for (const site of citySites) {
      const specBase = await tpl.cloneSlide(
        siteSpecTpl,
        {
          textReplacements: [
            ['Chennai', site.city],
            ['Tamil Nadu', site.state],
            ['Digital', site.mediaType],
            ['LED', site.illumination],
            ['30', String(site.width)],
            ['20', String(site.height)],
          ],
        },
        site
      );
      generatedSlideOrder.push(specBase);
    }
  }

  // Push thank you slide at the very end
  if (thankYouTpl && thankYouTpl !== coverTpl) {
    generatedSlideOrder.push(thankYouTpl);
  }

  console.log('Final Generated Slide Order:', generatedSlideOrder);

  // Set final slide order in presentation.xml
  await tpl.setFinalSlideOrder(generatedSlideOrder);

  const finalPresXml = await tpl.readText('ppt/presentation.xml');
  console.log('Final <p:sldIdLst> snippet:', finalPresXml.match(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/)?.[0]);
}

async function run() {
  await testGenerator('master.pptx');
  await testGenerator('template-1789471038153-fb1d78ae.pptx');
}

run().catch(console.error);
