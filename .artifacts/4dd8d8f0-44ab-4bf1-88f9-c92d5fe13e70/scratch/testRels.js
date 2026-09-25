const JSZip = require('jszip');
const fs = require('fs');

async function test() {
  const z1 = await JSZip.loadAsync(fs.readFileSync('reference/Adinn-Direct-Client-format/Adinn-Direct-Client-format.pptx'));
  let rels = await z1.file('ppt/slides/_rels/slide4.xml.rels').async('text');
  let slide = await z1.file('ppt/slides/slide4.xml').async('text');

  console.log('--- ORIGINAL RELS ---');
  console.log(rels);

  const newUrl = 'https://www.google.com/maps?q=9.9303,78.0955';

  // Replace Target on any relationship of Type hyperlink
  const updatedRels = rels.replace(
    (/(<Relationship[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"[^>]*Target=")[^"]*(")/g),
    `$1${newUrl}$2`
  );

  // Replace tooltip attribute on <a:hlinkClick> in slideXml
  const updatedSlide = slide.replace(
    /(<a:hlinkClick[^>]*tooltip=")[^"]*(")/g,
    `$1${newUrl}$2`
  );

  console.log('\n--- UPDATED RELS ---');
  console.log(updatedRels);

  console.log('\n--- TOOLTIP IN SLIDE XML BEFORE & AFTER ---');
  console.log('Before:', slide.match(/<a:hlinkClick[^>]*\/>/g));
  console.log('After:', updatedSlide.match(/<a:hlinkClick[^>]*\/>/g));
}

test().catch(console.error);
