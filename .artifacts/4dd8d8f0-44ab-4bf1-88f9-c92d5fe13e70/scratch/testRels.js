const JSZip = require('jszip');
const fs = require('fs');

async function test() {
  const z1 = await JSZip.loadAsync(fs.readFileSync('reference/Adinn-Direct-Client-format/Adinn-Direct-Client-format.pptx'));
  let rels = await z1.file('ppt/slides/_rels/slide4.xml.rels').async('text');

  const newUrl = 'https://www.google.com/maps?q=9.9303,78.0955';

  let updatedRels = rels.replace(
    /(<Relationship\b[^>]*?)\bTarget="[^"]*"([^>]*?Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"[^>]*\/>)/g,
    `$1Target="${newUrl}"$2`
  );

  console.log('UPDATED RELS:\n', updatedRels);
}

test().catch(console.error);
