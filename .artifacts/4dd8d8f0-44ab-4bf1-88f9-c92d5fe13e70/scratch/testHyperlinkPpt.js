const JSZip = require('jszip');
const fs = require('fs');

async function test() {
  const z1 = await JSZip.loadAsync(fs.readFileSync('reference/Adinn-Direct-Client-format/Adinn-Direct-Client-format.pptx'));
  let slide = await z1.file('ppt/slides/slide4.xml').async('text');
  let rels = await z1.file('ppt/slides/_rels/slide4.xml.rels').async('text');

  const locationUrl = 'https://www.google.com/maps?q=9.9303,78.0955';
  const escapedUrl = locationUrl.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  rels = rels
    .replace(
      /(<Relationship\b[^>]*?)\bTarget="[^"]*"([^>]*?Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"[^>]*\/>)/gi,
      `$1Target="${escapedUrl}"$2`
    )
    .replace(
      /(<Relationship\b[^>]*?Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"[^>]*?)\bTarget="[^"]*"([^>]*\/>)/gi,
      `$1Target="${escapedUrl}"$2`
    );

  slide = slide.replace(/(<a:hlinkClick[^>]*tooltip=")[^"]*(")/gi, `$1${escapedUrl}$2`);

  console.log('--- CHECK SLIDE XML HLINK ---');
  console.log(slide.match(/<a:hlinkClick[^>]*\/>/g));

  console.log('--- CHECK RELS XML HYPERLINK ---');
  console.log(rels.match(/<Relationship[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"[^>]*\/>/g));
}

test().catch(console.error);
