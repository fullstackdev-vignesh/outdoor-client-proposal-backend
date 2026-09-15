const JSZip = require('jszip');
const fs = require('fs');

async function run() {
  const buf = fs.readFileSync('assets/proposal-templates/template-1789471038153-fb1d78ae.pptx');
  const zip = await JSZip.loadAsync(buf);
  const presRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  let presentation = await zip.file('ppt/presentation.xml').async('string');

  console.log('Original presentation:', presentation);

  const baseName = 'slide4';
  const relMatch = presRels.match(new RegExp(`Id="(rId\\d+)"[^>]*Target="(?:/ppt/)?slides/${baseName}\\.xml"`)) ||
                   presRels.match(new RegExp(`Target="(?:/ppt/)?slides/${baseName}\\.xml"[^>]*Id="(rId\\d+)"`));

  console.log('relMatch for slide4:', relMatch ? relMatch[1] : 'null');
  if (relMatch) {
    const relId = relMatch[1];
    presentation = presentation.replace(new RegExp(`<p:sldId[^>]*r:id="${relId}"[^/>]*/>`), '');
    console.log('Updated presentation:', presentation);
  }
}

run().catch(console.error);
