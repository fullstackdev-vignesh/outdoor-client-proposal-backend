const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function inspectSlide5() {
  const refDir = path.join(process.cwd(), 'reference', 'adinn-customized-format', 'adinn-customized-format.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(refDir));
  const xml = await zip.file('ppt/slides/slide5.xml').async('string');

  const groupRe = /<p:grpSp>(?:(?!<\/p:grpSp>)[\s\S])*?<\/p:grpSp>/g;
  let match;
  while ((match = groupRe.exec(xml))) {
    const block = match[0];
    const nameMatch = block.match(/name="([^"]+)"/);
    const name = nameMatch ? nameMatch[1] : 'unknown';

    // Extract offX, offY, extCx, extCy
    const offX = block.match(/<a:off x="([0-9-]+)" y="([0-9-]+)"/);
    const extC = block.match(/<a:ext cx="([0-9-]+)" cy="([0-9-]+)"/);

    console.log(`Group: ${name}`);
    if (offX) console.log(`  Offset: x=${offX[1]}, y=${offX[2]}`);
    if (extC) console.log(`  Extent: cx=${extC[1]}, cy=${extC[2]}`);

    const spNames = [...block.matchAll(/name="([^"]+)"/g)].map(m => m[1]);
    console.log(`  Shapes: ${spNames.join(', ')}`);
  }
}

inspectSlide5();
