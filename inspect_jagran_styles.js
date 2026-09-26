const JSZip = require('jszip');
const fs = require('fs');

async function inspect(f) {
  console.log('=====', f);
  const buf = fs.readFileSync(f);
  const zip = await JSZip.loadAsync(buf);
  const stylesXml = await zip.file('xl/styles.xml').async('string');
  const m = stylesXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  const xfs = [...m[2].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((x) => x[0]);
  [2, 3, 5, 12, 16, 17, 21, 22, 23, 24, 25, 26, 27, 28, 29].forEach((i) => console.log(i, xfs[i]));
}

(async () => {
  await inspect(String.raw`C:\Users\User\Documents\outdoortemplate\Jagran_format\excel_format_1\TVS - Plan Format2.xlsx`);
  await inspect(String.raw`C:\Users\User\Documents\outdoortemplate\Jagran_format\excel_format_2\Coimbatore Propoasl Jagran2.xlsx`);
})();
