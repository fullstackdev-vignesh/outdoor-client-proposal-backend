const JSZip = require('jszip');
const fs = require('fs');

async function inspect(f, totalRow) {
  console.log('=====', f, 'totalRow', totalRow);
  const buf = fs.readFileSync(f);
  const zip = await JSZip.loadAsync(buf);
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const m = sheetXml.match(new RegExp(`<row r="${totalRow}"[^>]*>[\\s\\S]*?</row>`));
  console.log(m ? m[0] : 'NOT FOUND');
}

(async () => {
  await inspect(String.raw`C:\Users\User\Documents\outdoortemplate\Jagran_format\excel_format_1\TVS - Plan Format2.xlsx`, 89);
  await inspect(String.raw`C:\Users\User\Documents\outdoortemplate\Jagran_format\excel_format_2\Coimbatore Propoasl Jagran2.xlsx`, 11);
})();
