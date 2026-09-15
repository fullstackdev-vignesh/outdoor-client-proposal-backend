const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const MASTER_XLSX_PATH = path.join(__dirname, '..', '..', 'assets', 'proposal-templates', 'master.xlsx');

// The master template has pre-styled/pre-formulated empty rows 3-72 (row 73 is the closing
// border row) — we only ever write VALUES into existing cells in that range, never touch
// styles, merges, or the Area/Printing/Mounting/Total formulas already baked into the sheet.
const FIRST_DATA_ROW = 3;
const LAST_DATA_ROW = 72;

function xmlEscape(str) {
  return String(str ?? '').replace(/[<>&'"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
}

function setCell(xml, cellRef, value, { text } = {}) {
  const re = new RegExp(`<c r="${cellRef}"([^>]*)/>`);
  if (!re.test(xml)) return xml;
  if (text) {
    return xml.replace(re, `<c r="${cellRef}"$1 t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`);
  }
  return xml.replace(re, `<c r="${cellRef}"$1><v>${Number(value) || 0}</v></c>`);
}

/**
 * Populates the real Adinn Excel master template (master.xlsx) with proposal rows,
 * leaving all worksheet styling/formulas untouched. `rows` is an array of:
 * { siNo, city, media, location, qty, width, height, type, displayCostPerMonth, siteStatus, vendorName, vendorCost }
 */
async function generateExcelFromTemplate(rows) {
  const buf = fs.readFileSync(MASTER_XLSX_PATH);
  const zip = await JSZip.loadAsync(buf);
  let sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');

  const usable = rows.slice(0, LAST_DATA_ROW - FIRST_DATA_ROW + 1);
  usable.forEach((row, i) => {
    const r = FIRST_DATA_ROW + i;
    sheetXml = setCell(sheetXml, `A${r}`, row.siNo);
    sheetXml = setCell(sheetXml, `B${r}`, row.city, { text: true });
    sheetXml = setCell(sheetXml, `C${r}`, row.media, { text: true });
    sheetXml = setCell(sheetXml, `D${r}`, row.location, { text: true });
    sheetXml = setCell(sheetXml, `E${r}`, row.qty);
    sheetXml = setCell(sheetXml, `F${r}`, row.width);
    sheetXml = setCell(sheetXml, `G${r}`, row.height);
    sheetXml = setCell(sheetXml, `H${r}`, row.type, { text: true });
    sheetXml = setCell(sheetXml, `J${r}`, row.displayCostPerMonth);
    sheetXml = setCell(sheetXml, `N${r}`, row.siteStatus, { text: true });
    sheetXml = setCell(sheetXml, `O${r}`, row.vendorName, { text: true });
    sheetXml = setCell(sheetXml, `P${r}`, row.vendorCost);
  });

  let workbookXml = await zip.file('xl/workbook.xml').async('string');
  workbookXml = workbookXml.replace('<calcPr calcId="124519"/>', '<calcPr calcId="124519" fullCalcOnLoad="1"/>');

  zip.file('xl/worksheets/sheet1.xml', sheetXml);
  zip.file('xl/workbook.xml', workbookXml);

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { generateExcelFromTemplate, FIRST_DATA_ROW, LAST_DATA_ROW };
