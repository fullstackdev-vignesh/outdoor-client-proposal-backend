const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { getExcelConfig } = require('../config/excelTemplateConfigs');

const MASTER_XLSX_PATH = path.join(__dirname, '..', '..', 'assets', 'proposal-templates', 'master.xlsx');

function xmlEscape(str) {
  return String(str ?? '').replace(/[<>&'"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
}

// Overwrites one cell's value whether it started out empty (self-closing `<c .../>`) or
// already held a formula/value (`<c ...>...</c>`, as Printing/Mounting Cost do in some
// templates — they ship as area-derived estimates like `=I3*13`, or a real example proposal's
// leftover data). Either way only that cell's own value/type changes; its style attribute and
// every other cell/formula/merge in the sheet is untouched.
function setCell(xml, cellRef, value, { text } = {}) {
  const selfClosing = new RegExp(`<c r="${cellRef}"([^>]*)/>`);
  const withContent = new RegExp(`<c r="${cellRef}"([^>]*)>[\\s\\S]*?</c>`);

  const buildTag = (attrs) => {
    const cleanAttrs = attrs.replace(/\st="[^"]*"/, '');
    return text
      ? `<c r="${cellRef}"${cleanAttrs} t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`
      : `<c r="${cellRef}"${cleanAttrs}><v>${Number(value) || 0}</v></c>`;
  };

  if (selfClosing.test(xml)) return xml.replace(selfClosing, (m, attrs) => buildTag(attrs));
  if (withContent.test(xml)) return xml.replace(withContent, (m, attrs) => buildTag(attrs));
  return xml;
}

// Same in-place overwrite as setCell, but writes a SUM formula instead of a literal value —
// used for the Total row so the workbook keeps recalculating from the actual site rows
// instead of a value baked in at generation time.
function setCellFormula(xml, cellRef, formula) {
  const selfClosing = new RegExp(`<c r="${cellRef}"([^>]*)/>`);
  const withContent = new RegExp(`<c r="${cellRef}"([^>]*)>[\\s\\S]*?</c>`);
  const buildTag = (attrs) => `<c r="${cellRef}"${attrs.replace(/\st="[^"]*"/, '')}><f>${formula}</f><v>0</v></c>`;
  if (selfClosing.test(xml)) return xml.replace(selfClosing, (m, attrs) => buildTag(attrs));
  if (withContent.test(xml)) return xml.replace(withContent, (m, attrs) => buildTag(attrs));
  return xml;
}

// Returns the raw XML of one <row>, or null if it isn't present.
function extractRow(sheetXml, rowNum) {
  const m = sheetXml.match(new RegExp(`<row r="${rowNum}"[^>]*(?:/>|>[\\s\\S]*?</row>)`));
  return m ? m[0] : null;
}

function removeRow(sheetXml, rowNum) {
  return sheetXml.replace(new RegExp(`<row r="${rowNum}"[^>]*(?:/>|>[\\s\\S]*?</row>)`), '');
}

// Renumbers a standalone <row> element (its own r= attribute and every cell's r="COL#") from
// `oldNum` to `newNum` — used to move a template's total/spacer row to sit right after the
// last generated site row instead of always living at a fixed row number.
function renumberRow(rowXml, oldNum, newNum) {
  if (oldNum === newNum) return rowXml;
  // The row's own attribute is bare (`<row r="73" ...>`); every cell inside it is
  // column-prefixed (`r="A73"`) — both need renumbering, so they're handled separately.
  rowXml = rowXml.replace(new RegExp(`^<row r="${oldNum}"`), `<row r="${newNum}"`);
  return rowXml.replace(/r="([A-Z]+)(\d+)"/g, (m, col, num) => (Number(num) === oldNum ? `r="${col}${newNum}"` : m));
}

// Fills SUM formulas into the template's own (moved, renumbered) total row — only the
// columns the config lists get a formula over the actual generated site rows; everything
// else is left exactly as the template had it, and every style/merge on the row is whatever
// the template already had.
function writeTotalRowCells(rowXml, cfg, sumRangeEndRow, totalRow) {
  const rangeStart = cfg.sumRangeStartRow || cfg.firstDataRow;
  const range = (col) => `${col}${rangeStart}:${col}${sumRangeEndRow}`;
  if (cfg.totalLabelCell && cfg.totalLabelText) {
    rowXml = setCell(rowXml, `${cfg.totalLabelCell}${totalRow}`, cfg.totalLabelText, { text: true });
  }
  for (const col of cfg.totalColumns || []) {
    rowXml = setCellFormula(rowXml, `${col}${totalRow}`, `SUM(${range(col)})`);
  }
  return rowXml;
}

// Deletes the unused data rows after the last populated site row, then shifts every row
// below the data block up to close the gap. Rows firstDataRow..lastUsedRow are never
// touched/renumbered — only rows AFTER lastDataRow are affected — so none of the data rows'
// own (row-local) formulas need adjusting.
function removeUnusedRowsAndShiftTail(sheetXml, lastUsedRow, lastDataRow) {
  if (lastUsedRow >= lastDataRow) return sheetXml;
  const shift = lastDataRow - lastUsedRow;

  for (let r = lastUsedRow + 1; r <= lastDataRow; r++) {
    const rowRe = new RegExp(`<row r="${r}"[^>]*(?:/>|>[\\s\\S]*?</row>)`);
    sheetXml = sheetXml.replace(rowRe, '');
  }

  sheetXml = sheetXml.replace(/<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g, (whole, rowNum, rowAttrs, inner) => {
    const n = Number(rowNum);
    if (n <= lastDataRow) return whole;
    const newRowNum = n - shift;
    const newInner = inner.replace(/r="([A-Z]+)(\d+)"/g, (m, col, num) => (Number(num) === n ? `r="${col}${newRowNum}"` : m));
    return `<row r="${newRowNum}"${rowAttrs}>${newInner}</row>`;
  });
  sheetXml = sheetXml.replace(/<row r="(\d+)"([^>]*)\/>/g, (whole, rowNum, rowAttrs) => {
    const n = Number(rowNum);
    if (n <= lastDataRow) return whole;
    return `<row r="${n - shift}"${rowAttrs}/>`;
  });

  // Cosmetic: keep the sheet's declared dimension consistent with the new last row.
  sheetXml = sheetXml.replace(/(<dimension ref="[A-Z]+\d+:)([A-Z]+)(\d+)"/, (m, pre, col, lastRow) => `${pre}${col}${Number(lastRow) - shift}"`);

  return sheetXml;
}

async function resolveFirstSheetPath(zip) {
  const guess = 'xl/worksheets/sheet1.xml';
  if (zip.file(guess)) return guess;
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const firstSheet = wbXml.match(/<sheet[^>]*r:id="(rId\d+)"/);
  if (firstSheet) {
    const relMatch = relsXml.match(new RegExp(`Id="${firstSheet[1]}"[^>]*Target="([^"]+)"`));
    if (relMatch) return `xl/${relMatch[1]}`;
  }
  throw new Error('Could not locate a worksheet inside the Excel template');
}

function fillRowPerSite(sheetXml, usable, cfg) {
  usable.forEach((row, i) => {
    const r = cfg.firstDataRow + i;
    for (const [field, col] of Object.entries(cfg.columns)) {
      if (!(field in row) || row[field] === '' || row[field] == null) continue;
      const value = row[field];
      sheetXml = setCell(sheetXml, `${col}${r}`, value, { text: typeof value === 'string' });
    }
  });
  return sheetXml;
}

function fillBlockPerSite(sheetXml, usable, cfg) {
  usable.forEach((row, i) => {
    const r = cfg.firstDataRow + i * cfg.blockSize;
    for (const [field, col] of Object.entries(cfg.columns)) {
      if (!(field in row) || row[field] === '' || row[field] == null) continue;
      const value = row[field];
      sheetXml = setCell(sheetXml, `${col}${r}`, value, { text: typeof value === 'string' });
    }
    // The block's other row(s) carry fields with no matching Site data (e.g. a secondary
    // W/H pair) — zeroed so a real example proposal used as the master doesn't leak its own
    // old data into a fresh generation.
    for (let j = 1; j < cfg.blockSize; j++) {
      for (const col of cfg.secondaryRowColumns || []) {
        sheetXml = setCell(sheetXml, `${col}${r + j}`, 0);
      }
    }
  });
  return sheetXml;
}

/**
 * Populates a COPY of an Excel master template with proposal rows, leaving all worksheet
 * styling/formulas untouched, moves the template's OWN predefined final total row (already
 * styled/bordered in the uploaded file) to sit immediately after the last generated site row
 * with SUM formulas over just those rows, then trims the remaining unused trailing data rows.
 * `rows` is an array of generic field objects (see buildExcelRow in proposalFileGenerator.js);
 * which fields land in which column/row is entirely driven by `config`.
 *
 * `buffer` defaults to the bundled Adinn master.xlsx and `config` to its column map when no
 * specific template/format is supplied (today's existing default behaviour).
 */
async function generateExcelFromTemplate(rows, { buffer, config } = {}) {
  const buf = buffer || fs.readFileSync(MASTER_XLSX_PATH);
  const cfg = config || getExcelConfig('generic');

  if (buf.length >= 8 && buf.readUInt32LE(0) === 0xe011cfd0) {
    throw new Error('This Excel template is a legacy .xls file, which cannot be used as a generation source. Please re-save it as .xlsx and re-upload it in Excel Templates.');
  }

  const zip = await JSZip.loadAsync(buf);
  const sheetPath = await resolveFirstSheetPath(zip);
  let sheetXml = await zip.file(sheetPath).async('string');

  const blockSize = cfg.mode === 'block-per-site' ? cfg.blockSize || 1 : 1;
  const maxSites = Math.floor((cfg.lastDataRow - cfg.firstDataRow + 1) / blockSize);
  const usable = rows.slice(0, maxSites);

  sheetXml = cfg.mode === 'block-per-site' ? fillBlockPerSite(sheetXml, usable, cfg) : fillRowPerSite(sheetXml, usable, cfg);

  const lastUsedRow = usable.length ? cfg.firstDataRow + usable.length * blockSize - 1 : cfg.firstDataRow - 1;

  // Detach the template's own total row before the generic blank-row cleanup runs, so it
  // moves instead of being deleted or mistaken for one of the unused rows below it.
  const templateTotalRowXml = usable.length && cfg.totalRow ? extractRow(sheetXml, cfg.totalRow) : null;
  if (templateTotalRowXml) sheetXml = removeRow(sheetXml, cfg.totalRow);

  sheetXml = removeUnusedRowsAndShiftTail(sheetXml, lastUsedRow, cfg.lastDataRow);

  if (templateTotalRowXml) {
    const spacer = cfg.spacerRowsBeforeTotal || 0;
    const newTotalRowNum = lastUsedRow + spacer + 1;
    let totalRowXml = renumberRow(templateTotalRowXml, cfg.totalRow, newTotalRowNum);
    totalRowXml = writeTotalRowCells(totalRowXml, cfg, lastUsedRow, newTotalRowNum);
    // Insert right after whatever row now sits immediately before the total row — the last
    // data row when there's no spacer, or the (already-shifted) spacer row when there is one
    // — never before it, or the total row would land out of order ahead of its own spacer.
    const anchorRow = lastUsedRow + spacer;
    sheetXml = sheetXml.replace(new RegExp(`(<row r="${anchorRow}"[^>]*(?:/>|>[\\s\\S]*?</row>))`), `$1${totalRowXml}`);
  }

  if (zip.file('xl/workbook.xml')) {
    let workbookXml = await zip.file('xl/workbook.xml').async('string');
    if (workbookXml.includes('<calcPr')) {
      workbookXml = workbookXml.replace(/<calcPr([^/]*)\/>/, (m, attrs) =>
        attrs.includes('fullCalcOnLoad') ? m : `<calcPr${attrs} fullCalcOnLoad="1"/>`
      );
    } else {
      workbookXml = workbookXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    zip.file('xl/workbook.xml', workbookXml);
  }

  zip.file(sheetPath, sheetXml);

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { generateExcelFromTemplate };
