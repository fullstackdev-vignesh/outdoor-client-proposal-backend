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

// Finds (or adds) a custom numFmt entry for the given Excel format code — custom format ids
// must be >=164 per the OOXML spec (0-163 are reserved built-ins). Returns the id to use.
function ensureNumFmt(stylesXml, formatCode) {
  const existing = stylesXml.match(new RegExp(`<numFmt numFmtId="(\\d+)" formatCode="${formatCode.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`));
  if (existing) return { stylesXml, numFmtId: Number(existing[1]) };

  const usedIds = [...stylesXml.matchAll(/<numFmt numFmtId="(\d+)"/g)].map((m) => Number(m[1]));
  const numFmtId = Math.max(163, ...usedIds) + 1;
  const entry = `<numFmt numFmtId="${numFmtId}" formatCode="${formatCode}"/>`;

  if (stylesXml.includes('<numFmts')) {
    stylesXml = stylesXml
      .replace(/<numFmts count="(\d+)">/, (m, count) => `<numFmts count="${Number(count) + 1}">`)
      .replace('</numFmts>', `${entry}</numFmts>`);
  } else {
    stylesXml = stylesXml.replace(/(<styleSheet[^>]*>)/, `$1<numFmts count="1">${entry}</numFmts>`);
  }
  return { stylesXml, numFmtId };
}

// Clones cellXfs[oldIndex] (same font/fill/border/alignment) with its numFmtId swapped to
// `numFmtId` and appends it as a new style — the original style (and every other cell using
// it) is left completely untouched.
function cloneStyleWithNumFmt(stylesXml, oldIndex, numFmtId) {
  const m = stylesXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  const count = Number(m[1]);
  const xfs = [...m[2].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((x) => x[0]);
  const old = xfs[oldIndex];
  if (!old) return { stylesXml, newIndex: oldIndex };

  let newXf = old.includes('numFmtId="')
    ? old.replace(/numFmtId="\d+"/, `numFmtId="${numFmtId}"`)
    : old.replace('<xf ', `<xf numFmtId="${numFmtId}" `);
  if (!newXf.includes('applyNumberFormat=')) {
    newXf = newXf.replace('<xf ', '<xf applyNumberFormat="1" ');
  } else {
    newXf = newXf.replace(/applyNumberFormat="0"/, 'applyNumberFormat="1"');
  }

  const newIndex = count;
  stylesXml = stylesXml
    .replace(/<cellXfs count="(\d+)">/, `<cellXfs count="${count + 1}">`)
    .replace('</cellXfs>', `${newXf}</cellXfs>`);
  return { stylesXml, newIndex };
}

// Reformats every cell currently using one of `styleIds` to a cloned version of that same
// style with Indian comma-grouped digits (e.g. 1000000 -> displayed as 10,00,000) — matched
// by STYLE ID rather than column letter, so it still finds the right cells even after
// removeColumns/insertColumnBefore have shifted which letter a cost column actually lives at.
function applyIndianNumberFormat(sheetXml, stylesXml, styleIds) {
  const { stylesXml: withFmt, numFmtId } = ensureNumFmt(stylesXml, '#,##,##0');
  stylesXml = withFmt;
  const styleMap = {};
  for (const oldId of styleIds) {
    if (!sheetXml.includes(`s="${oldId}"`)) continue;
    const { stylesXml: cloned, newIndex } = cloneStyleWithNumFmt(stylesXml, oldId, numFmtId);
    stylesXml = cloned;
    styleMap[oldId] = newIndex;
  }
  for (const [oldId, newId] of Object.entries(styleMap)) {
    sheetXml = sheetXml.split(`s="${oldId}"`).join(`s="${newId}"`);
  }
  return { sheetXml, stylesXml };
}

function colToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}
function numToCol(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Removes whole columns (by letter) from EVERY row in the sheet, shifting everything to
// their right left by one to close the gap — also fixes up <cols> width defs, <mergeCells>,
// and the sheet's <dimension>. Every analyzed template's formulas only reference cells in
// their OWN row (never a fixed column further down/up), so only the cell's own address needs
// to move — formula text never needs rewriting here.
function removeColumns(sheetXml, lettersToRemove) {
  for (const removeNum of lettersToRemove.map(colToNum).sort((a, b) => b - a)) {
    // Note: the attrs group deliberately excludes '/' too, not just '>' — otherwise it can
    // swallow a self-closing tag's own "/" before the alternation below gets to see it,
    // which forces the "/>" branch to fail and the ">...</c>" branch to opportunistically
    // match all the way to some LATER cell's "</c>" instead (silently merging rows).
    sheetXml = sheetXml.replace(/<c r="([A-Z]+)(\d+)"([^/>]*)(\/>|>[\s\S]*?<\/c>)/g, (whole, col, row, attrs, tail) => {
      const n = colToNum(col);
      if (n === removeNum) return '';
      if (n < removeNum) return whole;
      return `<c r="${numToCol(n - 1)}${row}"${attrs}${tail}`;
    });
    sheetXml = sheetXml.replace(/<col ([^>]*)\/>/g, (whole, attrs) => {
      const minM = attrs.match(/min="(\d+)"/);
      const maxM = attrs.match(/max="(\d+)"/);
      if (!minM || !maxM) return whole;
      let min = Number(minM[1]);
      let max = Number(maxM[1]);
      if (min === removeNum && max === removeNum) return '';
      if (min >= removeNum) min -= 1;
      if (max >= removeNum) max -= 1;
      return `<col ${attrs.replace(/min="\d+"/, `min="${min}"`).replace(/max="\d+"/, `max="${max}"`)}/>`;
    });
    sheetXml = sheetXml.replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g, (whole, c1, r1, c2, r2) => {
      const n1 = colToNum(c1);
      const n2 = colToNum(c2);
      if (n1 === removeNum && n2 === removeNum) return '';
      const newN1 = n1 > removeNum ? n1 - 1 : n1;
      const newN2 = n2 > removeNum ? n2 - 1 : n2;
      return `<mergeCell ref="${numToCol(newN1)}${r1}:${numToCol(newN2)}${r2}"/>`;
    });
    sheetXml = sheetXml.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/, (whole, c1, r1, c2, r2) => {
      const n2 = colToNum(c2);
      return `<dimension ref="${c1}${r1}:${numToCol(n2 > removeNum ? n2 - 1 : n2)}${r2}"/>`;
    });
  }
  return sheetXml;
}

// Reverse of removeColumns for a single column: shifts every cell/col-width/mergeCell/
// dimension at or after `beforeLetter` one column to the RIGHT, opening up an empty column
// at `beforeLetter` — it does not create any cell content there itself (see insertCellInRow).
// Excel's absolute column limit (XFD) — a sheet's own trailing "rest of the sheet" <col> range
// (e.g. min="19" max="16384", covering every otherwise-undefined column) already sits AT this
// ceiling in most templates, so incrementing its max on every insertColumnBefore call (once per
// inserted fee column) pushes it past 16384 — an out-of-range value Excel rejects as corrupt
// "column information" and silently repairs on open, which can drop unrelated formatting (e.g.
// row borders) as a side effect. Clamping keeps every <col> range within the sheet's real bounds.
const MAX_EXCEL_COLUMN = 16384;

function insertColumnBefore(sheetXml, beforeLetter) {
  const insertNum = colToNum(beforeLetter);
  sheetXml = sheetXml.replace(/<c r="([A-Z]+)(\d+)"/g, (whole, col, row) => {
    const n = colToNum(col);
    return n >= insertNum ? `<c r="${numToCol(n + 1)}${row}"` : whole;
  });
  sheetXml = sheetXml.replace(/<col ([^>]*)\/>/g, (whole, attrs) => {
    const minM = attrs.match(/min="(\d+)"/);
    const maxM = attrs.match(/max="(\d+)"/);
    if (!minM || !maxM) return whole;
    let min = Number(minM[1]);
    let max = Number(maxM[1]);
    if (min >= insertNum) min = Math.min(min + 1, MAX_EXCEL_COLUMN);
    if (max >= insertNum) max = Math.min(max + 1, MAX_EXCEL_COLUMN);
    return `<col ${attrs.replace(/min="\d+"/, `min="${min}"`).replace(/max="\d+"/, `max="${max}"`)}/>`;
  });
  sheetXml = sheetXml.replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g, (whole, c1, r1, c2, r2) => {
    const n1 = colToNum(c1);
    const n2 = colToNum(c2);
    const newN1 = n1 >= insertNum ? n1 + 1 : n1;
    const newN2 = n2 >= insertNum ? n2 + 1 : n2;
    return `<mergeCell ref="${numToCol(newN1)}${r1}:${numToCol(newN2)}${r2}"/>`;
  });
  sheetXml = sheetXml.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/, (whole, c1, r1, c2, r2) => {
    const n2 = colToNum(c2);
    return `<dimension ref="${c1}${r1}:${numToCol(n2 >= insertNum ? n2 + 1 : n2)}${r2}"/>`;
  });
  return sheetXml;
}

// Splices a brand-new <c> element into a row at the correct sorted position (Excel expects
// cells within a row in ascending column order) — used after insertColumnBefore opens up a
// gap, since that only shifts EXISTING cells and never creates new ones.
function insertCellInRow(sheetXml, rowNum, cellRef, cellXml) {
  const rowRe = new RegExp(`(<row r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
  const m = sheetXml.match(rowRe);
  if (!m) return sheetXml;
  const [whole, openTag, inner, closeTag] = m;
  const targetNum = colToNum(cellRef.match(/[A-Z]+/)[0]);
  const cellMatches = [...inner.matchAll(/<c r="([A-Z]+)\d+"[^/>]*(?:\/>|>[\s\S]*?<\/c>)/g)];
  let insertPos = inner.length;
  for (const cm of cellMatches) {
    if (colToNum(cm[1]) > targetNum) {
      insertPos = cm.index;
      break;
    }
  }
  const newInner = inner.slice(0, insertPos) + cellXml + inner.slice(insertPos);
  return sheetXml.slice(0, m.index) + openTag + newInner + closeTag + sheetXml.slice(m.index + whole.length);
}

// Finds whichever existing <col min,max> range covers `colNum` and returns its own column-level
// `style="…"` (or undefined if that range has none). A column-level style applies to EVERY row
// in that column, including ones with no cell entry at all — which is why every pre-existing
// column in these templates still shows its default border/fill for blank rows far below the
// generated table, while a brand-new column with no such style looks visibly different there.
function getColStyle(sheetXml, colNum) {
  for (const m of sheetXml.matchAll(/<col ([^>]*)\/>/g)) {
    const minM = m[1].match(/min="(\d+)"/);
    const maxM = m[1].match(/max="(\d+)"/);
    if (!minM || !maxM) continue;
    if (colNum < Number(minM[1]) || colNum > Number(maxM[1])) continue;
    const styleM = m[1].match(/\bstyle="(\d+)"/);
    return styleM ? styleM[1] : undefined;
  }
  return undefined;
}

// Widens an existing <col> width entry for one column letter (or adds one if the sheet never
// explicitly defined that column's width) — cosmetic only, doesn't move any cells. `styleId`
// (only meaningful on the "brand new column" path — an existing column keeps whatever
// column-level style it already had) carries over a neighboring column's own style so blank
// rows far below the generated table render identically instead of the new column looking
// like a stray, unstyled gap.
function setColumnWidth(sheetXml, colLetter, width, styleId) {
  const num = colToNum(colLetter);
  let found = false;
  sheetXml = sheetXml.replace(/<col ([^>]*)\/>/g, (whole, attrs) => {
    const minM = attrs.match(/min="(\d+)"/);
    const maxM = attrs.match(/max="(\d+)"/);
    if (!minM || !maxM) return whole;
    const min = Number(minM[1]);
    const max = Number(maxM[1]);
    if (num < min || num > max) return whole;
    found = true;
    if (min === max) {
      let newAttrs = attrs.replace(/\scustomWidth="[^"]*"/, '');
      newAttrs = newAttrs.includes('width="') ? newAttrs.replace(/width="[^"]*"/, `width="${width}"`) : `${newAttrs} width="${width}"`;
      return `<col ${newAttrs} customWidth="1"/>`;
    }
    // This column shares a multi-column range with others — leave the range alone and add a
    // dedicated single-column override for just this one (inserted right after, so it wins).
    return `${whole}<col min="${num}" max="${num}" width="${width}" customWidth="1"/>`;
  });
  if (!found) {
    const styleAttr = styleId ? ` style="${styleId}"` : '';
    sheetXml = sheetXml.includes('<cols>')
      ? sheetXml.replace('</cols>', `<col min="${num}" max="${num}" width="${width}" customWidth="1"${styleAttr}/></cols>`)
      : sheetXml.replace(/(<sheetData)/, `<cols><col min="${num}" max="${num}" width="${width}" customWidth="1"${styleAttr}/></cols>$1`);
  }
  return sheetXml;
}

function addMergeCell(sheetXml, ref) {
  if (!sheetXml.includes('<mergeCells')) return sheetXml;
  return sheetXml
    .replace(/<mergeCells count="(\d+)">/, (m, count) => `<mergeCells count="${Number(count) + 1}">`)
    .replace('</mergeCells>', `<mergeCell ref="${ref}"/></mergeCells>`);
}

// Adinn (and, via config, ROTN): Vendor Name/Vendor Cost are unconditionally dropped from the
// sheet, and an Agency Comm / GST column is inserted immediately before the Total Cost anchor —
// but ONLY when the client has that percentage set — in that order, so Total Cost ends up right
// after whichever of them are present (matching the Jagran format's own Agency Comm -> GST ->
// Total column order and calculation: each fee = (running subtotal so far) * percent / 100, and
// Total Cost becomes the sum of the base cost columns plus whichever fees got inserted).
//
// `cfg.feeRangeStartCol` (default 'J', Adinn's Display Cost column) and `cfg.feeRangeBaseEndCol`
// (default 'L', Adinn's Mounting Cost column) describe the fixed base cost range each fee's
// running percentage is computed over before any fees are inserted. `cfg.feeStyleIds` (default
// Adinn's own {header:25, data:14, total:22}) lets a different uploaded file's own style ids be
// reused for the newly inserted cells so they visually match that file's existing columns.
function applyAdinnDynamicColumns(sheetXml, cfg, client, firstDataRow, lastUsedRow, totalRow) {
  if (cfg.removeColumns?.length) sheetXml = removeColumns(sheetXml, cfg.removeColumns);
  if (!cfg.feeColumnsBeforeAnchor) return sheetXml;

  const fees = [];
  if (client?.agencyComm) fees.push({ key: 'agencyComm', label: 'Agency Comm', percent: Number(client.agencyComm) });
  if (client?.gst) fees.push({ key: 'gst', label: 'GST', percent: Number(client.gst) });
  if (!fees.length) return sheetXml;

  const rangeStartCol = cfg.feeRangeStartCol || 'J';
  const baseEndCol = cfg.feeRangeBaseEndCol || 'L';
  const styleIds = { header: 25, data: 14, total: 22, ...(cfg.feeStyleIds || {}) };

  let totalCol = cfg.feeColumnsBeforeAnchor;
  const inserted = [];
  for (const fee of fees) {
    sheetXml = insertColumnBefore(sheetXml, totalCol);
    inserted.push({ ...fee, col: totalCol });
    totalCol = numToCol(colToNum(totalCol) + 1);
  }

  // Header label for each inserted fee column, on whichever row this format's OWN header text
  // actually lives on (default row 1, matching Adinn's real row1:row2 merged header) — and its
  // merged companion row IF this format actually merges one (Adinn does; ROTN's real header is
  // row1 ALONE with no merge anywhere in the sheet, confirmed against the real file, so forcing
  // a row-2 companion+merge there put a stray blue header-style fill one row above ROTN's real
  // header, on top of what's otherwise a genuinely blank spacer row).
  const headerRow = cfg.fixedExtraColumnHeaderRow || 1;
  const blankHeaderRow = 'fixedExtraColumnBlankHeaderRow' in cfg ? cfg.fixedExtraColumnBlankHeaderRow : 2;
  for (const fee of inserted) {
    const headerText = `${fee.label} ${fee.percent}%`;
    sheetXml = insertCellInRow(
      sheetXml,
      headerRow,
      `${fee.col}${headerRow}`,
      `<c r="${fee.col}${headerRow}" s="${styleIds.header}" t="inlineStr"><is><t>${xmlEscape(headerText)}</t></is></c>`
    );
    if (blankHeaderRow) {
      sheetXml = insertCellInRow(sheetXml, blankHeaderRow, `${fee.col}${blankHeaderRow}`, `<c r="${fee.col}${blankHeaderRow}" s="${styleIds.header}"/>`);
      const [top, bottom] = [headerRow, blankHeaderRow].sort((a, b) => a - b);
      sheetXml = addMergeCell(sheetXml, `${fee.col}${top}:${fee.col}${bottom}`);
    }
  }

  // Each fee is a running percentage of the subtotal built up so far (the base cost columns,
  // then each previously-inserted fee) — same compounding rule Jagran already uses for its own
  // Agency Comm -> GST columns. Total Cost is rewritten to sum everything up to (but not
  // including) itself.
  for (let r = firstDataRow; r <= lastUsedRow; r++) {
    let rangeEnd = baseEndCol;
    for (const fee of inserted) {
      sheetXml = insertCellInRow(sheetXml, r, `${fee.col}${r}`, `<c r="${fee.col}${r}" s="${styleIds.data}"/>`);
      sheetXml = setCellFormula(sheetXml, `${fee.col}${r}`, `ROUND(SUM(${rangeStartCol}${r}:${rangeEnd}${r})*${fee.percent}/100,2)`);
      rangeEnd = fee.col;
    }
    sheetXml = setCellFormula(sheetXml, `${totalCol}${r}`, `SUM(${rangeStartCol}${r}:${rangeEnd}${r})`);
  }

  // Total row: each inserted fee column gets its own SUM-down-the-column cell (didn't exist
  // before the column was inserted). The pre-existing Total Cost total-row cell moved to
  // `totalCol` via insertColumnBefore's generic shift, but its FORMULA TEXT still literally
  // says the old anchor's stale range (only cell addresses get renamed by the shift, never
  // formula text) — since that old anchor letter now holds a fee column's value instead, that
  // stale text must be rewritten to sum whatever column Total Cost actually ended up in.
  for (const fee of inserted) {
    sheetXml = insertCellInRow(sheetXml, totalRow, `${fee.col}${totalRow}`, `<c r="${fee.col}${totalRow}" s="${styleIds.total}"/>`);
    sheetXml = setCellFormula(sheetXml, `${fee.col}${totalRow}`, `SUM(${fee.col}${firstDataRow}:${fee.col}${lastUsedRow})`);
  }
  sheetXml = setCellFormula(sheetXml, `${totalCol}${totalRow}`, `SUM(${totalCol}${firstDataRow}:${totalCol}${lastUsedRow})`);

  // Jagran/ROTN both have genuine pre-existing bordered-blank spacer row(s) around the data —
  // one or more rows between the last data row and Total (`totalRow - lastUsedRow - 1` of
  // them), and for ROTN also one row ABOVE firstDataRow (its own SUM range starts at
  // cfg.sumRangeStartRow, one row earlier than the real data). Every other column already has
  // a cell there; these fee columns didn't get one, leaving a gap in the border/fill exactly
  // at that row — same class of bug already fixed for applyFixedExtraColumns's new columns.
  const spacerRows = [];
  for (let r = lastUsedRow + 1; r < totalRow; r++) spacerRows.push(r);
  if (cfg.sumRangeStartRow && cfg.sumRangeStartRow < firstDataRow) spacerRows.push(cfg.sumRangeStartRow);
  for (const fee of inserted) {
    const leftCol = numToCol(colToNum(fee.col) - 1);
    for (const r of spacerRows) {
      // Row 2 sometimes already got a cell here (the header-merge companion row above always
      // inserts one at row 2, which for ROTN happens to be the SAME row as its pre-data
      // spacer) — skip rather than insert a second, invalid duplicate `<c>` for the same ref.
      if (sheetXml.includes(`<c r="${fee.col}${r}"`)) continue;
      const s = getCellStyle(sheetXml, `${leftCol}${r}`);
      sheetXml = insertCellInRow(sheetXml, r, `${fee.col}${r}`, `<c r="${fee.col}${r}"${s ? ` s="${s}"` : ''}/>`);
    }
  }

  return sheetXml;
}

// Returns the `s="…"` style id of an existing cell (e.g. "B5"), or undefined if that exact
// cell has no style attribute (or doesn't exist). Used so a brand-new inserted column's cells
// can copy their immediate left neighbor's REAL style byte-for-byte instead of guessing a
// plausible-looking existing style id — a guessed id can still differ in some way (border
// weight, one missing side, font) that only shows up as a subtly "off" gridline once opened in
// Excel, which is exactly what happened the first time around.
function getCellStyle(sheetXml, cellRef) {
  const m = sheetXml.match(new RegExp(`<c r="${cellRef}"([^>]*?)(?:/>|>)`));
  if (!m) return undefined;
  const styleMatch = m[1].match(/\bs="(\d+)"/);
  return styleMatch ? styleMatch[1] : undefined;
}

// Brand-new columns that aren't part of the uploaded master file at all (e.g. Adinn's Media
// Type/Latitude/Longitude, Jagran's Media Type/Latitude/Longitude) — unlike Agency Comm/GST
// (conditional on the client, added by applyAdinnDynamicColumns/applyConditionalFeeColumns
// further below), these are unconditional for every proposal in that format, so they're
// inserted once, right up front, before any row gets its values filled in.
//
// Each entry's `before` letter is the column to open a gap at, evaluated against the sheet's
// CURRENT state at the moment that entry runs — i.e. it already accounts for every earlier
// entry in the list having shifted things first. `cfg.columns`/`totalColumns`/`removeColumns`/
// `feeRangeStartCol`/`feeRangeBaseEndCol`/`feeColumnsBeforeAnchor`/`conditionalFee*` for this
// format are written using the FINAL letters (after all of these insertions), since they
// always happen.
//
// Different uploaded formats put their actual header TEXT on a different row — Adinn's row1 is
// the real (merged row1:row2) header; Jagran's row1 is an entirely empty spacer row and row2
// alone holds the header text for every existing column, with no merge at all. `cfg.fixedExtraColumnHeaderRow`
// (default 1) / `cfg.fixedExtraColumnBlankHeaderRow` (default 2, or falsy to skip it and not
// merge at all) let each format match its own real structure instead of assuming Adinn's.
//
// Every inserted cell's style is copied fresh from whichever real column ends up immediately
// to its LEFT (per row) — never a fixed/guessed style id — so it's guaranteed to look exactly
// like its neighbor no matter which format or which column this happens to land next to.
//
// `col.decimalValue: true` (Latitude/Longitude) additionally clones that copied DATA-row style
// (once, not per-row — the same style id repeats for every data row already) down to a General
// number format before reuse. Without this, a coordinate like 9.3147566 silently displays
// rounded to "9" whenever its left neighbor happens to be a currency/whole-number column
// (Total Cost, Total (Incl. All), etc.) — copying the neighbor's border/fill is exactly right,
// but copying its number format along with it is not. The header/blank-header/total-row cells
// never hold a decimal VALUE (text label or genuinely blank), so only the data-row style needs
// this — cloning there alone is enough.
function applyFixedExtraColumns(sheetXml, stylesXml, cfg, firstDataRow, lastUsedRow, totalRow) {
  const extraColumns = cfg.fixedExtraColumns;
  const headerRow = cfg.fixedExtraColumnHeaderRow || 1;
  const blankHeaderRow = 'fixedExtraColumnBlankHeaderRow' in cfg ? cfg.fixedExtraColumnBlankHeaderRow : 2;
  // Some formats (Jagran) have `spacerRowsBeforeTotal` blank row(s) between the last real data
  // row and the Total row — a genuine pre-existing template row (bordered blank cells for
  // every ORIGINAL column) sitting right at cfg.lastDataRow+1..+N in the template's own
  // (still untrimmed) numbering, since this whole function runs before
  // removeUnusedRowsAndShiftTail renumbers anything. Without a matching cell here too, these
  // brand-new columns show a gap in the border exactly at that row, between an otherwise
  // fully-bordered data area and Total row.
  const spacerRows = [];
  for (let i = 1; i <= (cfg.spacerRowsBeforeTotal || 0); i++) spacerRows.push(cfg.lastDataRow + i);
  // ROTN also has a blank row BEFORE the first data row (its own SUM range starts at
  // `sumRangeStartRow`, row2, one row above `firstDataRow`, row3) — same "genuine pre-existing
  // bordered blank template row" situation as the after-data spacer, just on the other side.
  if (cfg.sumRangeStartRow && cfg.sumRangeStartRow < firstDataRow) {
    for (let r = cfg.sumRangeStartRow; r < firstDataRow; r++) spacerRows.push(r);
  }
  for (const col of extraColumns) {
    const leftCol = numToCol(colToNum(col.before) - 1);
    const styleAttr = (row) => {
      const s = getCellStyle(sheetXml, `${leftCol}${row}`);
      return s ? ` s="${s}"` : '';
    };
    const headerStyleAttr = styleAttr(headerRow);
    const blankHeaderStyleAttr = blankHeaderRow ? styleAttr(blankHeaderRow) : '';
    const totalStyleAttr = totalRow ? styleAttr(totalRow) : '';

    sheetXml = insertColumnBefore(sheetXml, col.before);
    sheetXml = insertCellInRow(
      sheetXml,
      headerRow,
      `${col.before}${headerRow}`,
      `<c r="${col.before}${headerRow}"${headerStyleAttr} t="inlineStr"><is><t>${xmlEscape(col.headerText)}</t></is></c>`
    );
    if (blankHeaderRow) {
      sheetXml = insertCellInRow(sheetXml, blankHeaderRow, `${col.before}${blankHeaderRow}`, `<c r="${col.before}${blankHeaderRow}"${blankHeaderStyleAttr}/>`);
      const [top, bottom] = [headerRow, blankHeaderRow].sort((a, b) => a - b);
      sheetXml = addMergeCell(sheetXml, `${col.before}${top}:${col.before}${bottom}`);
    }

    // Bordered placeholder cell (previously had no `s=` at all, which rendered with no
    // gridlines in this template's gridlines-off view) for every generated site row.
    let dataStyleAttr = null; // resolved once per column, reused for every data row
    for (let r = firstDataRow; r <= lastUsedRow; r++) {
      if (dataStyleAttr === null) {
        if (col.decimalValue && stylesXml) {
          const neighborStyleId = getCellStyle(sheetXml, `${leftCol}${r}`);
          if (neighborStyleId !== undefined) {
            const cloned = cloneStyleWithNumFmt(stylesXml, Number(neighborStyleId), 0); // 0 = built-in General
            stylesXml = cloned.stylesXml;
            dataStyleAttr = ` s="${cloned.newIndex}"`;
          } else {
            dataStyleAttr = '';
          }
        } else {
          dataStyleAttr = styleAttr(r);
        }
      }
      sheetXml = insertCellInRow(sheetXml, r, `${col.before}${r}`, `<c r="${col.before}${r}"${dataStyleAttr}/>`);
    }
    // The Total row has no cell at all for these brand-new columns (they don't exist in the
    // original template), which otherwise leaves that one row's worth of grey fill/border
    // missing right under them — a plain styled blank (no value/formula needed here).
    if (totalRow) {
      sheetXml = insertCellInRow(sheetXml, totalRow, `${col.before}${totalRow}`, `<c r="${col.before}${totalRow}"${totalStyleAttr}/>`);
    }
    for (const r of spacerRows) {
      sheetXml = insertCellInRow(sheetXml, r, `${col.before}${r}`, `<c r="${col.before}${r}"${styleAttr(r)}/>`);
    }
  }
  return { sheetXml, stylesXml };
}

// Jagran-only: unlike Adinn/ROTN, the uploaded master file already has native Agency Comm/GST
// columns baked in with a FIXED percentage in both the header label and every row's formula
// (e.g. "Agency comm. @ 2%", always computed regardless of who the client is). This makes them
// conditional on the actual proposal's client, matching Adinn's rule: a fee column is removed
// entirely when the client doesn't have that percentage set, and rewritten (header label +
// every row's formula, using the CLIENT'S real percentage instead of the file's fixed one) when
// they do. `cfg.conditionalFeeColumns` lists each fee in file column order (e.g. Agency Comm
// before GST); `cfg.conditionalFeeBaseColumns` are the always-present cost columns the running
// SUM starts from; `cfg.conditionalFeeTotalCol` is the final Total column after them.
function applyConditionalFeeColumns(sheetXml, cfg, client, firstDataRow, lastUsedRow, totalRow) {
  const feeDefs = cfg.conditionalFeeColumns;
  if (!feeDefs?.length) return sheetXml;

  const present = feeDefs.filter((f) => client?.[f.key]);
  const absent = feeDefs.filter((f) => !present.includes(f));

  if (absent.length) sheetXml = removeColumns(sheetXml, absent.map((f) => f.col));

  // Recompute each surviving column's final letter after removal — the original file's
  // base-cost / fee / total columns are always consecutive, so removing an absent one shifts
  // every letter after it one to the left.
  const orderedOriginal = [...cfg.conditionalFeeBaseColumns, ...feeDefs.map((f) => f.col), cfg.conditionalFeeTotalCol];
  const survivors = orderedOriginal.filter((letter) => !absent.some((f) => f.col === letter));
  const startNum = colToNum(cfg.conditionalFeeBaseColumns[0]);
  const letterMap = {};
  survivors.forEach((orig, i) => {
    letterMap[orig] = numToCol(startNum + i);
  });

  const baseStartCol = letterMap[cfg.conditionalFeeBaseColumns[0]];
  const baseEndCol = letterMap[cfg.conditionalFeeBaseColumns[cfg.conditionalFeeBaseColumns.length - 1]];
  const newTotalCol = letterMap[cfg.conditionalFeeTotalCol];

  // Header label — same cell the file's own fixed-percentage text already occupies, now
  // rewritten with the client's real percentage.
  for (const fee of present) {
    const col = letterMap[fee.col];
    const pct = Number(client[fee.key]);
    sheetXml = setCell(sheetXml, `${col}2`, `${fee.label} @ ${pct}%`, { text: true });
  }

  // Every data row's own fee/Total formulas — present fees keep the file's own compounding
  // rule (running % of the subtotal so far), just with the client's real percentage and the
  // post-removal column letters; Total always sums the base columns through whichever fee
  // column (if any) ends up last.
  for (let r = firstDataRow; r <= lastUsedRow; r++) {
    let rangeEnd = baseEndCol;
    for (const fee of present) {
      const col = letterMap[fee.col];
      const pct = Number(client[fee.key]);
      sheetXml = setCellFormula(sheetXml, `${col}${r}`, `SUM(${baseStartCol}${r}:${rangeEnd}${r})*${pct}%`);
      rangeEnd = col;
    }
    sheetXml = setCellFormula(sheetXml, `${newTotalCol}${r}`, `SUM(${baseStartCol}${r}:${rangeEnd}${r})`);
  }

  // Total row (already moved/renumbered to `totalRow` by the generic pipeline before this runs)
  // — same stale-formula-text problem as Adinn's own dynamic columns: whichever fee columns
  // survive need a fresh SUM-down-the-column formula, and Total's own total-row cell (whatever
  // letter it ended up at) needs its range rewritten to match.
  for (const fee of present) {
    const col = letterMap[fee.col];
    sheetXml = setCellFormula(sheetXml, `${col}${totalRow}`, `SUM(${col}${firstDataRow}:${col}${lastUsedRow})`);
  }
  sheetXml = setCellFormula(sheetXml, `${newTotalCol}${totalRow}`, `SUM(${newTotalCol}${firstDataRow}:${newTotalCol}${lastUsedRow})`);

  return sheetXml;
}

// Returns the raw XML of one <row>, or null if it isn't present.
function extractRow(sheetXml, rowNum) {
  const m = sheetXml.match(new RegExp(`<row r="${rowNum}"[^/>]*(?:/>|>[\\s\\S]*?</row>)`));
  return m ? m[0] : null;
}

function removeRow(sheetXml, rowNum) {
  return sheetXml.replace(new RegExp(`<row r="${rowNum}"[^/>]*(?:/>|>[\\s\\S]*?</row>)`), '');
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
    const rowRe = new RegExp(`<row r="${r}"[^/>]*(?:/>|>[\\s\\S]*?</row>)`);
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
    // W/H pair). When collapseSecondaryRow will delete that row outright (see
    // collapseSecondaryRows below), zeroing it first is pointless — skip it.
    if (!cfg.collapseSecondaryRow) {
      for (let j = 1; j < cfg.blockSize; j++) {
        for (const col of cfg.secondaryRowColumns || []) {
          sheetXml = setCell(sheetXml, `${col}${r + j}`, 0);
        }
      }
    }
  });
  return sheetXml;
}

// ROTN-only: rather than leave the block's 2nd row visible with a meaningless "0  0" (no
// Site field maps to that secondary W/H pair), delete it outright and shift everything below
// up by one row per site — turning the visible sheet into a clean one-row-per-site table.
// The Area formula on the surviving row (originally `G3*H3+G4*H4`, folding in the deleted
// row's W/H) is rewritten to just `G3*H3` so it stays numerically correct once that term is
// gone. Processed from the LAST site backward so row numbers for sites not yet processed stay
// valid while earlier ones are being deleted/shifted.
function collapseSecondaryRows(sheetXml, cfg, usableCount) {
  // Pass 1: delete each block's 2nd row and shift everything below up, working from the LAST
  // site backward so a not-yet-processed site's row number stays valid until its own turn.
  for (let i = usableCount - 1; i >= 0; i--) {
    const r = cfg.firstDataRow + i * cfg.blockSize;
    sheetXml = removeUnusedRowsAndShiftTail(sheetXml, r, r + 1);
  }
  // Pass 2: only NOW are all rows at their final, stable numbers (firstDataRow..
  // firstDataRow+usableCount-1) — rewrite every self-referencing per-row formula (Area, and
  // any other column the config lists) so each points at its OWN final row. Doing this DURING
  // pass 1 would go stale: shifting only renames cell addresses, never rewrites formula text,
  // so an earlier site's deletion shifting a LATER (already-rewritten) row down would leave
  // that row's formula still pointing at its old row number — this bit ROTN in practice: its
  // "Display Duration Cost" (N) column's formula (`M{r}/30*L{r}`) is an Excel "shared formula"
  // whose one HOST cell carries the literal text; if collapsing shifts that specific host row
  // without rewriting it, it keeps referencing its old (now wrong) row number.
  for (let i = 0; i < usableCount; i++) {
    const r = cfg.firstDataRow + i;
    for (const { column, build } of cfg.selfReferencingFormulas || []) {
      sheetXml = setCellFormula(sheetXml, `${column}${r}`, build(r, cfg.columns));
    }
  }
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
async function generateExcelFromTemplate(rows, { buffer, config, client } = {}) {
  const buf = buffer || fs.readFileSync(MASTER_XLSX_PATH);
  const cfg = config || getExcelConfig('generic');

  if (buf.length >= 8 && buf.readUInt32LE(0) === 0xe011cfd0) {
    throw new Error('This Excel template is a legacy .xls file, which cannot be used as a generation source. Please re-save it as .xlsx and re-upload it in Excel Templates.');
  }

  const zip = await JSZip.loadAsync(buf);
  const sheetPath = await resolveFirstSheetPath(zip);
  let sheetXml = await zip.file(sheetPath).async('string');
  // Loaded up front (not only when cfg.indianCommaStyleIds is set) because
  // applyFixedExtraColumns may also need to clone a style — e.g. Latitude/Longitude copy
  // their left neighbor's border/fill exactly, but that neighbor is very often a currency
  // column whose number format has 0 decimal places, which would otherwise silently round a
  // coordinate like 9.3147566 down to 9.
  const stylesPath = 'xl/styles.xml';
  let stylesXml = zip.file(stylesPath) ? await zip.file(stylesPath).async('string') : null;

  const blockSize = cfg.mode === 'block-per-site' ? cfg.blockSize || 1 : 1;
  const maxSites = Math.floor((cfg.lastDataRow - cfg.firstDataRow + 1) / blockSize);
  const usable = rows.slice(0, maxSites);

  if (cfg.fixedExtraColumns?.length && usable.length) {
    // `blockSize` is 1 for row-per-site formats (Adinn/Jagran), so this is just
    // firstDataRow + count - 1 there — but for block-per-site (ROTN, 2 rows per site), it must
    // span every block's OWN 2 rows too (the 2nd/secondary row gets a placeholder cell here
    // just like its real columns already do, even though collapseSecondaryRows deletes that
    // row later — same as how secondaryRowColumns are already handled).
    const lastRawRow = cfg.firstDataRow + usable.length * blockSize - 1;
    const result = applyFixedExtraColumns(sheetXml, stylesXml, cfg, cfg.firstDataRow, lastRawRow, cfg.totalRow);
    sheetXml = result.sheetXml;
    stylesXml = result.stylesXml;
  }

  // insertColumnBefore only ever moves CELL ADDRESSES, never rewrites a formula's own text —
  // fine for columns nothing references, but Adinn's Area column formula literally names the
  // (now-shifted) Qty/Width/Height cells, so inserting Media Type ahead of them leaves it
  // pointing at the wrong cells unless rewritten here (row-per-site only; ROTN's own block
  // mode already rewrites its equivalent formulas inside collapseSecondaryRows).
  if (cfg.mode !== 'block-per-site' && cfg.selfReferencingFormulas?.length && usable.length) {
    for (let r = cfg.firstDataRow; r <= cfg.firstDataRow + usable.length - 1; r++) {
      for (const { column, build } of cfg.selfReferencingFormulas) {
        sheetXml = setCellFormula(sheetXml, `${column}${r}`, build(r, cfg.columns));
      }
    }
  }

  sheetXml = cfg.mode === 'block-per-site' ? fillBlockPerSite(sheetXml, usable, cfg) : fillRowPerSite(sheetXml, usable, cfg);

  // fillRowPerSite/fillBlockPerSite skip any field whose value is '' (so leaving a field out of
  // buildExcelRow's output, or a site simply not having that data, doesn't blank out cells other
  // templates might rely on keeping their own static example text) — but a blank `rationale`
  // (site has no SiteInfo linked) must actually CLEAR the cell instead, or a jagran-excel-2 row
  // whose site was swapped in would keep showing whatever unrelated Rationale text the master
  // file's own original example proposal had for that row number.
  if (cfg.columns.rationale) {
    const effectiveBlockSize = cfg.mode === 'block-per-site' ? blockSize : 1;
    usable.forEach((row, i) => {
      const r = cfg.firstDataRow + i * effectiveBlockSize;
      sheetXml = setCell(sheetXml, `${cfg.columns.rationale}${r}`, row.rationale || '', { text: true });
    });
  }

  // Cosmetic header-cell overrides — e.g. the uploaded template's own header says "Media
  // Vehicle", shown as "Media Type" instead, without touching the original uploaded file.
  for (const { cell, text } of cfg.headerRenames || []) {
    sheetXml = setCell(sheetXml, cell, text, { text: true });
  }
  // Widen columns whose auto-fit width is too narrow to show real values (e.g. Area shows
  // "#####" once real numbers replace the blank template cells). For a brand-new column (no
  // existing <col> entry of its own — e.g. Media Type/Latitude/Longitude) this also carries
  // over its left neighbor's own column-level style, so it looks consistent below the table
  // instead of being the one column with no style rendering blank rows differently.
  for (const { col, width } of cfg.columnWidths || []) {
    const neighborColStyle = getColStyle(sheetXml, colToNum(col) - 1);
    sheetXml = setColumnWidth(sheetXml, col, width, neighborColStyle);
  }

  // Collapsing turns each used site's 2-row block into a single visible row (deleting the
  // block's 2nd row rather than leaving it blank/zeroed) — every row number at/after
  // firstDataRow shifts up by one per collapsed site, including the template's own
  // lastDataRow/totalRow constants, so those are adjusted locally for the rest of this run.
  let lastDataRow = cfg.lastDataRow;
  let totalRow = cfg.totalRow;
  let effectiveBlockSize = blockSize;
  if (cfg.mode === 'block-per-site' && cfg.collapseSecondaryRow && usable.length) {
    sheetXml = collapseSecondaryRows(sheetXml, cfg, usable.length);
    lastDataRow -= usable.length;
    totalRow -= usable.length;
    effectiveBlockSize = 1;
  }

  const lastUsedRow = usable.length ? cfg.firstDataRow + usable.length * effectiveBlockSize - 1 : cfg.firstDataRow - 1;

  // Detach the template's own total row before the generic blank-row cleanup runs, so it
  // moves instead of being deleted or mistaken for one of the unused rows below it.
  const templateTotalRowXml = usable.length && totalRow ? extractRow(sheetXml, totalRow) : null;
  if (templateTotalRowXml) sheetXml = removeRow(sheetXml, totalRow);

  sheetXml = removeUnusedRowsAndShiftTail(sheetXml, lastUsedRow, lastDataRow);

  if (templateTotalRowXml) {
    const spacer = cfg.spacerRowsBeforeTotal || 0;
    const newTotalRowNum = lastUsedRow + spacer + 1;
    let totalRowXml = renumberRow(templateTotalRowXml, totalRow, newTotalRowNum);
    totalRowXml = writeTotalRowCells(totalRowXml, cfg, lastUsedRow, newTotalRowNum);
    // Insert right after whatever row now sits immediately before the total row — the last
    // data row when there's no spacer, or the (already-shifted) spacer row when there is one
    // — never before it, or the total row would land out of order ahead of its own spacer.
    const anchorRow = lastUsedRow + spacer;
    sheetXml = sheetXml.replace(new RegExp(`(<row r="${anchorRow}"[^/>]*(?:/>|>[\\s\\S]*?</row>))`), `$1${totalRowXml}`);

    if (cfg.removeColumns?.length || cfg.feeColumnsBeforeAnchor) {
      sheetXml = applyAdinnDynamicColumns(sheetXml, cfg, client, cfg.firstDataRow, lastUsedRow, newTotalRowNum);
    }
    if (cfg.conditionalFeeColumns?.length) {
      sheetXml = applyConditionalFeeColumns(sheetXml, cfg, client, cfg.firstDataRow, lastUsedRow, newTotalRowNum);
    }
  } else if (cfg.removeColumns?.length) {
    sheetXml = removeColumns(sheetXml, cfg.removeColumns);
  }

  // Indian comma-grouped display (e.g. 1000000 -> "10,00,000") for whichever style ids the
  // config says hold cost/area values — matched by style id (not column letter) so it still
  // finds the right cells even after removeColumns/insertColumnBefore shifted them, and
  // automatically covers the dynamically-inserted Agency Comm/GST cells too since those reuse
  // Total Cost's own style ids.
  if (cfg.indianCommaStyleIds?.length && stylesXml) {
    const result = applyIndianNumberFormat(sheetXml, stylesXml, cfg.indianCommaStyleIds);
    sheetXml = result.sheetXml;
    stylesXml = result.stylesXml;
  }
  if (stylesXml) zip.file(stylesPath, stylesXml);

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

  // Every uploaded master file ships its own xl/calcChain.xml — Excel's cell-recalculation-order
  // cache, built for the ORIGINAL file's exact formula layout. None of the row/column
  // insert/delete/move operations above touch it, so after heavy edits (ROTN's block collapsing,
  // Adinn/ROTN/Jagran's fee column insert/remove, moving the Total row) it no longer matches the
  // sheet's actual formulas — Excel detects the mismatch on open and silently "repairs" the file,
  // which can drop formatting (e.g. row borders) as a side effect. It's dropped entirely here
  // (from the zip, [Content_Types].xml, and workbook.xml.rels) rather than attempting to keep it
  // in sync — it's purely a performance cache, not required for correctness, and `fullCalcOnLoad`
  // (set above) already makes Excel recompute every formula fresh the moment the file opens, so
  // losing the cache has no visible effect beyond a very slightly slower first calculation.
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    if (zip.file('[Content_Types].xml')) {
      let contentTypes = await zip.file('[Content_Types].xml').async('string');
      contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
      zip.file('[Content_Types].xml', contentTypes);
    }
    if (zip.file('xl/_rels/workbook.xml.rels')) {
      let workbookRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
      workbookRels = workbookRels.replace(/<Relationship[^>]*Type="[^"]*\/calcChain"[^>]*\/>/, '');
      zip.file('xl/_rels/workbook.xml.rels', workbookRels);
    }
  }

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { generateExcelFromTemplate };
