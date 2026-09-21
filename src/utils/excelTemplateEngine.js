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
    if (min >= insertNum) min += 1;
    if (max >= insertNum) max += 1;
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

// Widens an existing <col> width entry for one column letter (or adds one if the sheet
// never explicitly defined that column's width) — cosmetic only, doesn't move any cells.
function setColumnWidth(sheetXml, colLetter, width) {
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
    sheetXml = sheetXml.includes('<cols>')
      ? sheetXml.replace('</cols>', `<col min="${num}" max="${num}" width="${width}" customWidth="1"/></cols>`)
      : sheetXml.replace(/(<sheetData)/, `<cols><col min="${num}" max="${num}" width="${width}" customWidth="1"/></cols>$1`);
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

  // Header (row 1 label + merged row 2 blank cell) for each inserted fee column.
  for (const fee of inserted) {
    const headerText = `${fee.label} ${fee.percent}%`;
    sheetXml = insertCellInRow(sheetXml, 1, `${fee.col}1`, `<c r="${fee.col}1" s="${styleIds.header}" t="inlineStr"><is><t>${xmlEscape(headerText)}</t></is></c>`);
    sheetXml = insertCellInRow(sheetXml, 2, `${fee.col}2`, `<c r="${fee.col}2" s="${styleIds.header}"/>`);
    sheetXml = addMergeCell(sheetXml, `${fee.col}1:${fee.col}2`);
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

  return sheetXml;
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

  const blockSize = cfg.mode === 'block-per-site' ? cfg.blockSize || 1 : 1;
  const maxSites = Math.floor((cfg.lastDataRow - cfg.firstDataRow + 1) / blockSize);
  const usable = rows.slice(0, maxSites);

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
  // "#####" once real numbers replace the blank template cells).
  for (const { col, width } of cfg.columnWidths || []) {
    sheetXml = setColumnWidth(sheetXml, col, width);
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
  if (cfg.indianCommaStyleIds?.length) {
    const stylesPath = 'xl/styles.xml';
    if (zip.file(stylesPath)) {
      let stylesXml = await zip.file(stylesPath).async('string');
      const result = applyIndianNumberFormat(sheetXml, stylesXml, cfg.indianCommaStyleIds);
      sheetXml = result.sheetXml;
      zip.file(stylesPath, result.stylesXml);
    }
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
