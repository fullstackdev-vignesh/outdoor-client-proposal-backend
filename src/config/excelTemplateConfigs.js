// Per-format Excel layout maps — one entry per real uploaded template FORMAT, built by
// inspecting each vendor's actual .xlsx structure once (header row, data-row range, existing
// total-row position/formulas, block size). Generation (excelTemplateEngine.js) reads these
// instead of guessing column letters or assuming every template looks like Adinn's.
//
// mode: 'row-per-site'   — one data row per site, fixed column letters.
// mode: 'block-per-site' — a fixed-size group of rows per site (e.g. a 2nd row carries a
//                          secondary width/height pair with no matching Site field — left blank).
//
// `totalRow` / `sumRangeStartRow` describe the template's OWN predefined final total row
// (already styled/bordered in the uploaded file) — the engine moves that exact row to sit
// right after the last generated site row instead of building a new one from scratch.

const ADINN_EXCEL_1 = {
  mode: 'row-per-site',
  firstDataRow: 3,
  lastDataRow: 72,
  totalRow: 73,
  sumRangeStartRow: null, // null => same as firstDataRow
  totalLabelCell: 'A',
  totalLabelText: 'Total',
  hasTermsAfterTotal: true,
  columns: {
    siNo: 'A', city: 'B', media: 'C', location: 'D', qty: 'E', width: 'F', height: 'G',
    type: 'H', displayCostPerMonth: 'J', printingCost: 'K', mountingCost: 'L', siteStatus: 'N',
  },
  totalColumns: ['E', 'I', 'J', 'K', 'L', 'M'],
  // Vendor Name/Vendor Cost are unused — dropped from the sheet entirely (not just left
  // blank). Agency Comm / GST are inserted right before Total Cost, in that order, ONLY when
  // the client has that percentage set — see applyAdinnDynamicColumns in excelTemplateEngine.js.
  removeColumns: ['O', 'P'],
  feeColumnsBeforeAnchor: 'M',
};

// ROTN — real uploaded proposal file used as master. Header row1, one blank spacer row2,
// then one 2-row BLOCK per site (row1-of-block = main record, row2-of-block = a secondary
// W/H pair only — no matching Site field, left blank/zeroed), one blank spacer row, then the
// template's own "Grand Total" row (label in column E) with SUM formulas already starting
// from row2 (the spacer), not row3.
const ROTN_EXCEL_1 = {
  mode: 'block-per-site',
  blockSize: 2,
  firstDataRow: 3,
  lastDataRow: 76,
  totalRow: 78,
  sumRangeStartRow: 2,
  spacerRowsBeforeTotal: 1,
  totalLabelCell: 'E',
  totalLabelText: 'Grand Total',
  hasTermsAfterTotal: false,
  columns: {
    siNo: 'A', state: 'B', city: 'C', media: 'D', location: 'E', type: 'F', width: 'G', height: 'H',
    qty: 'J', durationDays: 'L', displayCostPerMonth: 'M', printingCost: 'O', mountingCost: 'P',
  },
  // No Site field maps to the block's 2nd row (a secondary W/H pair) — rather than leave it
  // visible with a meaningless "0  0", it's deleted outright (collapseSecondaryRows in
  // excelTemplateEngine.js). Both of these per-row formulas self-reference their own row and
  // must be rewritten once collapsing settles each site onto its final row number — Area (K)
  // originally folds in the deleted row's W/H too; Display Duration Cost (N) is an Excel
  // "shared formula" whose literal host text would otherwise keep pointing at its pre-collapse
  // row.
  secondaryRowColumns: ['G', 'H'],
  collapseSecondaryRow: true,
  selfReferencingFormulas: [
    { column: 'K', build: (r, cols) => `${cols.width}${r}*${cols.height}${r}` },
    { column: 'N', build: (r, cols) => `${cols.displayCostPerMonth}${r}/30*${cols.durationDays}${r}` },
  ],
  totalColumns: ['J', 'K', 'M', 'N', 'O', 'P', 'Q'],
};

// Jagran — both real uploaded proposal files share the exact same column layout (header
// text, order, and formulas are identical); they only differ in how many site rows the
// specific uploaded file happens to have, and where its own total row therefore sits.
// Row1 blank, row2-3 two-line header (Sizes merged into W/H sub-columns), row4 spacer/
// border, data from row5, one blank spacer row, then the template's own total row (no text
// label — just SUM formulas, already present in both files).
const JAGRAN_COLUMNS = {
  siNo: 'A', state: 'B', city: 'C', media: 'D', location: 'E', width: 'F', height: 'G',
  qty: 'H', type: 'K', displayCostPerMonth: 'L', durationDays: 'M', mountingCost: 'O',
  printingCost: 'P',
};
const JAGRAN_TOTAL_COLUMNS = ['H', 'I', 'L', 'N', 'O', 'P', 'Q', 'R', 'S'];

const JAGRAN_EXCEL_1 = {
  mode: 'row-per-site',
  firstDataRow: 5,
  lastDataRow: 87,
  totalRow: 89,
  sumRangeStartRow: null,
  spacerRowsBeforeTotal: 1,
  hasTermsAfterTotal: false,
  columns: JAGRAN_COLUMNS,
  totalColumns: JAGRAN_TOTAL_COLUMNS,
};

const JAGRAN_EXCEL_2 = {
  ...JAGRAN_EXCEL_1,
  lastDataRow: 9,
  totalRow: 11,
};

const EXCEL_CONFIGS = {
  generic: ADINN_EXCEL_1,
  'adinn-excel-1': ADINN_EXCEL_1,
  'rotn-excel-1': ROTN_EXCEL_1,
  'jagran-excel-1': JAGRAN_EXCEL_1,
  'jagran-excel-2': JAGRAN_EXCEL_2,
};

function getExcelConfig(formatKey) {
  return EXCEL_CONFIGS[formatKey] || EXCEL_CONFIGS.generic;
}

module.exports = { EXCEL_CONFIGS, getExcelConfig };
