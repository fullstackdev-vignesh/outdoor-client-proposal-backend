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
  // Area/Display/Printing/Mounting (style 11) and Total Cost (style 14) ship in "General"
  // format (plain "250000", no grouping) — reformatted to Indian comma grouping
  // ("2,50,000"). 22/23 are the same columns' styles in the moved Total row (the border-row
  // style), and 14/22 are what applyAdinnDynamicColumns reuses for the Agency Comm/GST cells
  // it inserts, so this list also covers those automatically.
  indianCommaStyleIds: [11, 14, 22, 23],
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
  // The real uploaded ROTN file has NO Agency Comm/GST columns at all — unlike Adinn (whose
  // Vendor columns get replaced) or Jagran (whose fee columns already exist, fixed-percentage),
  // ROTN's Total Cost (Q) is simply Display Duration Cost + Printing + Mounting (N+O+P). Reuses
  // the exact same "insert only when the client has that percentage set" mechanism Adinn uses
  // (applyAdinnDynamicColumns in excelTemplateEngine.js) — feeRangeStartCol/feeRangeBaseEndCol
  // describe ROTN's own base cost range (N:P, not Adinn's J:L), and feeStyleIds reuse ROTN's own
  // neighboring header/data/total-row style ids so inserted cells match this file's look instead
  // of Adinn's.
  feeColumnsBeforeAnchor: 'Q',
  feeRangeStartCol: 'N',
  feeRangeBaseEndCol: 'P',
  feeStyleIds: { header: 11, data: 25, total: 17 },
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
  // The uploaded template's own header cell says "Media Vehicle" — shown as "Media Type"
  // instead, without touching the original uploaded file (only this generated copy).
  headerRenames: [{ cell: 'D2', text: 'Media Type' }],
  // "Sq. ft" (I) ships narrow (~7.3) and needs widening for its own per-row value. "Display
  // Cost per month" (L, ~18.4) and "Cost as per Duration" (N, ~16.1) are already wide enough
  // for a single site's own value, but the Total row uses an accounting number format
  // (reserves extra invisible padding for sign/parens alignment) in a bolder font, which
  // needs more room than any individual row — so these are widened PAST their original size,
  // not down to some flat value, or the fix would make the Total row worse, not better.
  columnWidths: [
    { col: 'I', width: 12 },
    { col: 'L', width: 22 },
    { col: 'N', width: 20 },
  ],
  // Unlike Adinn/ROTN, the real uploaded file already has these two fee columns natively baked
  // in (Q: "Agency comm. @ 2%", R: "GST @ 18%", S: "Total (Incl. All)") — always shown, always
  // computed with that FIXED percentage regardless of who the client actually is. This makes
  // them conditional (removed when the client has no such percentage set, rewritten with the
  // client's real percentage — both the header label and every row's formula — when they do),
  // matching Adinn's own conditional rule, via applyConditionalFeeColumns in
  // excelTemplateEngine.js. Order matters: Agency Comm before GST, same as the file's own layout.
  conditionalFeeColumns: [
    { key: 'agencyComm', col: 'Q', label: 'Agency comm.' },
    { key: 'gst', col: 'R', label: 'GST' },
  ],
  conditionalFeeBaseColumns: ['N', 'O', 'P'],
  conditionalFeeTotalCol: 'S',
};

const JAGRAN_EXCEL_2 = {
  ...JAGRAN_EXCEL_1,
  lastDataRow: 9,
  totalRow: 11,
  // Format 2's sheet has two extra columns format 1 doesn't (T Availability, U Rationale) —
  // Availability maps to the site's current status (Available/Booked/Blocked). Rationale maps
  // to the site's linked SiteInfo master data's description (blank when none linked) — cleared
  // via the rationale-specific fillRowPerSite fix-up in excelTemplateEngine.js rather than left
  // showing the master file's own unrelated leftover example text for that row.
  columns: { ...JAGRAN_COLUMNS, siteStatus: 'T', rationale: 'U' },
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
