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
  // Media Type (right after Media Code) and Latitude/Longitude (right after Site Status) are
  // NOT part of the uploaded master.xlsx at all — applyFixedExtraColumns in
  // excelTemplateEngine.js inserts all three fresh at generation time, in this exact order
  // (each `before` letter is evaluated against the sheet AFTER the previous entry already
  // shifted things). Every letter below is the FINAL position once all three insertions have
  // happened — i.e. what the original template's own D onward (Location..Site Status) and
  // O/P (Vendor Name/Cost) become after being pushed right by them.
  fixedExtraColumns: [
    { before: 'D', headerText: 'Media Type' },
    { before: 'P', headerText: 'Latitude', decimalValue: true },
    { before: 'Q', headerText: 'Longitude', decimalValue: true },
  ],
  columns: {
    siNo: 'A', city: 'B', media: 'C', mediaType: 'D', location: 'E', qty: 'F', width: 'G', height: 'H',
    type: 'I', displayCostPerMonth: 'K', printingCost: 'L', mountingCost: 'M', siteStatus: 'O',
    latitude: 'P', longitude: 'Q',
  },
  // Column C's header text in the uploaded master says "Media" — shown as "Media Code" instead
  // (mediaId data itself is unchanged), so it reads clearly next to the new Media Type column.
  headerRenames: [{ cell: 'C1', text: 'Media Code' }],
  columnWidths: [
    { col: 'D', width: 14 },
    { col: 'P', width: 11 },
    { col: 'Q', width: 11 },
  ],
  // Area (J) and Total Cost (N) are both formulas in the ORIGINAL master file that
  // insertColumnBefore can't fix on its own — it moves the cell holding a formula, but never
  // rewrites the formula's own text, which still names the pre-shift column letters. Rewritten
  // here with the final post-insertion letters (see applyFixedExtraColumns's caller).
  // Total Cost is rewritten unconditionally (even when the client has no Agency Comm/GST %,
  // in which case applyAdinnDynamicColumns never runs and would otherwise leave this stale) —
  // when fees ARE present, applyAdinnDynamicColumns's own rewrite simply overwrites this again.
  selfReferencingFormulas: [
    { column: 'J', build: (r, cols) => `${cols.qty}${r}*${cols.width}${r}*${cols.height}${r}` },
    { column: 'N', build: (r, cols) => `${cols.displayCostPerMonth}${r}+${cols.printingCost}${r}+${cols.mountingCost}${r}` },
  ],
  totalColumns: ['F', 'J', 'K', 'L', 'M', 'N'],
  // Vendor Name/Vendor Cost are unused — dropped from the sheet entirely (not just left
  // blank). Agency Comm / GST are inserted right before Total Cost, in that order, ONLY when
  // the client has that percentage set — see applyAdinnDynamicColumns in excelTemplateEngine.js.
  removeColumns: ['R', 'S'],
  feeColumnsBeforeAnchor: 'N',
  feeRangeStartCol: 'K',
  feeRangeBaseEndCol: 'M',
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
// Media Type (right after Media/D — confirmed via the real file's own header text, "Media",
// at D1 with NO row1:row2 merge anywhere in the whole sheet) and Latitude/Longitude (right
// after Total Cost, before "Slot Loop") are not part of the real uploaded file at all —
// inserted fresh at generation time, same mechanism as Adinn/Jagran. Every letter below is the
// FINAL position once Media Type has already shifted everything from Location onward one
// column right.
const ROTN_EXCEL_1 = {
  mode: 'block-per-site',
  blockSize: 2,
  firstDataRow: 3,
  lastDataRow: 76,
  totalRow: 78,
  sumRangeStartRow: 2,
  spacerRowsBeforeTotal: 1,
  totalLabelCell: 'F',
  totalLabelText: 'Grand Total',
  hasTermsAfterTotal: false,
  // The real file's header lives on row1 alone with no merge at all (confirmed against the
  // actual sheet XML) — unlike Adinn/Jagran, which both merge a header cell down into row2.
  fixedExtraColumnHeaderRow: 1,
  fixedExtraColumnBlankHeaderRow: null,
  fixedExtraColumns: [
    { before: 'E', headerText: 'Media Type' },
    { before: 'S', headerText: 'Latitude', decimalValue: true },
    { before: 'T', headerText: 'Longitude', decimalValue: true },
  ],
  // The uploaded file's own header cell literally says "Media" — shown as "Media Code" instead
  // (mediaId data itself is unchanged), so it reads clearly next to the new Media Type column.
  headerRenames: [{ cell: 'D1', text: 'Media Code' }],
  columnWidths: [
    { col: 'E', width: 14 },
    { col: 'S', width: 11 },
    { col: 'T', width: 11 },
  ],
  columns: {
    siNo: 'A', state: 'B', city: 'C', media: 'D', mediaType: 'E', location: 'F', type: 'G', width: 'H', height: 'I',
    qty: 'K', durationDays: 'M', displayCostPerMonth: 'N', printingCost: 'P', mountingCost: 'Q',
    latitude: 'S', longitude: 'T',
  },
  // No Site field maps to the block's 2nd row (a secondary W/H pair) — rather than leave it
  // visible with a meaningless "0  0", it's deleted outright (collapseSecondaryRows in
  // excelTemplateEngine.js). Both of these per-row formulas self-reference their own row and
  // must be rewritten once collapsing settles each site onto its final row number — Area (L,
  // shifted from K) originally folds in the deleted row's W/H too; Display Duration Cost (O,
  // shifted from N) is an Excel "shared formula" whose literal host text would otherwise keep
  // pointing at its pre-collapse row.
  secondaryRowColumns: ['H', 'I'],
  collapseSecondaryRow: true,
  selfReferencingFormulas: [
    { column: 'L', build: (r, cols) => `${cols.width}${r}*${cols.height}${r}` },
    { column: 'O', build: (r, cols) => `${cols.displayCostPerMonth}${r}/30*${cols.durationDays}${r}` },
  ],
  totalColumns: ['K', 'L', 'N', 'O', 'P', 'Q', 'R'],
  // The real uploaded ROTN file has NO Agency Comm/GST columns at all — unlike Adinn (whose
  // Vendor columns get replaced) or Jagran (whose fee columns already exist, fixed-percentage),
  // ROTN's Total Cost (R, shifted from Q) is simply Display Duration Cost + Printing + Mounting
  // (O+P+Q). Reuses the exact same "insert only when the client has that percentage set"
  // mechanism Adinn uses (applyAdinnDynamicColumns in excelTemplateEngine.js) —
  // feeRangeStartCol/feeRangeBaseEndCol describe ROTN's own (now-shifted) base cost range
  // (O:Q, not Adinn's J:L), and feeStyleIds reuse ROTN's own neighboring header/data/total-row
  // style ids so inserted cells match this file's look instead of Adinn's.
  feeColumnsBeforeAnchor: 'R',
  feeRangeStartCol: 'O',
  feeRangeBaseEndCol: 'Q',
  feeStyleIds: { header: 11, data: 25, total: 17 },
};

// Jagran — both real uploaded proposal files share the exact same column layout (header
// text, order, and formulas are identical); they only differ in how many site rows the
// specific uploaded file happens to have, and where its own total row therefore sits.
// Row1 blank, row2-3 two-line header (Sizes merged into W/H sub-columns), row4 spacer/
// border, data from row5, one blank spacer row, then the template's own total row (no text
// label — just SUM formulas, already present in both files).
// Media Type (right after Media Code) and Latitude/Longitude are NOT part of either real
// uploaded Jagran file at all — applyFixedExtraColumns in excelTemplateEngine.js inserts them
// fresh at generation time (same mechanism Adinn uses). Every letter below is the FINAL
// position once Media Type has already shifted everything from Location onward one column
// right (real files inspected directly — column D holds mediaId with the ORIGINAL template
// header mislabelled "Media Vehicle", not the site's actual mediaType; Media Type must insert
// AFTER Media/D, i.e. before Location/E, not before Media itself).
const JAGRAN_COLUMNS = {
  siNo: 'A', state: 'B', city: 'C', media: 'D', mediaType: 'E', location: 'F', width: 'G', height: 'H',
  qty: 'I', type: 'L', displayCostPerMonth: 'M', durationDays: 'N', mountingCost: 'P',
  printingCost: 'Q',
  // Format 1's own position (right after Total (Incl. All)) — format 2 overrides both to 'V'/'W'
  // since its Latitude/Longitude sit right after Availability instead.
  latitude: 'U', longitude: 'V',
};
const JAGRAN_TOTAL_COLUMNS = ['I', 'J', 'M', 'O', 'P', 'Q', 'R', 'S', 'T'];
// Sq. ft (J) and Cost as per Duration (O) are both formulas in the ORIGINAL real files
// (confirmed by inspecting them directly: `=F5*G5*H5` and `=L5/30*M5`) that insertColumnBefore
// can't fix on its own — it moves the cell holding a formula, but never rewrites the formula's
// own text, which still names the pre-shift Width/Height/Qty/Display Cost/Duration letters.
// Rewritten here with the final post-insertion letters. Mounting/Printing are ALSO formulas in
// the real files, but those get unconditionally overwritten with the site's real cost value by
// fillRowPerSite anyway, so their stale formula text is harmless and left alone.
const JAGRAN_SELF_REFERENCING_FORMULAS = [
  { column: 'J', build: (r, cols) => `${cols.width}${r}*${cols.height}${r}*${cols.qty}${r}` },
  { column: 'O', build: (r, cols) => `${cols.displayCostPerMonth}${r}/30*${cols.durationDays}${r}` },
];

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
  selfReferencingFormulas: JAGRAN_SELF_REFERENCING_FORMULAS,
  // The uploaded template's own header cell says "Media Vehicle" — column D actually holds the
  // site's mediaId, so this reads "Media Code" (matching Adinn's own naming), not "Media Type"
  // (the new column inserted right after it, via fixedExtraColumns below, gets that label).
  headerRenames: [{ cell: 'D2', text: 'Media Code' }],
  // Both real files' actual header TEXT lives in row2, merged with row3 below it (row1 is a
  // genuinely empty, unused spacer row for every existing column — confirmed against the real
  // file: every single-value header column merges row2:row3, e.g. D2:D3, F2:F3, T2:T3; only the
  // "Sizes" pair merges horizontally instead, G2:H2, since G3/H3 hold their own "W"/"H" labels).
  fixedExtraColumnHeaderRow: 2,
  fixedExtraColumnBlankHeaderRow: 3,
  fixedExtraColumns: [
    { before: 'E', headerText: 'Media Type' },
    // Format 1 has no Availability/Rationale column — Total (Incl. All)/T is the last real
    // column, so these two just append as brand-new trailing columns.
    { before: 'U', headerText: 'Latitude', decimalValue: true },
    { before: 'V', headerText: 'Longitude', decimalValue: true },
  ],
  // "Sq. ft" (J, shifted from I) ships narrow (~7.3) and needs widening for its own per-row
  // value. "Display Cost per month" (M, shifted from L, ~18.4) and "Cost as per Duration" (O,
  // shifted from N, ~16.1) are already wide enough for a single site's own value, but the Total
  // row uses an accounting number format (reserves extra invisible padding for sign/parens
  // alignment) in a bolder font, which needs more room than any individual row — so these are
  // widened PAST their original size, not down to some flat value, or the fix would make the
  // Total row worse, not better.
  columnWidths: [
    { col: 'E', width: 14 },
    { col: 'J', width: 12 },
    { col: 'M', width: 22 },
    { col: 'O', width: 20 },
    { col: 'U', width: 11 },
    { col: 'V', width: 11 },
  ],
  // Unlike Adinn/ROTN, the real uploaded file already has these two fee columns natively baked
  // in (R: "Agency comm. @ 2%", S: "GST @ 18%", T: "Total (Incl. All)") — always shown, always
  // computed with that FIXED percentage regardless of who the client actually is. This makes
  // them conditional (removed when the client has no such percentage set, rewritten with the
  // client's real percentage — both the header label and every row's formula — when they do),
  // matching Adinn's own conditional rule, via applyConditionalFeeColumns in
  // excelTemplateEngine.js. Order matters: Agency Comm before GST, same as the file's own layout.
  conditionalFeeColumns: [
    { key: 'agencyComm', col: 'R', label: 'Agency comm.' },
    { key: 'gst', col: 'S', label: 'GST' },
  ],
  conditionalFeeBaseColumns: ['O', 'P', 'Q'],
  conditionalFeeTotalCol: 'T',
};

const JAGRAN_EXCEL_2 = {
  ...JAGRAN_EXCEL_1,
  lastDataRow: 9,
  totalRow: 11,
  // Format 2's sheet has two extra columns format 1 doesn't (Availability, Rationale) —
  // Availability maps to the site's current status (Available/Booked/Blocked). Rationale maps
  // to the site's linked SiteInfo master data's description (blank when none linked) — cleared
  // via the rationale-specific fillRowPerSite fix-up in excelTemplateEngine.js rather than left
  // showing the master file's own unrelated leftover example text for that row. Latitude/
  // Longitude go right after Availability (not appended at the very end like format 1), so
  // Rationale ends up shifted two columns further right than it otherwise would be.
  columns: { ...JAGRAN_COLUMNS, siteStatus: 'U', latitude: 'V', longitude: 'W', rationale: 'X' },
  fixedExtraColumns: [
    { before: 'E', headerText: 'Media Type' },
    { before: 'V', headerText: 'Latitude', decimalValue: true },
    { before: 'W', headerText: 'Longitude', decimalValue: true },
  ],
  columnWidths: [...JAGRAN_EXCEL_1.columnWidths.filter((c) => c.col !== 'U' && c.col !== 'V'), { col: 'V', width: 11 }, { col: 'W', width: 11 }],
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
