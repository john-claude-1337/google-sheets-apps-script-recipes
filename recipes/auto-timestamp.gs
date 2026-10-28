// From the Sheets Automation Pack by Plainwork. Written by Claude (an AI model made by Anthropic).
// MIT licensed. Not affiliated with Google or Anthropic.

// =====================================================================
// RECIPE 2 — Auto-timestamp: write the date/time in a column when a row is edited
// Trigger: simple onEdit (no setup). If you already have an onEdit, call r2_timestamp(e) from it.
// =====================================================================
const R2_CONFIG = {
  sheetName: 'Sheet1',
  watchColumns: [1, 2, 3],   // edits in these columns update the stamp
  stampColumn: 5,            // where the timestamp goes (5 = E)
  headerRows: 1,
  onlyFirstTime: false,      // true = never overwrite an existing stamp ("created at")
};

function r2_shouldStamp(sheetName, row, col, numCols, cfg) {
  if (sheetName !== cfg.sheetName || row <= cfg.headerRows) return false;
  // a pasted block touches columns col .. col+numCols-1
  for (let c = col; c < col + numCols; c++) {
    if (c !== cfg.stampColumn && cfg.watchColumns.indexOf(c) !== -1) return true;
  }
  return false;
}

function r2_timestamp(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (!r2_shouldStamp(sheet.getName(), range.getRow(), range.getColumn(), range.getNumColumns(), R2_CONFIG)) return;
  const stamps = sheet.getRange(range.getRow(), R2_CONFIG.stampColumn, range.getNumRows(), 1);
  const current = stamps.getValues();
  const now = new Date();
  stamps.setValues(current.map(r => [R2_CONFIG.onlyFirstTime && r[0] !== '' ? r[0] : now]));
}

function onEdit(e) {
  r2_timestamp(e);
}
