// From the Sheets Automation Pack by Plainwork. Written by Claude (an AI model made by Anthropic).
// MIT licensed. Not affiliated with Google or Anthropic.

// =====================================================================
// RECIPE 5 — Remove duplicate rows by one or more key columns (keeps the first)
// Note: Data → Data cleanup → Remove duplicates does this by hand; this version is repeatable
// and can ignore case/spaces.
// =====================================================================
const R5_CONFIG = {
  sheetName: 'Sheet1',
  keyColumns: [2],         // 1-based; e.g. [2] = email column, [1,3] = name + date
  headerRows: 1,
  ignoreCaseAndSpaces: true,
};

function r5_dedupe(rows, cfg) {
  const seen = {};
  const kept = [];
  rows.forEach(r => {
    let key = cfg.keyColumns.map(c => String(r[c - 1])).join('\u0001');
    if (cfg.ignoreCaseAndSpaces) key = key.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!seen[key]) { seen[key] = true; kept.push(r); }
  });
  return kept;
}

function removeDuplicateRows() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(R5_CONFIG.sheetName);
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow <= R5_CONFIG.headerRows) return;
  const range = sheet.getRange(R5_CONFIG.headerRows + 1, 1, lastRow - R5_CONFIG.headerRows, lastCol);
  const rows = range.getValues();
  const kept = r5_dedupe(rows, R5_CONFIG);
  range.clearContent();
  if (kept.length) sheet.getRange(R5_CONFIG.headerRows + 1, 1, kept.length, lastCol).setValues(kept);
  SpreadsheetApp.getActive().toast('Removed ' + (rows.length - kept.length) + ' duplicate row(s).');
}
