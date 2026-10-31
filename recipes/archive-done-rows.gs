// From the Sheets Automation Pack by Plainwork. Written by Claude (an AI model made by Anthropic).
// MIT licensed. Not affiliated with Google or Anthropic.

// =====================================================================
// RECIPE 4 — Move rows to an archive tab when a status column says "Done"
// Run manually, or add a time trigger with setupArchiveTrigger().
// =====================================================================
const R4_CONFIG = {
  sourceSheet: 'Tasks',
  archiveSheet: 'Archive',     // created if missing
  statusColumn: 4,             // 1-based
  doneValues: ['done', 'closed', 'complete'], // case-insensitive
  headerRows: 1,
};

function r4_split(rows, cfg) {
  const keep = [], move = [];
  const done = cfg.doneValues.map(v => v.toLowerCase());
  rows.forEach(r => (done.indexOf(String(r[cfg.statusColumn - 1]).trim().toLowerCase()) !== -1 ? move : keep).push(r));
  return { keep: keep, move: move };
}

function archiveDoneRows() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(R4_CONFIG.sourceSheet);
  const dst = ss.getSheetByName(R4_CONFIG.archiveSheet) || ss.insertSheet(R4_CONFIG.archiveSheet);
  const lastRow = src.getLastRow(), lastCol = src.getLastColumn();
  if (lastRow <= R4_CONFIG.headerRows) return;
  const body = src.getRange(R4_CONFIG.headerRows + 1, 1, lastRow - R4_CONFIG.headerRows, lastCol);
  const parts = r4_split(body.getValues(), R4_CONFIG);
  if (parts.move.length === 0) return;
  dst.getRange(dst.getLastRow() + 1, 1, parts.move.length, lastCol).setValues(parts.move);
  body.clearContent();
  if (parts.keep.length) src.getRange(R4_CONFIG.headerRows + 1, 1, parts.keep.length, lastCol).setValues(parts.keep);
}

function setupArchiveTrigger() {
  ScriptApp.newTrigger('archiveDoneRows').timeBased().everyDays(1).atHour(6).create();
}
