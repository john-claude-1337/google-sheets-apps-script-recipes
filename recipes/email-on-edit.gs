// From the Sheets Automation Pack by Plainwork. Written by Claude (an AI model made by Anthropic).
// MIT licensed. Not affiliated with Google or Anthropic.

// =====================================================================
// RECIPE 1 — Email me when a cell in a watched column changes
// Trigger: installable "On edit" (simple onEdit can't send email). Run setupEditEmailTrigger() once.
// =====================================================================
const R1_CONFIG = {
  sheetName: 'Sheet1',          // tab to watch
  watchColumns: [3, 4],         // 1-based column numbers (3 = C)
  headerRows: 1,                // edits in these top rows are ignored
  notifyEmail: 'you@example.com',
  subjectPrefix: '[Sheet change]',
};

function r1_shouldNotify(sheetName, row, col, cfg) {
  return sheetName === cfg.sheetName && row > cfg.headerRows && cfg.watchColumns.indexOf(col) !== -1;
}

function r1_buildMessage(sheetName, a1, oldValue, newValue, url) {
  const was = oldValue === undefined ? '(empty or multi-cell edit)' : String(oldValue);
  return 'Cell ' + sheetName + '!' + a1 + ' changed.\n\nOld value: ' + was +
    '\nNew value: ' + (newValue === undefined ? '(empty)' : String(newValue)) + '\n\nOpen: ' + url;
}

function r1_onEditEmail(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (!r1_shouldNotify(sheet.getName(), range.getRow(), range.getColumn(), R1_CONFIG)) return;
  const body = r1_buildMessage(sheet.getName(), range.getA1Notation(), e.oldValue, e.value,
    e.source.getUrl());
  MailApp.sendEmail(R1_CONFIG.notifyEmail, R1_CONFIG.subjectPrefix + ' ' + range.getA1Notation(), body);
}

function setupEditEmailTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'r1_onEditEmail')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('r1_onEditEmail').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}
