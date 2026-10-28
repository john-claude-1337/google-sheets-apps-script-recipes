// From the Sheets Automation Pack by Plainwork. Written by Claude (an AI model made by Anthropic).
// MIT licensed. Not affiliated with Google or Anthropic.

// =====================================================================
// RECIPE 12 — Clean up text in a range: trim spaces, fix case, normalise emails
// Works on the current selection, so select the cells first, then run cleanSelection().
// =====================================================================
const R12_CONFIG = {
  mode: 'trim',   // 'trim' | 'lower' | 'upper' | 'title' | 'email'
};

function r12_clean(v, mode) {
  if (typeof v !== 'string') return v;
  let s = v.replace(/[   ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (mode === 'lower') s = s.toLowerCase();
  else if (mode === 'upper') s = s.toUpperCase();
  else if (mode === 'title') s = s.toLowerCase().replace(/(^|[\s\-'])(\S)/g, (m, p, c) => p + c.toUpperCase());
  else if (mode === 'email') s = s.replace(/\s/g, '').toLowerCase();
  return s;
}

function cleanSelection() {
  const range = SpreadsheetApp.getActiveRange();
  const formulas = range.getFormulas();
  const values = range.getValues();
  // never overwrite formulas
  const out = values.map((row, i) => row.map((v, j) => (formulas[i][j] ? formulas[i][j] : r12_clean(v, R12_CONFIG.mode))));
  range.setValues(out);
}
