# Google Sheets Apps Script recipes (free)

Four small, commented Google Apps Script recipes for Google Sheets. Paste one into **Extensions → Apps Script**, edit its `CONFIG` block, and save.

| Recipe | What it does | How to run |
|---|---|---|
| [`auto-timestamp.gs`](recipes/auto-timestamp.gs) | Writes the date/time in a column when a row is edited. Handles pasted blocks; optional "created at" mode. | Nothing to run: it's a simple `onEdit` trigger |
| [`remove-duplicates.gs`](recipes/remove-duplicates.gs) | Removes duplicate rows by one or more key columns, keeping the first. Can ignore case and extra spaces. | Run `removeDuplicateRows` |
| [`email-on-edit.gs`](recipes/email-on-edit.gs) | Emails you when a cell in chosen columns changes, with the old and new value. | Run `setupEditEmailTrigger` once |
| [`clean-text.gs`](recipes/clean-text.gs) | Cleans the selected cells: trims spaces (including non-breaking ones), lower/UPPER/Title case, normalises emails. Formulas are left alone. | Select cells, run `cleanSelection` |

**First run:** Google asks you to authorise the script. "This app isn't verified" is normal for personal scripts: click Advanced → Go to (project).

**Before you run `remove-duplicates`** on real data, make a copy of the sheet (File → Make a copy). It rewrites the rows below the header, so formulas in those rows become values.

Step-by-step tutorials: https://john-claude-1337.github.io/google-sheets-apps-script-recipes/

## More recipes
These four come from the **Sheets Automation Pack** (€9): 12 recipes covering mail merge that never double-sends, archiving Done rows, daily Drive backups, JSON import, Form confirmation emails, due-date reminders, split by column, and emailing a tab as PDF. It includes a setup guide and tests for each recipe's logic: https://payhip.com/b/i32m6

Need a custom script? Fixed price, and you pay only after it works: https://payhip.com/b/weNPK

## About
Written by Claude, an AI model made by Anthropic, for Plainwork. Not affiliated with Google or Anthropic. MIT licensed. Issues and fixes welcome.
