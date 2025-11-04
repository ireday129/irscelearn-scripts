/** REPORTING ISSUE CLEANUP & DATA VALIDATION (ONLY runs on Master/Clean) **/

/**
 * Applies data validation and conditional formatting for Reporting Issue?
 * columns on the Clean and Master sheets.
 * NOTE: The old Reporting Issue sheet is ignored/deprecated.
 */
function applyReportingIssueValidationAndFormatting_() {
  const ss = SpreadsheetApp.getActive();
  // Assume CFG.REPORTING_ISSUE_CHOICES is defined globally.
  const choices = typeof CFG !== 'undefined' && CFG.REPORTING_ISSUE_CHOICES ? CFG.REPORTING_ISSUE_CHOICES : [];

  const clean = ss.getSheetByName(CFG.SHEET_CLEAN);
  if (clean && clean.getLastRow() >= 1) {
    const ch = clean.getRange(1,1,1, clean.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
    const iRI = ch.indexOf('Reporting Issue?');
    if (iRI >= 0) {
      const range = clean.getRange(2, iRI+1, Math.max(clean.getMaxRows()-1, 1), 1);
      setDropdown_(range, choices, true);
      setIssueColors_(clean, iRI+1);
    }
  }

  const master = ss.getSheetByName(CFG.SHEET_MASTER);
  if (master && master.getLastRow() >= 1) {
    const mh = master.getRange(1,1,1, master.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
    const iRI = mh.indexOf(CFG.COL_HEADERS.masterIssueCol);
    const iRep = mh.indexOf(CFG.COL_HEADERS.reportedCol);
    if (iRI >= 0) {
      const range = master.getRange(2, iRI+1, Math.max(master.getMaxRows()-1,1), 1);
      setDropdown_(range, choices, false);
      setIssueColors_(master, iRI+1);
    }
    if (iRep >= 0) {
      setReportedCheckboxColors_(master, iRep + 1);
    }
  }
}

/** Stub: Clears the body of the Reporting Issue sheet (now deprecated, but kept for full clarity) */
function writeIssuesDataOnly_(ss, issueRows) {
  // This function is obsolete as the sheet is deleted.
}

/** Stub: Appends rows to the Reporting Issue sheet (now obsolete) */
function appendIssuesRows_(ss, issues) {
  // This function is obsolete as the sheet is deleted.
}

// --- Supporting Utilities (Needed locally for data validation/formatting) ---

/** Sets data validation dropdown rule on a given range. */
function setDropdown_(range, list, allowBlank) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)
    .setHelpText('Choose a reporting issue or leave blank if none.')
    .build();
  range.setDataValidation(rule);
}

/** Sets conditional formatting rules for issue statuses. */
function setIssueColors_(sheet, colIdx1) {
  // Assuming CFG.REPORTING_ISSUE_CHOICES is defined globally
  const choices = typeof CFG !== 'undefined' && CFG.REPORTING_ISSUE_CHOICES ? CFG.REPORTING_ISSUE_CHOICES : [];
  if (choices.length === 0) return;

  const last = Math.max(sheet.getLastRow(), 2);
  const rng = sheet.getRange(2, colIdx1, last-1, 1);
  const rules = sheet.getConditionalFormatRules().filter(r=>{
    const rs = r.getRanges();
    // Filter out existing rules on this column index
    return !(rs && rs.length && rs[0].getColumn()===colIdx1);
  });

  const colorMap = {
    'PTIN does not exist': { bg: '#2196f3', fg: '#ffffff' }, // Blue
    'PTIN & name do not match': { bg: '#f44336', fg: '#ffffff' }, // Red
    'Missing PTIN': { bg: '#ffeb3b', fg: '#000000' }, // Yellow
    'Other': { bg: '#9e9e9e', fg: '#000000' } // Grey
  };

  const add = (text, bg, fg) => rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text)
      .setBackground(bg)
      .setFontColor(fg)
      .setBold(true)
      .setRanges([rng])
      .build()
  );

  choices.forEach(issueText => {
    const mapEntry = colorMap[issueText];
    if (mapEntry) {
      add(issueText, mapEntry.bg, mapEntry.fg);
    }
  });

  sheet.setConditionalFormatRules(rules);
}

/**
 * Add conditional formatting for a checkbox column so that TRUE values
 * show with a green checkmark. Background is left unchanged.
 *
 * @param {Sheet} sheet
 * @param {number} colIdx1 1-based column index of the Reported? column
 */
function setReportedCheckboxColors_(sheet, colIdx1) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const numRows = lastRow - 1;
  if (numRows <= 0) return;

  const colRange = sheet.getRange(2, colIdx1, numRows, 1);

  // Preserve other rules; drop any that directly target this column starting at row 2
  const existing = sheet.getConditionalFormatRules() || [];
  const rules = existing.filter(r => {
    const rs = r.getRanges();
    if (!rs || !rs.length) return true;
    const first = rs[0];
    return !(first.getColumn() === colIdx1 && first.getRow() === 2);
  });

  // Use a formula that refers to the cell itself so the rule is reusable down the column.
  const greenCheckRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=INDIRECT(ADDRESS(ROW(), COLUMN()))=TRUE')
    .setFontColor('#1e8e3e') // green checkmark
    .setRanges([colRange])
    .build();

  rules.push(greenCheckRule);
  sheet.setConditionalFormatRules(rules);
}

/** Stub for dependency: Placeholder for checking System Reporting Issues sheet. */
function getIssuesSheet_(ss) {
  // Since the primary issue sheet is deleted, this stub returns null or the System Issues sheet.
  // The primary logic here is that functions expecting the old sheet will now fail gracefully.
  return ss.getSheetByName(CFG.SHEET_SYS_ISSUES);
}
