/** Map group "Reporting"/"Reporting Info" headers with common variants/typos (case-insensitive). */
function mapGroupHeadersFlexible_(sheet){
  const hdr = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const lower = hdr.map(h=>h.toLowerCase());

  function idxOfAny(names){
    for (const n of names){
      const j = lower.indexOf(String(n).toLowerCase());
      if (j>=0) return j;
    }
    return -1;
  }

  // Flexible aliases (case-insensitive)
  const map = {
    first:     idxOfAny(['attendee first name','first name','fname']),
    last:      idxOfAny(['attendee last name','last name','lname']),
    ptin:      idxOfAny(['attendee ptin','ptin']),
    email:     idxOfAny(['email','attendee email','e-mail']),
    prog:      idxOfAny(['program number','program #','program id']),
    progName:  idxOfAny(['program name','course name']),
    hours:     idxOfAny(['ce hours awarded','ce hours','hours','ce hour(s)']),
    comp:      idxOfAny(['program completion date','completion date','completed at','date completed']),
    issue:     idxOfAny(['reporting issue?','reporting issue']),
    reported:  idxOfAny(['reported?','reported']),
    reportedAt:idxOfAny(['reported at','date reported'])
  };

  // Sheet must have: First+Last, (PTIN or Email), and (Program Name or Program Number)
  const haveIdentity = (map.first >= 0 && map.last >= 0);
  const haveKey      = (map.ptin >= 0 || map.email >= 0);
  const haveProgram  = (map.progName >= 0 || map.prog >= 0);

  const missing = [];
  if (!haveIdentity) {
    if (map.first < 0) missing.push('Attendee First Name');
    if (map.last  < 0) missing.push('Attendee Last Name');
  }
  if (!haveKey) missing.push('Attendee PTIN or Email');
  if (!haveProgram) missing.push('Program Name or Program Number');

  return { ok: missing.length===0, missing, ...map };
}

/** Public entrypoint for the menu item: “Sync Group Sheets (strict)” */
function syncGroupSheets() {
  try {
    if (typeof syncGroupSheetsStrict === 'function') return syncGroupSheetsStrict();
    if (typeof doGroupSyncAll_ === 'function')       return doGroupSyncAll_();
    if (typeof runGroupSync === 'function')          return runGroupSync();
    if (typeof groupSyncMain === 'function')         return groupSyncMain();
    if (typeof syncGroupsFlexible === 'function')    return syncGroupsFlexible();

    const fns = Object.keys(this)
      .filter(k => typeof this[k] === 'function' && /group.*sync|sync.*group|doGroup/i.test(k))
      .sort();
    toast_(
      'No concrete group sync function found for syncGroupSheets(). ' +
      (fns.length ? 'Candidates: ' + fns.join(', ') : 'No group-sync-like functions detected.'),
      true
    );
  } catch (err) {
    toast_('syncGroupSheets failed: ' + (err && err.message ? err.message : err), true);
    Logger.log(err && err.stack ? err.stack : err);
  }
}

/** Alias used by menus that reference syncGroupSheetsMenu */
function syncGroupSheetsMenu() {
  return syncGroupSheets();
}

/** Ensure global exports in Apps Script runtime (defensive) */
try {
  this.syncGroupSheets = this.syncGroupSheets || syncGroupSheets;
  this.syncGroupSheetsMenu = this.syncGroupSheetsMenu || syncGroupSheetsMenu;
} catch (e) {
  // no-op
}

/**
 * === Strict Group Sync ===
 * Reads Master, filters rows by Group, and writes each group's rows to its linked sheet.
 * - Uses Courses sheet to convert Program Number -> Program Name.
 * - Accepts either Program Name or Program Number columns on target.
 * - Keeps Reporting Issue? rows (do NOT filter them out).
 * - Clears & rewrites the data body only (keeps header row).
 *
 * Requires:
 *  - CFG.SHEET_MASTER (e.g., "Master")
 *  - A catalog sheet named "Group Config" (preferred) with: Group ID, Group Name, Spreadsheet URL
 *    or legacy "Groups" with: Group, Sheet URL
 *  - A sheet named "Courses" with: Program Number, Program Name
 */
function syncGroupSheetsStrict() {
  const ss = SpreadsheetApp.getActive();

  // --- Load Master
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) { toast_('Master is empty; nothing to sync.', true); return; }
  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mm   = mapHeaders_(mHdr);
  if (mm.group == null) { toast_('Master is missing "Group" column.', true); return; }

  // Build courses map: Program Number -> Program Name
  const coursesMap = loadCoursesMap_(ss);

  // Read group catalog (Group Config preferred)
  const groups = readGroupsCatalog_(ss);
  if (!groups.length) { toast_('No groups found in the "Group Config"/"Groups" catalog.', true); return; }

  const mBody = mVals.slice(1);
  let totalSheets = 0, totalCells = 0, visited = 0;

  groups.forEach(entry => {
    visited++;
    const gName = String(entry.groupName || entry.group || '').trim();
    const gUrl  = String(entry.url || '').trim();
    const gId   = String(entry.groupId || '').trim();
    if (!gName || !gUrl) {
      Logger.log(`Group catalog entry missing name or url: ${JSON.stringify(entry)}`);
      return;
    }

    // Master "Group" is the ID column in this workbook
    const iGroupId   = (mm.group != null) ? mm.group : findHeaderIndex_(mHdr, 'Group ID');
    const iGroupName = findHeaderIndex_(mHdr, 'Group Name');

    // Filter Master rows for this group (prefer match by ID, fallback to name)
    const rows = mBody.filter(r => {
      const name = iGroupName >= 0 ? String(r[iGroupName] || '').trim() : '';
      const id   = iGroupId   >= 0 ? String(r[iGroupId]   || '').trim() : '';
      return (gId && id && id === gId) || (name && name === gName);
    });

    if (!rows.length) {
      Logger.log(`Group "${gName}": no Master rows; skipping.`);
      return;
    }

    // Open target spreadsheet
    let targetSS;
    try {
      targetSS = openSpreadsheetByUrlOrId_(gUrl);
    } catch (e) {
      Logger.log(`Group "${gName}": cannot open target (${gUrl}): ${e.message}`);
      return;
    }
    const targetSheet = targetSS.getSheets()[0]; // convention: first sheet
    if (!targetSheet) {
      Logger.log(`Group "${gName}": no sheets found in target spreadsheet.`);
      return;
    }

    // Ensure target headers are understood; accept Program Name OR Program Number
    const gMap = mapGroupHeadersFlexible_(targetSheet);
    if (!gMap.ok) {
      Logger.log(`Group "${gName}": missing headers: ${gMap.missing.join(', ')}`);
      toast_(`Group "${gName}" missing required headers: ${gMap.missing.join(', ')}`, true);
      return;
    }

    // Build output rows aligned to target header order
    const hdr = targetSheet.getRange(1,1,1,targetSheet.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
    const lower = hdr.map(h=>h.toLowerCase());
    const out = [];

    const wantDate = (v) => (v instanceof Date) ? v : (parseDate_(v) || v);

    rows.forEach(row => {
      const arr = new Array(hdr.length).fill('');

      const first = mm.firstName != null ? String(row[mm.firstName]||'').trim() : '';
      const last  = mm.lastName  != null ? String(row[mm.lastName] ||'').trim() : '';
      const ptin  = mm.ptin      != null ? formatPtinP0_(row[mm.ptin]||'')      : '';
      const email = mm.email     != null ? String(row[mm.email]    ||'').toLowerCase().trim() : '';
      const progN = mm.program   != null ? String(row[mm.program]  ||'').toUpperCase().replace(/\s+/g,'').trim() : '';
      const hrs   = mm.hours     != null ? row[mm.hours] : '';
      const comp  = mm.completion!= null ? wantDate(row[mm.completion]) : '';
      const issue = mm.masterIssueCol != null ? String(row[mm.masterIssueCol]||'').trim() : '';
      const reported    = mm.reportedCol    != null ? row[mm.reportedCol] : '';
      const reportedAt  = mm.reportedAtCol  != null ? wantDate(row[mm.reportedAtCol]) : '';

      const progName = coursesMap.get(progN) || '';

      // Place values by exact header text (case-insensitive lookup)
      placeByHeader_(arr, lower, 'Attendee First Name', first);
      placeByHeader_(arr, lower, 'Attendee Last Name',  last);
      placeByHeader_(arr, lower, 'Attendee PTIN',       ptin);
      placeByHeader_(arr, lower, 'Email',               email);

      if (gMap.progName >= 0) {
        placeByHeader_(arr, lower, 'Program Name', progName || progN);
      } else if (gMap.prog >= 0) {
        placeByHeader_(arr, lower, 'Program Number', progN);
      }

      placeByHeader_(arr, lower, 'CE Hours Awarded', hrs);
      placeByHeader_(arr, lower, 'CE Hours', hrs); // some sheets use this label
      placeByHeader_(arr, lower, 'Program Completion Date', comp);

      placeByHeader_(arr, lower, 'Reporting Issue?', issue);
      placeByHeader_(arr, lower, 'Reported?', truthy_(reported));
      placeByHeader_(arr, lower, 'Reported At', reportedAt);

      out.push(arr);
    });

    // Clear body & write
    const lastRow = targetSheet.getLastRow();
    const lastCol = targetSheet.getLastColumn();
    if (lastRow > 1) targetSheet.getRange(2,1,lastRow-1,lastCol).clearContent();

    if (out.length) {
      targetSheet.getRange(2,1,out.length, lastCol).setValues(out);

      // Date formatting if present
      const iComp = lower.indexOf('program completion date');
      if (iComp >= 0) targetSheet.getRange(2, iComp+1, out.length, 1).setNumberFormat('mm/dd/yyyy');
      const iRepAt = lower.indexOf('reported at');
      if (iRepAt >= 0) targetSheet.getRange(2, iRepAt+1, out.length, 1).setNumberFormat('mm/dd/yyyy');

      // NEW: Enforce checkbox on Reported? column
      const iRep = lower.indexOf('reported?');
      if (iRep >= 0) {
        const repRange = targetSheet.getRange(2, iRep+1, Math.max(out.length,1), 1);
        const rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
        repRange.setDataValidation(rule);
      }


      // ===== Enhancements for group sheets =====

      // 1) Freeze + protect the header row (use RANGE.protect to avoid sheet↔range mismatch)
      try {
        targetSheet.setFrozenRows(1);

        // Remove any existing Range protections that exactly cover row 1 (header)
        (targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || []).forEach(p => {
          const r = p.getRange && p.getRange();
          if (r && r.getRow() === 1 && r.getNumRows() === 1) p.remove();
        });

        // Create a RANGE protection directly on the header row
        const headerRange = targetSheet.getRange(1, 1, 1, lastCol);
        const headerProt  = headerRange.protect().setDescription('Protect header row');
        headerProt.removeEditors(headerProt.getEditors()); // keep only owner
        // Optional: allow viewers/editors to see warning instead of hard block
        // headerProt.setWarningOnly(true);
      } catch (e) {
        Logger.log('Header protect failed (non-fatal): ' + e.message);
      }

      // 2) Alternating row banding (neutral) — use RANGE.applyRowBanding
      try {
        (targetSheet.getBandings() || []).forEach(b => b.remove());
        const bandRange = targetSheet.getRange(1, 1, Math.max(targetSheet.getLastRow(), 2), lastCol);
        const band = bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
        band.setHeaderRowColor(null);
      } catch (e) {
        Logger.log('Banding failed (non-fatal): ' + e.message);
      }

      // 3) Auto-size columns + enforce sane formats
      try {
        // Turn off wrapping on header and data so widths reflect full titles/content
        const dataLastRow = Math.max(out.length, 1);
        targetSheet.getRange(1, 1, 1, lastCol).setWrap(false);
        if (dataLastRow > 0) {
          targetSheet.getRange(2, 1, dataLastRow, lastCol).setWrap(false);
        }

        // Auto-resize all columns (twice helps when formats update lengths)
        for (let pass = 0; pass < 2; pass++) {
          for (let c = 1; c <= lastCol; c++) targetSheet.autoResizeColumn(c);
        }

        // Hours number format (CE Hours Awarded / CE Hours)
        let iHours = lower.indexOf('ce hours awarded');
        if (iHours < 0) iHours = lower.indexOf('ce hours');
        if (iHours >= 0) {
          targetSheet.getRange(2, iHours + 1, Math.max(out.length, 1), 1).setNumberFormat('0');
        }
      } catch (e) {
        Logger.log('Autosize/format failed (non-fatal): ' + e.message);
      }

      // 4) Filter view on full table
      try {
        const prevFilter = targetSheet.getFilter();
        if (prevFilter) prevFilter.remove();
        targetSheet.getRange(1, 1, Math.max(targetSheet.getLastRow(), 2), lastCol).createFilter();
      } catch (e) {
        Logger.log('Filter creation failed (non-fatal): ' + e.message);
      }

      // --- Conditional formatting (reset to only our yellow issue rule + new green text rule for Reported?) ---
      try {
        const existing = targetSheet.getConditionalFormatRules() || [];
        const iIssue = lower.indexOf('reporting issue?');
        const iReported = lower.indexOf('reported?');
        const newRules = [];
        // 1) Yellow for nonblank Reporting Issue?
        if (iIssue >= 0) {
          const issueRange = targetSheet.getRange(2, iIssue + 1, Math.max(out.length, 1), 1);
          const issueColLetter = colToA1_(iIssue + 1);
          const issueFormula = `=LEN($${issueColLetter}${issueRange.getRow()})>0`;
          const issueRule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(issueFormula)
            .setBackground('#fff59d') // light yellow
            .setRanges([issueRange])
            .build();
          newRules.push(issueRule);
        }
        // 2) Green text for Reported? TRUE (background white)
        if (iReported >= 0) {
          const reportedRange = targetSheet.getRange(2, iReported + 1, Math.max(out.length, 1), 1);
          const reportedColLetter = colToA1_(iReported + 1);
          const rowStart = reportedRange.getRow();
          const reportedFormula = `=$${reportedColLetter}${rowStart}=TRUE`;
          const reportedRule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(reportedFormula)
            .setBackground('#ffffff')
            .setFontColor('#1e8e3e')
            .setRanges([reportedRange])
            .build();
          newRules.push(reportedRule);
        }
        // Overwrite any prior rules
        targetSheet.setConditionalFormatRules(newRules);
      } catch (e) {
        Logger.log('Conditional-format reset failed (non-fatal): ' + e.message);
      }

      // 6) Duplicate guard (same Attendee PTIN + Program Name) highlighted light red
      try {
        const iPTIN     = lower.indexOf('attendee ptin');
        const iProgName = lower.indexOf('program name');
        if (iPTIN >= 0 && iProgName >= 0) {
          const dataRange = targetSheet.getRange(2, 1, Math.max(out.length,1), lastCol);
          const aPTIN = colToA1_(iPTIN+1);
          const aProg = colToA1_(iProgName+1);
          const dupRule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=COUNTIFS($${aPTIN}:$${aPTIN},$${aPTIN}2,$${aProg}:$${aProg},$${aProg}2)>1`)
            .setBackground('#fde2e1')
            .setRanges([dataRange])
            .build();
          const rules = targetSheet.getConditionalFormatRules() || [];
          rules.push(dupRule);
          targetSheet.setConditionalFormatRules(rules);
        }
      } catch (e) {
        Logger.log('Duplicate CF failed (non-fatal): ' + e.message);
      }

      // 7) Email normalization (lowercase)
      try {
        const iEmail = lower.indexOf('email');
        if (iEmail >= 0) {
          const emailRange = targetSheet.getRange(2, iEmail+1, out.length, 1);
          const emails = emailRange.getValues().map(r => [String(r[0]||'').trim().toLowerCase()]);
          if (emails.length) emailRange.setValues(emails);
        }
      } catch (e) {
        Logger.log('Email normalize failed (non-fatal): ' + e.message);
      }

      // Shared sync timestamp for summary block
      const syncTimestamp = new Date();
      const syncTimestampStr = Utilities.formatDate(syncTimestamp, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');

      // 10) Summary block: totals + distinct attendee counts
      try {
        const iHours2  = (() => { const x = lower.indexOf('ce hours awarded'); return x >= 0 ? x : lower.indexOf('ce hours'); })();
        const iIssue2  = lower.indexOf('reporting issue?');
        const iEmail2  = lower.indexOf('email');
        const iPtin2   = lower.indexOf('attendee ptin');
        const iFirst2  = lower.indexOf('attendee first name');
        const iLast2   = lower.indexOf('attendee last name');
        const iRep2    = lower.indexOf('reported?');

        // Distinct attendees across the sheet
        const allAttendees = new Set();
        for (let r = 0; r < out.length; r++) {
          const emailVal = iEmail2 >= 0 ? String(out[r][iEmail2] || '').trim().toLowerCase() : '';
          const ptinVal  = iPtin2  >= 0 ? String(out[r][iPtin2]  || '').trim().toUpperCase() : '';
          const fVal     = iFirst2 >= 0 ? String(out[r][iFirst2] || '').trim().toLowerCase() : '';
          const lVal     = iLast2  >= 0 ? String(out[r][iLast2]  || '').trim().toLowerCase() : '';
          const k = emailVal || ptinVal || (fVal + '|' + lVal);
          if (k.replace(/\|/g,'').length) allAttendees.add(k);
        }

        // Distinct attendees who have an issue
        const issueAttendees = new Set();
        if (iIssue2 >= 0) {
          for (let r = 0; r < out.length; r++) {
            const issueVal = String(out[r][iIssue2] || '').trim();
            if (!issueVal) continue;
            const emailVal = iEmail2 >= 0 ? String(out[r][iEmail2] || '').trim().toLowerCase() : '';
            const ptinVal  = iPtin2  >= 0 ? String(out[r][iPtin2]  || '').trim().toUpperCase() : '';
            const fVal     = iFirst2 >= 0 ? String(out[r][iFirst2] || '').trim().toLowerCase() : '';
            const lVal     = iLast2  >= 0 ? String(out[r][iLast2]  || '').trim().toLowerCase() : '';
            const k = emailVal || ptinVal || (fVal + '|' + lVal);
            if (k.replace(/\|/g,'').length) issueAttendees.add(k);
          }
        }

        // Distinct attendees who have Reported? TRUE
        const reportedAttendees = new Set();
        if (iRep2 >= 0) {
          for (let r = 0; r < out.length; r++) {
            if (truthy_(out[r][iRep2])) {
              const emailVal = iEmail2 >= 0 ? String(out[r][iEmail2] || '').trim().toLowerCase() : '';
              const ptinVal  = iPtin2  >= 0 ? String(out[r][iPtin2]  || '').trim().toUpperCase() : '';
              const fVal     = iFirst2 >= 0 ? String(out[r][iFirst2] || '').trim().toLowerCase() : '';
              const lVal     = iLast2  >= 0 ? String(out[r][iLast2]  || '').trim().toLowerCase() : '';
              const k = emailVal || ptinVal || (fVal + '|' + lVal);
              if (k.replace(/\|/g,'').length) reportedAttendees.add(k);
            }
          }
        }

        // Place the summary well below the data (2 blank rows after the data)
        const summaryRow = targetSheet.getLastRow() + 2;
        // 6 rows: header + 4 metrics + 1 report last updated
        targetSheet.getRange(summaryRow, 1, 6, 2).clearContent();

        // Header: just 'Summary'
        targetSheet.getRange(summaryRow, 1, 1, 2)
          .setValues([['Summary', '']])
          .setFontWeight('bold');

        // Total CE Hours (if present, only where Reported? is TRUE)
        if (iHours2 >= 0 && iRep2 >= 0) {
          const hoursCol = colToA1_(iHours2+1);
          const repCol = colToA1_(iRep2+1);
          const endRow = 1 + out.length;
          targetSheet.getRange(summaryRow + 1, 1).setValue('Total CE Hours:');
          targetSheet.getRange(summaryRow + 1, 2)
            .setFormula(`=SUMIFS(${hoursCol}2:${hoursCol}${endRow}, ${repCol}2:${repCol}${endRow}, TRUE)`);
        }

        // Total Students = distinct attendee count on this sheet
        targetSheet.getRange(summaryRow + 2, 1).setValue('Total Students:');
        targetSheet.getRange(summaryRow + 2, 2).setValue(countDistinctAttendeesOnSheet_(targetSheet, lower, lastCol));

        // Students with Issues = distinct attendee count with nonblank issue
        targetSheet.getRange(summaryRow + 3, 1).setValue('Students with Issues:');
        targetSheet.getRange(summaryRow + 3, 2).setValue(issueAttendees.size);

        // Reported Students = distinct attendees with any row Reported? TRUE
        targetSheet.getRange(summaryRow + 4, 1).setValue('Reported Students:');
        targetSheet.getRange(summaryRow + 4, 2).setValue(reportedAttendees.size);

        // Report Last Updated row (bold)
        targetSheet.getRange(summaryRow + 5, 1, 1, 2)
          .setValues([['Report Last Updated:', syncTimestampStr]])
          .setFontWeight('bold');

        // Keep columns auto-sized so headers are readable
        for (let pass = 0; pass < 2; pass++) {
          for (let c = 1; c <= lastCol; c++) targetSheet.autoResizeColumn(c);
        }
      } catch (e) {
        Logger.log('Summary block failed (non-fatal): ' + e.message);
      }
    }

    totalSheets++;
    totalCells += out.length * lastCol;
    Logger.log(`Group "${gName}" (${gId || 'no-id'}): wrote ${out.length} row(s) to ${targetSS.getName()}.`);
  });

  toast_(`Group sync done: visited ${visited}, wrote ${totalSheets} sheet(s), ~${totalCells} cells.`);
}

/** Helper: place value into the array by header label (lowercased array of headers). */
function placeByHeader_(arr, lower, label, value) {
  const i = lower.indexOf(String(label||'').toLowerCase());
  if (i >= 0) arr[i] = value;
}

/** Open a spreadsheet by URL or Spreadsheet ID */
function openSpreadsheetByUrlOrId_(ref) {
  const s = String(ref||'').trim();
  if (!s) throw new Error('Empty spreadsheet reference.');
  if (/^https?:\/\//i.test(s)) return SpreadsheetApp.openByUrl(s);
  return SpreadsheetApp.openById(s);
}

/** Read "Courses" into Map: Program Number -> Program Name */
function loadCoursesMap_(ss) {
  const sh = ss.getSheetByName('Courses');
  const map = new Map();
  if (!sh || sh.getLastRow() < 2) return map;
  const vals = sh.getDataRange().getValues();
  const hdr  = normalizeHeaderRow_(vals[0]);
  const lower= hdr.map(h=>h.toLowerCase());
  const iNum = lower.indexOf('program number');
  const iNam = lower.indexOf('program name');
  if (iNum < 0 || iNam < 0) return map;
  for (let r=1;r<vals.length;r++){
    const num = String(vals[r][iNum]||'').toUpperCase().replace(/\s+/g,'').trim();
    const nam = String(vals[r][iNam]||'').trim();
    if (num) map.set(num, nam);
  }
  return map;
}

/** Read "Group Config" (preferred) or "Groups" catalog: returns [{groupName, url, groupId}] */
function readGroupsCatalog_(ss) {
  // Preferred: Group Config with headers: Group ID, Group Name, Spreadsheet URL
  let sh = ss.getSheetByName('Group Config');
  if (!sh || sh.getLastRow() < 2) {
    // Fallback legacy: Groups with headers: Group, Sheet URL
    sh = ss.getSheetByName('Groups');
  }
  if (!sh || sh.getLastRow() < 2) return [];

  const vals = sh.getDataRange().getValues();
  const hdr  = normalizeHeaderRow_(vals[0]);
  const lower= hdr.map(h=>h.toLowerCase());

  const idxOfAny = (names) => {
    for (const n of names) {
      const i = lower.indexOf(String(n).toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };

  const iId  = idxOfAny(['group id','id']);
  const iNam = idxOfAny(['group name','group']);
  const iUrl = idxOfAny(['spreadsheet url','sheet url','url']);

  if (iNam < 0 || iUrl < 0) {
    toast_('Group catalog missing "Group Name/Group" and/or "Spreadsheet URL/Sheet URL" headers.', true);
    return [];
  }

  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const groupName = String(vals[r][iNam] || '').trim();
    const url       = String(vals[r][iUrl] || '').trim();
    const groupId   = iId >= 0 ? String(vals[r][iId] || '').trim() : '';
    if (groupName && url) out.push({ groupName, url, groupId });
  }
  return out;
}

/** Convert column index (1-based) to A1 letter */
function colToA1_(c) {
  let s = '';
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

/**
 * Count distinct attendees on the TARGET SHEET by scanning all rows (no caps).
 * Key priority: Email > Attendee PTIN > First+Last.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} lower Lowercased header labels for the sheet's first row
 * @param {number} lastCol Last column index of the sheet
 */
function countDistinctAttendeesOnSheet_(sheet, lower, lastCol) {
  const totalRows = sheet.getLastRow();
  if (totalRows <= 1) return 0; // header only

  // Indices based on header labels
  const iEmail = lower.indexOf('email');
  const iPtin  = lower.indexOf('attendee ptin');
  const iFirst = lower.indexOf('attendee first name');
  const iLast  = lower.indexOf('attendee last name');

  // Read the full body (row 2 .. lastRow) in one pass
  const numRows = totalRows - 1;
  const body = sheet.getRange(2, 1, numRows, lastCol).getValues();

  const seen = new Set();
  for (let r = 0; r < body.length; r++) {
    const row = body[r];
    const emailVal = iEmail >= 0 ? String(row[iEmail] || '').trim().toLowerCase() : '';
    const ptinVal  = iPtin  >= 0 ? String(row[iPtin]  || '').trim().toUpperCase() : '';
    const fVal     = iFirst >= 0 ? String(row[iFirst] || '').trim().toLowerCase() : '';
    const lVal     = iLast  >= 0 ? String(row[iLast]  || '').trim().toLowerCase() : '';
    const key = emailVal || ptinVal || (fVal + '|' + lVal);
    if (key.replace(/\|/g, '').length) seen.add(key);
  }
  return seen.size;
}

/** Export the strict function for menus */
try {
  this.syncGroupSheetsStrict = this.syncGroupSheetsStrict || syncGroupSheetsStrict;
} catch (e) {}