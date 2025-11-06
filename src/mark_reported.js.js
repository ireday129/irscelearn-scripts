function markCleanAsReported() {
  const ss = SpreadsheetApp.getActive();
  const clean = mustGet_(ss, CFG.SHEET_CLEAN);
  const master = mustGet_(ss, CFG.SHEET_MASTER);

  const cVals = clean.getDataRange().getValues();
  if (cVals.length <= 1) {
    toast_('No rows on Clean to mark as reported.', true);
    return;
  }

  const cHdr = cVals[0].map(s => String(s || '').trim());
  const cMap = mapCleanHeaders_(cHdr);
  const cBody = cVals.slice(1);

  const ci = {
    first:     cMap.firstName,
    last:      cMap.lastName,
    ptin:      cMap.ptin,
    email:     cMap.email,
    program:   cMap.program,
    hours:     cMap.hours,
    completion:cMap.completion
  };

  // Require at least (Email + Program) OR (PTIN + Program)
  const emailModeOk = ci.email >= 0 && ci.program >= 0;
  const ptinModeOk  = ci.ptin  >= 0 && ci.program >= 0;
  if (!emailModeOk && !ptinModeOk) {
    toast_('Clean sheet missing Email/PTIN + Program columns; cannot mark reported.', true);
    return;
  }

  // --- Master ---
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) {
    toast_('Master is empty; nothing to update.', true);
    return;
  }

  const mHdr  = normalizeHeaderRow_(mVals[0]);
  const mm    = mapHeaders_(mHdr);
  const mBody = mVals.slice(1);

  const idxPtin       = mm.ptin;
  const idxEmail      = mm.email;
  const idxProg       = mm.program;
  const idxReported   = mm.reportedCol;
  const idxReportedAt = mm.reportedAtCol;

  if (idxProg == null || idxReported == null || idxReportedAt == null) {
    toast_('Master missing Program/Reported?/Reported At columns; cannot mark reported.', true);
    return;
  }

  // Build lookup maps for Master: Email+Program and PTIN+Program
  const byEmailProg = new Map();
  const byPtinProg  = new Map();

  const normProg = v => normalizeProgram_(v || '');

  for (let r = 0; r < mBody.length; r++) {
    const row  = mBody[r];
    const prog = normProg(row[idxProg]);
    if (!prog) continue;

    if (idxEmail != null) {
      const em = String(row[idxEmail] || '').toLowerCase().trim();
      if (em) byEmailProg.set(em + '|' + prog, r);
    }
    if (idxPtin != null) {
      const pt = formatPtinP0_(row[idxPtin] || '');
      if (pt) byPtinProg.set(pt + '|' + prog, r);
    }
  }

  const now = new Date();
  let updated = 0;

  // rows we’ll append to Reported Hours
  const reportedRows = [];

  // --- Walk Clean rows and mark matching Master rows as reported ---
  for (let i = 0; i < cBody.length; i++) {
    const row = cBody[i];
    const prog = normProg(row[ci.program]);
    if (!prog) continue;

    const em = ci.email >= 0 ? String(row[ci.email] || '').toLowerCase().trim() : '';
    const pt = ci.ptin  >= 0 ? formatPtinP0_(row[ci.ptin] || '') : '';

    let masterIndex = null;
    if (em) {
      const k = em + '|' + prog;
      if (byEmailProg.has(k)) masterIndex = byEmailProg.get(k);
    }
    if (masterIndex == null && pt) {
      const k = pt + '|' + prog;
      if (byPtinProg.has(k)) masterIndex = byPtinProg.get(k);
    }
    if (masterIndex == null) continue; // no match in Master

    const mrow = mBody[masterIndex];

    // Flip Master flags
    mrow[idxReported]   = true;
    mrow[idxReportedAt] = now;
    updated++;

    // --- Build a robust row for Reported Hours ---
    // Prefer Clean values, but backfill from Master when missing.
    const firstFromClean = ci.first >= 0 ? row[ci.first] : '';
    const lastFromClean  = ci.last  >= 0 ? row[ci.last]  : '';
    const ptFromClean    = ci.ptin  >= 0 ? formatPtinP0_(row[ci.ptin] || '') : '';

    const firstFromMaster = mm.firstName != null ? mrow[mm.firstName] : '';
    const lastFromMaster  = mm.lastName  != null ? mrow[mm.lastName]  : '';
    const ptFromMaster    = idxPtin      != null ? formatPtinP0_(mrow[idxPtin] || '') : '';

    const firstVal = firstFromClean || firstFromMaster || '';
    const lastVal  = lastFromClean  || lastFromMaster  || '';
    const ptinVal  = ptFromClean    || ptFromMaster    || '';

    const progVal = ci.program   >= 0 ? row[ci.program]   : (idxProg != null       ? mrow[idxProg]       : '');
    const hoursVal= ci.hours     >= 0 ? row[ci.hours]     : (mm.hours != null      ? mrow[mm.hours]      : '');
    const compVal = ci.completion>= 0 ? row[ci.completion]: (mm.completion != null ? mrow[mm.completion] : '');

    reportedRows.push({
      first: firstVal,
      last: lastVal,
      ptin: ptinVal,
      prog: progVal,
      hours: hoursVal,
      completion: compVal,
      reportedAt: now
    });
  }

  // Write Master back if anything changed
  if (updated) {
    master.getRange(2, 1, mBody.length, mHdr.length).setValues(mBody);
  }

  // --- Append to Reported Hours sheet (from Clean→Master matches) ---
  try {
    if (reportedRows.length) {
      const reportedName =
        (typeof CFG !== 'undefined' && CFG.SHEET_REPORTED_HOURS)
          ? CFG.SHEET_REPORTED_HOURS
          : 'Reported Hours';

      const rh = ss.getSheetByName(reportedName);
      if (!rh) {
        toast_('Reported Hours sheet not found; skipping append.', true);
      } else {
        const hdr = rh.getRange(1, 1, 1, rh.getLastColumn())
          .getValues()[0]
          .map(x => String(x || '').trim());
        const lower = hdr.map(h => h.toLowerCase());

        // Re-use clean header mapper (it understands our common labels)
        const hMap = mapCleanHeaders_(hdr);
        const iF   = hMap.firstName;
        const iL   = hMap.lastName;
        const iP   = hMap.ptin;
        const iG   = hMap.program;
        const iH   = hMap.hours;
        const iC   = hMap.completion;

        const iDateReported = lower.indexOf('date reported');

        if ([iF, iL, iP, iG, iH, iC].some(v => v < 0) || iDateReported < 0) {
          toast_(
            'Reported Hours sheet missing one of: Attendee First Name, Attendee Last Name, ' +
            'Attendee PTIN, Program Number, CE Hours, Program Completion Date, Date Reported.',
            true
          );
        } else {
          const startRow = rh.getLastRow() + 1;
          const data = reportedRows.map(r => {
            const arr = new Array(hdr.length).fill('');

            arr[iF] = r.first || '';
            arr[iL] = r.last  || '';
            arr[iP] = formatPtinP0_(r.ptin || '');
            arr[iG] = normalizeProgram_(r.prog || '');
            arr[iH] = r.hours || '';
            arr[iC] = r.completion || '';
            arr[iDateReported] = r.reportedAt;

            return arr;
          });

          rh.getRange(startRow, 1, data.length, hdr.length).setValues(data);

          // Make Date Reported look like a proper timestamp
          rh.getRange(startRow, iDateReported + 1, data.length, 1)
            .setNumberFormat('mm/dd/yyyy h:mm am/pm');
        }
      }
    }
  } catch (e) {
    Logger.log('Appending to Reported Hours failed (non-fatal): ' + e.message);
  }

  // Clear Clean sheet body (keep headers)
  if (clean.getLastRow() > 1) {
    clean
      .getRange(2, 1, clean.getLastRow() - 1, clean.getLastColumn())
      .clearContent();
  }

  // Update program totals & sync RH→Master if those helpers exist
  try {
    if (typeof updateProgramReportedTotals === 'function') {
      updateProgramReportedTotals();
    }
  } catch (e) {
    Logger.log('updateProgramReportedTotals failed: ' + e.message);
  }

  try {
    if (typeof syncMasterWithReportedHours === 'function') {
      syncMasterWithReportedHours(true);
    }
  } catch (e) {
    Logger.log('syncMasterWithReportedHours failed: ' + e.message);
  }

  // Stamp Reporting Stats!B6 with the exact finish time in explicit EST
  try {
    const statsName =
      (typeof CFG !== 'undefined' && CFG.SHEET_REPORTING_STATS)
        ? CFG.SHEET_REPORTING_STATS
        : 'Reporting Stats';
    const statsSh = mustGet_(ss, statsName);

    const stamp = Utilities.formatDate(
      now,
      'America/New_York',
      "MMM d, yyyy h:mm a 'EST'"
    );
    statsSh.getRange('B6').setValue(stamp);
  } catch (e) {
    Logger.log('Failed to write Reporting Stats!B6 timestamp: ' + e.message);
  }

  // Persist the exact finish time for downstream summaries (e.g., group sync)
  try {
    PropertiesService.getScriptProperties()
      .setProperty('LAST_CE_HOURS_REPORTED_AT', now.toISOString());
  } catch (e) {
    Logger.log('Failed to set LAST_CE_HOURS_REPORTED_AT: ' + e.message);
  }

  // Refresh roster highlighting based on reported hours, if available
  try {
    if (typeof highlightRosterFromReportedHours === 'function') {
      highlightRosterFromReportedHours();
    }
  } catch (e) {
    Logger.log('highlightRosterFromReportedHours failed (non-fatal): ' + e.message);
  }

  toast_('Mark as Reported complete. Clean cleared. Updated ' + updated + ' Master row(s).');
}

function markCleanAsReportedMenu() {
  return markCleanAsReported();
}

// Defensive global exports for Apps Script
try {
  this.markCleanAsReported = this.markCleanAsReported || markCleanAsReported;
  this.markCleanAsReportedMenu = this.markCleanAsReportedMenu || markCleanAsReportedMenu;
} catch (e) {
  // no-op in environments where `this` is not the global
}