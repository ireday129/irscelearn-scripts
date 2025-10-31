/** mark_reported.js.gs
 * MARK CLEAN AS REPORTED (resumable)
 * - Marks matching Master rows as Reported? TRUE, sets Reported At
 * - Clears Master "Reporting Issue?" and "Last Updated"
 * - Appends a ledger row to "Reported Hours"
 * - Skips any Clean row with a non-blank "Reporting Issue?"
 * - ✅ Sets Roster.Valid? = TRUE for each successfully reported row (by PTIN or Email)
 * - On final batch: clears Clean body, updates Reporting Stats, and syncs RH → Master
 *
 * Requires utils: mustGet_, mapCleanHeaders_, normalizeHeaderRow_, mapHeaders_,
 * normalizeProgram_, formatPtinP0_, hasValue_, parseDate_, formatToMDY_, toast_,
 * mapRosterHeaders_
 * Requires batching helper: runJob(jobKey, stepFn, limit) defined elsewhere.
 */

/** Entry point from menu (fresh run) */
function markCleanAsReported() {
  PropertiesService.getScriptProperties().deleteProperty('JOB/MARK_REPORTED');
  runMarkReportedBatch();
}

/** Kicks the resumable batch */
function runMarkReportedBatch() {
  // runJob must exist in your triggers/util file
  runJob('JOB/MARK_REPORTED', stepMarkReported_, 600);
}

/** Batch step: returns { processed, done } */
function stepMarkReported_(offset, limit) {
  const ss = SpreadsheetApp.getActive();
  const clean = mustGet_(ss, CFG.SHEET_CLEAN);
  const cleanVals = clean.getDataRange().getValues();
  if (cleanVals.length <= 1) return { processed: 0, done: true };

  // Map Clean headers
  const cHdr = cleanVals[0].map(s => String(s || '').trim());
  const cMap = mapCleanHeaders_(cHdr);
  const iCF = cMap.firstName, iCL = cMap.lastName, iCP = cMap.ptin, iCE = cMap.email,
        iCG = cMap.program,   iCH = cMap.hours,    iCC = cMap.completion, iCI = cMap.issue;

  // Read Master
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) return { processed: 0, done: true };

  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mm   = mapHeaders_(mHdr);

  // Required Master columns
  const need = ['email', 'program', 'reportedCol', 'reportedAtCol', 'masterIssueCol'];
  const missing = need.filter(k => mm[k] == null);
  if (missing.length) { toast_('Master missing columns: ' + missing.join(', '), true); return { processed: 0, done: true }; }

  // Optional: Last Updated column (case-insensitive)
  const iLastUpdated = findHeaderIndex_(mHdr, 'Last Updated'); // -1 if absent

  // Build lookups on Master: Program+Email and Program+PTIN
  const mBody = mVals.slice(1);
  const idxPE = new Map(); // prog|email -> rowIndex (0-based within mBody)
  const idxPP = new Map(); // prog|ptin  -> rowIndex
  for (let r = 0; r < mBody.length; r++) {
    const row  = mBody[r];
    const prog = normalizeProgram_(row[mm.program]);
    const email= String(row[mm.email] || '').toLowerCase().trim();
    const ptin = mm.ptin != null ? formatPtinP0_(row[mm.ptin] || '') : '';
    if (prog && email) idxPE.set(prog + '|' + email, r);
    if (prog && ptin)  idxPP.set(prog + '|' + ptin,  r);
  }

  // Batch window
  const start = 1 + (offset || 0);
  const end   = Math.min(start + (limit || 500), cleanVals.length);

  if (start >= end) {
    // Finished: clear Clean body, refresh Reporting Stats, optional RH→Master sync
    if (clean.getLastRow() > 1) clean.getRange(2, 1, clean.getLastRow() - 1, clean.getLastColumn()).clearContent();
    try { if (typeof updateProgramReportedTotals === 'function') updateProgramReportedTotals(); } catch (e) { Logger.log('updateProgramReportedTotals failed: ' + e.message); }
    try { if (typeof syncMasterWithReportedHours === 'function') syncMasterWithReportedHours(true); } catch (e) { Logger.log('syncMasterWithReportedHours failed: ' + e.message); }
    toast_('Mark as Reported complete. Clean cleared.');
    return { processed: 0, done: true };
  }

  let updated = 0;
  const now = new Date();
  const toReportedHours = []; // staged ledger rows

  // ✅ Collect identifiers to flip Roster.Valid? = TRUE after we write Master
  const processedPtins  = new Set();
  const processedEmails = new Set();

  for (let r = start; r < end; r++) {
    const crow = cleanVals[r];

    // Skip rows with a value in "Reporting Issue?"
    if (iCI >= 0 && String(crow[iCI] || '').trim() !== '') continue;

    const prog  = normalizeProgram_(crow[iCG]);
    const email = String(crow[iCE] || '').toLowerCase().trim();
    const ptin  = formatPtinP0_(crow[iCP] || '');

    if (!prog || (!email && !ptin)) continue;

    // Find Master row (prefer Program+Email; fallback Program+PTIN)
    let mi = -1;
    if (email && idxPE.has(prog + '|' + email)) mi = idxPE.get(prog + '|' + email);
    else if (ptin && idxPP.has(prog + '|' + ptin)) mi = idxPP.get(prog + '|' + ptin);
    if (mi < 0) continue;

    const mrow = mBody[mi];

    // Mark as Reported
    mrow[mm.reportedCol] = true;
    if (mm.reportedAtCol != null) mrow[mm.reportedAtCol] = now;

    // Clear issue + last updated
    if (mm.masterIssueCol != null) mrow[mm.masterIssueCol] = '';
    if (iLastUpdated >= 0)        mrow[iLastUpdated]      = '';

    // Non-destructive sync of identity/program fields
    if (mm.firstName != null && iCF >= 0 && hasValue_(crow[iCF])) mrow[mm.firstName] = crow[iCF];
    if (mm.lastName  != null && iCL >= 0 && hasValue_(crow[iCL])) mrow[mm.lastName]  = crow[iCL];
    if (mm.ptin      != null && iCP >= 0 && hasValue_(ptin))       mrow[mm.ptin]      = ptin;
    if (mm.email     != null && iCE >= 0 && hasValue_(email))      mrow[mm.email]     = email;
    if (mm.hours     != null && iCH >= 0 && hasValue_(crow[iCH]))  mrow[mm.hours]     = crow[iCH];
    if (mm.completion!= null && iCC >= 0 && hasValue_(crow[iCC]))  mrow[mm.completion]= parseDate_(crow[iCC]);

    // Stage ledger row for Reported Hours
    toReportedHours.push({
      'Attendee First Name': (iCF>=0 ? crow[iCF] : ''),
      'Attendee Last Name':  (iCL>=0 ? crow[iCL] : ''),
      'Attendee PTIN':       ptin,
      'Program Number':      prog,
      'CE Hours':            (iCH>=0 ? crow[iCH] : ''),
      'Email':               email, // (you said you may remove Email column from RH; this still writes it if present)
      'Program Completion Date': (iCC>=0 ? formatToMDY_(crow[iCC]) : ''),
      'Date Reported':       formatToMDY_(now)
    });

    // ✅ Track for Roster flip
    if (ptin)  processedPtins.add(ptin);
    if (email) processedEmails.add(email);

    updated++;
  }

  if (updated) {
    master.getRange(2, 1, mBody.length, mHdr.length).setValues(mBody);
  }

  if (toReportedHours.length) appendToReportedHours_(ss, toReportedHours);

  // ✅ Flip Roster.Valid? = TRUE for all successfully reported PTIN/Email seen in this batch
  try { setRosterValidForKeys_(processedPtins, processedEmails); } catch (e) { Logger.log('setRosterValidForKeys_ failed: ' + e.message); }

  return { processed: (end - start), done: false };
}

/** Append staged rows to the "Reported Hours" sheet */
function appendToReportedHours_(ss, rows) {
  if (!rows || !rows.length) return;
  const sheetName = 'Reported Hours';
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { toast_(`${sheetName} sheet not found; cannot append reported rows.`, true); return; }

  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const startRow = sh.getLastRow() + 1;
  const numCols  = hdr.length;

  const data = rows.map(r => {
    const arr = new Array(numCols).fill('');
    hdr.forEach((header, index) => {
      const key = String(header || '').replace(/\?$/, '').trim();
      switch (key) {
        case 'Attendee First Name':       arr[index] = r['Attendee First Name']; break;
        case 'Attendee Last Name':        arr[index] = r['Attendee Last Name'];  break;
        case 'Attendee PTIN':             arr[index] = r['Attendee PTIN']; break;
        case 'Program Number':            arr[index] = r['Program Number'];      break;
        case 'CE Hours':                  arr[index] = r['CE Hours'];            break;
        case 'Email':                     arr[index] = r['Email'];               break; // harmless if column not present
        case 'Program Completion Date':   arr[index] = r['Program Completion Date']; break;
        case 'Date Reported':             arr[index] = r['Date Reported'];       break;
      }
    });
    return arr;
  });

  sh.getRange(startRow, 1, data.length, numCols).setValues(data);

  const iReportedDate = hdr.map(h => String(h||'').trim()).indexOf('Date Reported');
  if (iReportedDate >= 0) {
    sh.getRange(startRow, iReportedDate + 1, data.length, 1).setNumberFormat('mm/dd/yyyy');
  }
}

/** Case-insensitive header finder */
function findHeaderIndex_(hdrArray, label) {
  const lower = hdrArray.map(h => String(h||'').trim().toLowerCase());
  return lower.indexOf(String(label||'').trim().toLowerCase());
}

/** ✅ Helper: Set Roster.Valid? = TRUE for any matching PTIN or Email */
function setRosterValidForKeys_(ptinSet, emailSet) {
  if ((!ptinSet || ptinSet.size === 0) && (!emailSet || emailSet.size === 0)) return;

  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(CFG.SHEET_ROSTER);
  if (!roster) { toast_(`Sheet "${CFG.SHEET_ROSTER}" not found.`, true); return; }

  const rMap = mapRosterHeaders_(roster);
  if (!rMap) { toast_('Roster headers missing or renamed.', true); return; }
  if (rMap.valid == null || rMap.valid < 0) { toast_('Roster has no "Valid?" column.', true); return; }

  const lastRow = roster.getLastRow();
  if (lastRow < 2) return;

  const rng = roster.getRange(2, 1, lastRow - 1, roster.getLastColumn());
  const vals = rng.getValues();

  let changed = 0;
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const email = String(row[rMap.email] || '').toLowerCase().trim();
    const ptin  = formatPtinP0_(row[rMap.ptin] || '');
    const hit = (email && emailSet && emailSet.has(email)) || (ptin && ptinSet && ptinSet.has(ptin));
    if (hit && !parseBool_(row[rMap.valid])) {
      row[rMap.valid] = true;
      changed++;
    }
  }

  if (changed) rng.setValues(vals);
  if (changed) toast_(`Roster Valid? updated: ${changed} row(s) set TRUE after reporting.`);
}