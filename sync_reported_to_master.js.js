/**
 * Ensure every row in "Reported Hours" is reflected in "Master".
 * Key = PTIN + Program Number (email no longer used in RH).
 * - Updates existing rows (marks Reported?, sets Reported At, syncs hours/date)
 * - Appends new rows if missing (without touching headers)
 * - Clears "Reporting Issue?" and "Last Updated" on affected Master rows
 */
function syncMasterWithReportedHours(quiet) {
  const ss = SpreadsheetApp.getActive();
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const rh     = mustGet_(ss, 'Reported Hours');

  // Optional: sweep duplicates in RH first (if you added this helper)
  try { if (typeof enforceReportedHoursUniqueness_ === 'function') enforceReportedHoursUniqueness_(); } catch(e){ Logger.log(e); }

  const rhVals = rh.getDataRange().getValues();
  if (rhVals.length <= 1) { if(!quiet) toast_('Reported Hours is empty; nothing to sync.'); return; }

  // Reported Hours header map (email column may not exist)
  const rhHdr   = rhVals[0].map(s=>String(s||'').trim());
  const rhLower = rhHdr.map(h=>h.toLowerCase());
  const rhIdx   = (label) => rhLower.indexOf(String(label||'').toLowerCase().trim());

  const iRhF   = rhIdx('attendee first name');
  const iRhL   = rhIdx('attendee last name');
  const iRhPT  = rhIdx('ptin');              // support either header
  const iRhPT2 = rhIdx('attendee ptin');     // support either header
  const iRhP   = rhIdx('program number');
  const iRhH   = rhIdx('ce hours');
  const iRhC   = rhIdx('program completion date');
  const iRhDR  = rhIdx('date reported');

  const iRhPTIN = iRhPT >= 0 ? iRhPT : iRhPT2;
  if (iRhPTIN < 0 || iRhP < 0) {
    if(!quiet) toast_('Reported Hours missing PTIN and/or Program Number columns.', true);
    return;
  }

  // Master header map
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) { if(!quiet) toast_('Master is empty; cannot sync.', true); return; }
  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mm   = mapHeaders_(mHdr);

  // Optional columns on Master
  const mLower = mHdr.map(h=>h.toLowerCase());
  const iM_LastUpdated = mLower.indexOf('last updated'); // may be -1

  // Build index on Master by Program+PTIN
  const mBody   = mVals.slice(1);
  const idxPP   = new Map(); // prog|ptin -> rowIndex
  for (let r=0; r<mBody.length; r++){
    const row  = mBody[r];
    const prog = normalizeProgram_(row[mm.program] || '');
    const ptin = mm.ptin!=null ? formatPtinP0_(row[mm.ptin]||'') : '';
    if (prog && ptin) idxPP.set(prog + '|' + ptin, r);
  }

  const toAppend = [];
  let updates = 0, appends = 0;

  // Walk Reported Hours rows and upsert into Master (PTIN+Program)
  for (let r=1; r<rhVals.length; r++){
    const rr   = rhVals[r];
    const ptin = formatPtinP0_(rr[iRhPTIN] || '');
    const prog = normalizeProgram_(rr[iRhP]   || '');
    if (!prog || !ptin) continue;

    const first = iRhF >= 0 ? String(rr[iRhF]||'').trim() : '';
    const last  = iRhL >= 0 ? String(rr[iRhL]||'').trim() : '';
    const hours = iRhH >= 0 ? rr[iRhH] : '';
    const comp  = iRhC >= 0 ? rr[iRhC] : '';
    const dRep  = iRhDR>= 0 ? rr[iRhDR] : new Date();

    const key = prog + '|' + ptin;
    const mi  = idxPP.get(key);

    if (mi !== undefined) {
      // UPDATE existing master row
      const mrow = mBody[mi];
      if (mm.reportedCol    != null) mrow[mm.reportedCol]    = true;
      if (mm.reportedAtCol  != null) mrow[mm.reportedAtCol]  = (dRep instanceof Date ? dRep : parseDate_(dRep) || new Date());
      if (mm.masterIssueCol != null) mrow[mm.masterIssueCol] = '';  // Clear any issue
      if (iM_LastUpdated    >= 0)    mrow[iM_LastUpdated]    = '';  // Clear Last Updated

      // Fill blanks for identity fields
      if (mm.firstName != null && isBlankCell_(mrow[mm.firstName]) && first) mrow[mm.firstName] = first;
      if (mm.lastName  != null && isBlankCell_(mrow[mm.lastName])  && last)  mrow[mm.lastName]  = last;
      if (mm.ptin      != null && isBlankCell_(mrow[mm.ptin])      && ptin)  mrow[mm.ptin]      = ptin;

      // Always update program data (safe): hours + completion
      if (mm.hours     != null && hasValue_(hours)) mrow[mm.hours]      = hours;
      if (mm.completion!= null && hasValue_(comp))  mrow[mm.completion] = parseDate_(comp) || comp;

      updates++;
    } else {
      // APPEND new master row
      const newRow = new Array(mHdr.length).fill('');
      if (mm.firstName    != null) newRow[mm.firstName]    = first;
      if (mm.lastName     != null) newRow[mm.lastName]     = last;
      if (mm.ptin         != null) newRow[mm.ptin]         = ptin;
      /* no email to set from RH */
      if (mm.program      != null) newRow[mm.program]      = prog;
      if (mm.hours        != null) newRow[mm.hours]        = hours;
      if (mm.completion   != null) newRow[mm.completion]   = parseDate_(comp) || comp;
      if (mm.reportedCol  != null) newRow[mm.reportedCol]  = true;
      if (mm.reportedAtCol!= null) newRow[mm.reportedAtCol]= (dRep instanceof Date ? dRep : parseDate_(dRep) || new Date());
      if (mm.masterIssueCol!=null) newRow[mm.masterIssueCol] = '';
      if (iM_LastUpdated  >= 0)    newRow[iM_LastUpdated]  = '';

      toAppend.push(newRow);
      appends++;
    }
  }

  // Write back updates
  if (updates) {
    master.getRange(2, 1, mBody.length, mHdr.length).setValues(mBody);
  }
  // Append new rows (no header changes)
  if (toAppend.length) {
    master.getRange(master.getLastRow()+1, 1, toAppend.length, mHdr.length).setValues(toAppend);
  }

  if (!quiet) toast_(`Reported Hours → Master sync: ${updates} updated, ${appends} appended.`);
}