function recheckMaster() {
  try {
    clearMasterIssuesFromFixedIssues_(true); // safe stub

    backfillMasterPtinFromRoster_(true);

    const ss = SpreadsheetApp.getActive();
    const master = mustGet_(ss, CFG.SHEET_MASTER);
    const mVals = master.getDataRange().getValues();
    if (mVals.length<=1) { toast_('Master is empty.', true); return; }

    const mHdr = normalizeHeaderRow_(mVals[0]);
    const mMap = mapHeaders_(mHdr);
    const roster = getRosterMap_(ss);

    // Valid *format* only — does NOT bless P00000000
    const ptinRe = /^P0\d{7}$/i;

    const out = [];
    for (let i=1;i<mVals.length;i++){
      const row   = mVals[i];
      const first = String(row[mMap.firstName]||'').trim();
      const last  = String(row[mMap.lastName]||'').trim();
      const ptin  = formatPtinP0_(row[mMap.ptin]||'');
      const mIssue= String(row[mMap.masterIssueCol]||'').trim();

      // If the row is already Reported? -> do not compute issues
      if (parseBool_(row[mMap.reportedCol])) {
        out.push(['']); // leave blank
        continue;
      }

      // Preserve existing "PTIN does not exist" exactly as requested
      if (mIssue === 'PTIN does not exist') {
        out.push(['PTIN does not exist']);
        continue;
      }

      // Compute fresh status
      let status = 'Good';

      if (!ptin) {
        status = 'Missing PTIN';
      } else if (ptin.toUpperCase() === 'P00000000') {
        // Explicitly treat all-zero PTIN as invalid
        status = 'PTIN does not exist';
      } else if (!ptinRe.test(ptin)) {
        status = 'PTIN does not exist';
      } else if (roster) {
        const ro = roster.get(ptin);
        if (ro && !namesMatchFull_(first, last, ro.first, ro.last)) {
          status = 'PTIN & name do not match';
        }
      }

      // Respect "Fixed" marker
      if (/^fixed$/i.test(mIssue)) status = 'Good';

      out.push([status==='Good' ? '' : mapFreeTextToStandardIssue_(status)]);
    }

    if (mMap.masterIssueCol!=null) {
      master.getRange(2, mMap.masterIssueCol+1, out.length, 1).setValues(out);
    }

    toast_('Master rechecked (sticky "PTIN does not exist" preserved).', false);
  } catch(e) {
    toast_('Recheck failed: ' + e.message, true);
    Logger.log(e.stack||e.message);
  }
}