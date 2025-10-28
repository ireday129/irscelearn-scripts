/**
 * Public function to run Build Clean Upload via the custom menu.
 * (NON-RESUMABLE version to avoid ReferenceError: runJob)
 * * This function orchestrates the preparation of the Clean sheet by:
 * 1. Running necessary pre-cleanups (dedupe Master/Roster, backfill PTIN).
 * 2. Processing Master rows, filtering out known unresolved issues and reported rows.
 * 3. Consolidating reported/clean rows and writing them to the Clean sheet.
 */
/**
 * Stub for missing dependency in BuildClean and RecheckMaster
 * (Row-level roster fixes can be added here later)
 */
function applyReportingFixes(quiet) {
  // For now: do nothing, but prevent errors
  if (!quiet) {
    console.log('applyReportingFixes called but no fix logic implemented yet.');
  }
  return;
}
function buildCleanUpload() {
  const ss = SpreadsheetApp.getActive();

  // --- 0. Pre-Cleanup (mimics start of batch job) ---
  try {
    // Functions assumed to be globally defined or in other linked files:
    // ingestSystemReportingIssues(true); 
    // syncMasterFromIssueSheet_();
  } catch (e) {
    Logger.log('Initial ingest failed: ' + (e.stack || e.message));
  }
  // Functions assumed to be globally defined or in other linked files:
  dedupeMasterByEmailProgram(true);
  applyReportingFixes(true);
  dedupeRosterByEmail(true);
  backfillMasterPtinFromRoster_(true);

  // --- 1. Setup ---
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) {
    // These helpers must be available globally
    writeCleanDataOnly_(ss, []);
    toast_('Master sheet is empty. Clean upload generated with 0 rows.', false);
    return;
  }

  // Clear Clean body
  writeCleanDataOnly_(ss, []);

  const header = normalizeHeaderRow_(mVals[0]);
  const mMap = mapHeaders_(header);

  const need = ['firstName','lastName','ptin','email','program','hours','completion','group','masterIssueCol','reportedCol'];
  const missing = need.filter(k => mMap[k] == null);
  if (missing.length) {
    toast_('Master is missing columns for Clean Build: ' + missing.join(', '), true);
    return;
  }

  // --- 2. Core Processing ---
  const unresolved = buildUnresolvedIssueIndex_();
  const roster = getRosterMap_(ss);
  const ptinRe = /^P0\d{7}$/i;

  const cleanCandidates = new Map(); // Key: PTIN|Program -> record (Last row wins)

  for (let r = 1; r < mVals.length; r++) {
    const row = mVals[r];
    const rec = {
      firstName: String(row[mMap.firstName]||'').trim(),
      lastName:  String(row[mMap.lastName]||'').trim(),
      ptin:      formatPtinP0_(row[mMap.ptin]||''),
      email:     String(row[mMap.email]||'').toLowerCase().trim(),
      program:   normalizeProgram_(row[mMap.program]),
      hours:     row[mMap.hours],
      completion: row[mMap.completion],
      group:     String(row[mMap.group]||'').trim(),
      issueVal:  String(row[mMap.masterIssueCol]||'').trim(),
      reported:  parseBool_(row[mMap.reportedCol])
    };

    // 1. HARD EXCLUSION: If the row has already been reported, skip it entirely.
    if (rec.reported === true) {
        continue; 
    }

    // 2. HARD EXCLUSION: skip if this record matches any UNRESOLVED issue
    const keyPP = rec.ptin ? (rec.ptin + '|' + rec.program) : '';
    const keyEP = rec.email ? (rec.email + '|' + rec.program) : '';
    const keyNP = (rec.firstName && rec.lastName)
      ? (rec.firstName.trim().toLowerCase() + ' ' + rec.lastName.trim().toLowerCase() + '|' + rec.program)
      : '';

    if (rec.program && (
        (keyPP && unresolved.byPtinProg.has(keyPP)) ||
        (keyEP && unresolved.byEmailProg.has(keyEP)) ||
        (keyNP && unresolved.byNameProg.has(keyNP))
    )) {
      continue;
    }
    
    // 3. NEW HARD EXCLUSION: If the Reporting Issue column has ANY value, skip it.
    // This implements the strict rule to only process rows that are completely blank in that column.
    if (rec.issueVal !== '') {
        continue;
    }

    // Status classification for any *new* problems
    let status = 'Good';
    if (!rec.ptin) status = 'Missing PTIN';
    else if (!ptinRe.test(rec.ptin)) status = 'PTIN does not exist';
    const ro = roster ? roster.get(rec.ptin) : null;
    if (status === 'Good' && ro && !namesMatchFull_(rec.firstName, rec.lastName, ro.first, ro.last)) {
      status = 'PTIN & name do not match';
    }
    // NOTE: The previous line 'if (/^fixed$/i.test(rec.issueVal)) status = 'Good';' 
    // has been removed because the row is already skipped if rec.issueVal is not empty.

    // --- 3. Partition Logic ---
    if (status === 'Good') {
      // Send good rows to Clean 
      const cleanKey = rec.ptin + '|' + rec.program;
      if (rec.ptin && rec.program) {
        cleanCandidates.set(cleanKey, rec);
      }
    } else {
      // Logic for handling new issues is simplified: issues are expected to be flagged
      // on the Master sheet directly by recheckMaster.
    }
  }

  // --- 4. Final Write ---
  const cleanRows = Array.from(cleanCandidates.values()).map(r => ({
    first: r.firstName,
    last:  r.lastName,
    ptin:  r.ptin,
    email: r.email,
    prog:  r.program,
    hours: r.hours,
    completion: normalizeCompletionForUpload_(r.completion)
  }));

  appendToClean_(ss, cleanRows);

  // --- 5. Finalize & Cleanup ---
  try {
    // ingestSystemReportingIssues(true);
    // syncMasterFromIssueSheet_();
  } catch (e) {
    Logger.log('Finalize ingest failed: ' + (e.stack || e.message));
  }
  // Functions assumed to be globally defined or in other linked files:
  applyReportingIssueValidationAndFormatting_();
  updateRosterValidityFromIssues_();
  // syncGroupSheets(true); 
  
  toast_(`Clean Upload Complete: ${cleanRows.length} rows generated.`, false);
}
/** Sweep Master: normalize PTINs (PO->P0) and flag P00000000 as "PTIN does not exist". */
function sweepPtinNormalizationAndIssues(quiet){
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(CFG.SHEET_MASTER);
  if (!master) { if(!quiet) toast_(`Sheet "${CFG.SHEET_MASTER}" not found.`, true); return; }

  const vals = master.getDataRange().getValues();
  if (vals.length <= 1) return;

  const hdr = normalizeHeaderRow_(vals[0]);
  const mm  = mapHeaders_(hdr);
  if (mm.ptin == null || mm.masterIssueCol == null) {
    if (!quiet) toast_('Master is missing PTIN and/or Reporting Issue? columns.', true);
    return;
  }

  const body = vals.slice(1);
  let changed = 0;

  for (let i=0; i<body.length; i++){
    const row = body[i];

    // 1) Normalize PTIN: fix PO→P0 first, then format to P0####### if possible
    let raw = String(row[mm.ptin] || '').trim();
    if (raw) {
      // fix common typo "PO" (letter O) at start
      if (/^pO/i.test(raw)) raw = 'P0' + raw.slice(2);
      const normalized = formatPtinP0_(raw);
      if (normalized !== row[mm.ptin]) {
        row[mm.ptin] = normalized;
        changed++;
      }

      // 2) Flag obvious invalid P00000000 (unless explicitly marked Fixed)
      if (normalized === 'P00000000') {
        const curIssue = String(row[mm.masterIssueCol] || '').trim();
        if (!/^fixed$/i.test(curIssue) && curIssue !== 'PTIN does not exist') {
          row[mm.masterIssueCol] = 'PTIN does not exist';
          changed++;
        }
      }
    }
  }

  if (changed) {
    master.getRange(2, 1, body.length, hdr.length).setValues(body);
    if (!quiet) toast_(`PTIN sweep complete: ${changed} cell change${changed===1?'':'s'}.`);
  } else if (!quiet) {
    toast_('PTIN sweep: no changes needed.');
  }
}