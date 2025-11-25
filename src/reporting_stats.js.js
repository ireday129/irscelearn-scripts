/** reporting_stats.js
 * Updates per-program totals on "Reporting Stats" using the "Reported Hours" sheet.
 * Counts rows in REPORTED HOURS grouped by Program Number.
 * Needs utils: mustGet_, normalizeProgram_.
 */

function updateProgramReportedTotals() {
  const ss = SpreadsheetApp.getActive();
  const reported = mustGet_(ss, 'Reported Hours');
  const stats    = mustGet_(ss, 'Reporting Stats');

  // Program Number -> target cell on Reporting Stats (B11..B18 as specified)
  const PROGRAM_COUNT_CELLS = {
    'RTPMH-A-00004-25-S': 'B11', // Annual Federal Tax Refresher
    'RTPMH-T-00010-25-S': 'B12', // Tax Cuts & Jobs Act Walkthrough
    'RTPMH-T-00009-25-S': 'B13', // 1040 Schedule C: Business or Hobby
    'RTPMH-T-00008-25-S': 'B14', // 1040 Schedule A: Itemized Deductions
    'RTPMH-T-00007-25-S': 'B15', // Earned Income Tax Credit: Who and Why
    'RTPMH-T-00006-25-S': 'B16', // Child & Dependent Care Credit Decoded
    'RTPMH-T-00003-25-S': 'B17', // Mastering IRS Authorizations: 8821, 2848 & the CAF Unit
    'RTPMH-E-00005-25-S': 'B18'  // Circular 230: Tax Pro Bible
  };

  // Build normalized keys and zeroed counters
  const keys = Object.keys(PROGRAM_COUNT_CELLS);
  const wanted = new Set(keys.map(normalizeProgram_));
  const counts = {};
  keys.forEach(k => counts[normalizeProgram_(k)] = 0);

  const vals = reported.getDataRange().getValues();
  if (vals.length <= 1) {
    // nothing logged yet; write zeros
    for (const addr of Object.values(PROGRAM_COUNT_CELLS)) stats.getRange(addr).setValue(0);
    return;
  }

  // Map headers (case-insensitive)
  const hdr = vals[0].map(s => String(s||'').trim().toLowerCase());
  const iProgram = hdr.indexOf('program number');
  if (iProgram < 0) throw new Error('Reported Hours is missing "Program Number" column.');

  // Tally by normalized Program Number
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const prog = normalizeProgram_(row[iProgram]);
    if (!prog) continue;
    if (wanted.has(prog)) counts[prog] = (counts[prog] || 0) + 1;
  }

  // Write counts to target cells (leave formatting alone)
  for (const [progRaw, addr] of Object.entries(PROGRAM_COUNT_CELLS)) {
    const key = normalizeProgram_(progRaw);
    stats.getRange(addr).setValue(counts[key] || 0);
  }
}

/** Optional: call from your nightly job */
function nightlyUpdateReportingStats() {
  updateTotalReportedCEHours();
  // toast_('Reporting Stats updated from Reported Hours.'); // enable if you like toasts
}

/**
 * Update Reporting Stats — CE Hours total and per-program counts from Reported Hours
 * Places per-program counts in B11..B18, and total CE hours in B5.
 */
function updateTotalReportedCEHours() {
  const ss = SpreadsheetApp.getActive();
  const reported = mustGet_(ss, 'Reported Hours');
  const stats    = mustGet_(ss, 'Reporting Stats');

  // Program Number -> target cell on Reporting Stats (B11..B18 as specified)
  const PROGRAM_COUNT_CELLS = {
    'RTPMH-A-00004-25-S': 'B11', // Annual Federal Tax Refresher
    'RTPMH-T-00010-25-S': 'B12', // Tax Cuts & Jobs Act Walkthrough
    'RTPMH-T-00009-25-S': 'B13', // 1040 Schedule C: Business or Hobby
    'RTPMH-T-00008-25-S': 'B14', // 1040 Schedule A: Itemized Deductions
    'RTPMH-T-00007-25-S': 'B15', // Earned Income Tax Credit: Who and Why
    'RTPMH-T-00006-25-S': 'B16', // Child & Dependent Care Credit Decoded
    'RTPMH-T-00003-25-S': 'B17', // Mastering IRS Authorizations: 8821, 2848 & the CAF Unit
    'RTPMH-E-00005-25-S': 'B18'  // Circular 230: Tax Pro Bible
  };

  const keys   = Object.keys(PROGRAM_COUNT_CELLS);
  const wanted = new Set(keys.map(normalizeProgram_));
  const counts = {};
  keys.forEach(k => counts[normalizeProgram_(k)] = 0);

  const vals = reported.getDataRange().getValues();

  // If no data rows, zero out everything and exit
  if (vals.length <= 1) {
    for (const addr of Object.values(PROGRAM_COUNT_CELLS)) {
      stats.getRange(addr).setValue(0);
    }
    stats.getRange('B5').setValue(0);
    return;
  }

  // Map headers (case-insensitive)
  const hdr = vals[0].map(s => String(s || '').trim().toLowerCase());
  const iProgram = hdr.indexOf('program number');
  if (iProgram < 0) throw new Error('Reported Hours is missing "Program Number" column.');

  // Try to locate CE Hours column by header with several candidates; fall back to column E (index 4)
  const CE_HOURS_HEADER_CANDIDATES = [
    'ce hours reported',
    'ce hours',
    'hours reported',
    'hours'
  ];

  let iHours = -1;
  for (const cand of CE_HOURS_HEADER_CANDIDATES) {
    const idx = hdr.indexOf(cand);
    if (idx >= 0) {
      iHours = idx;
      break;
    }
  }

  if (iHours < 0) {
    // Fallback: assume column E (index 4) if we still couldn't find a match
    iHours = 4;
  }

  let totalHours = 0;

  // Walk data rows once: tally per-program counts AND total CE hours
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];

    // Program count logic
    const prog = normalizeProgram_(row[iProgram]);
    if (prog && wanted.has(prog)) {
      counts[prog] = (counts[prog] || 0) + 1;
    }

    // CE hours total
    const v = row[iHours];
    const num = Number(v);
    if (!isNaN(num)) totalHours += num;
  }

  // Write per-program counts
  for (const [progRaw, addr] of Object.entries(PROGRAM_COUNT_CELLS)) {
    const key = normalizeProgram_(progRaw);
    stats.getRange(addr).setValue(counts[key] || 0);
  }

  // Write total CE hours to B5
  stats.getRange('B5').setValue(totalHours);
}