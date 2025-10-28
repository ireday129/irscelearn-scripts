/**
 * De-duplicate "Reported Hours" by (PTIN + Program Number).
 * Winner = most recent by Date Reported, else Program Completion Date, else row order (later wins).
 * Leaves headers untouched and rewrites the body with winners only.
 */
function dedupeReportedHoursMostRecentWinner(quiet) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Reported Hours');
  if (!sh || sh.getLastRow() <= 1) return;

  const vals = sh.getDataRange().getValues();
  const hdr  = vals[0].map(h => String(h || '').trim());
  const lower= hdr.map(h => h.toLowerCase());

  // Header indices (email is optional and ignored here)
  const iPT   = lower.indexOf('ptin');
  const iProg = lower.indexOf('program number');
  const iDR   = lower.indexOf('date reported');
  const iPCD  = lower.indexOf('program completion date');

  if (iPT < 0 || iProg < 0) {
    if (!quiet) toast_('Reported Hours missing PTIN and/or Program Number columns.', true);
    return;
  }

  // Helper: robust date → ms since epoch. Falls back to NaN if unparseable.
  function ts(v) {
    // Prefer your global parseDate_ if present
    const d = (typeof parseDate_ === 'function') ? parseDate_(v) : new Date(v);
    return (d instanceof Date && !isNaN(d.getTime())) ? d.getTime() : NaN;
  }

  // Walk body and keep the most recent by (PTIN|PROGRAM)
  const body = vals.slice(1);
  const keepMap = new Map();  // key -> { row, idx, scoreTs }
  for (let r = 0; r < body.length; r++) {
    const row = body[r];

    const ptin = String(row[iPT] || '').trim().toUpperCase();
    const prog = String(row[iProg] || '').trim().toUpperCase().replace(/\s+/g, '');
    if (!ptin || !prog) continue;

    const key = prog + '|' + ptin;

    // Compute recency score
    const dr  = iDR  >= 0 ? ts(row[iDR])  : NaN;
    const pcd = iPCD >= 0 ? ts(row[iPCD]) : NaN;

    // score: prefer Date Reported; if NaN, use Program Completion Date; if both NaN, use row index
    let scoreTs;
    if (!isNaN(dr)) scoreTs = dr;
    else if (!isNaN(pcd)) scoreTs = pcd;
    else scoreTs = -Infinity; // will fall back to row order below

    const current = keepMap.get(key);
    if (!current) {
      keepMap.set(key, { row, idx: r, scoreTs, rowOrder: r });
      continue;
    }

    // Decide winner:
    // 1) higher scoreTs wins
    // 2) if equal/NaN, later row (greater rowOrder) wins
    const a = current;
    const b = { row, idx: r, scoreTs, rowOrder: r };

    const aScore = a.scoreTs;
    const bScore = b.scoreTs;

    let chooseB = false;
    if (aScore === bScore) {
      // tie or both NaN → later row wins
      chooseB = (b.rowOrder > a.rowOrder);
    } else {
      // strict greater timestamp wins; treat NaN as lowest
      const aEff = isNaN(aScore) ? -Infinity : aScore;
      const bEff = isNaN(bScore) ? -Infinity : bScore;
      chooseB = (bEff > aEff);
    }
    if (chooseB) keepMap.set(key, b);
  }

  // Rebuild final set of winners, sorted by their original row order (stable & readable)
  const winners = Array.from(keepMap.values())
    .sort((x,y) => x.rowOrder - y.rowOrder)
    .map(x => x.row);

  // Rewrite the sheet body (headers untouched)
  const numCols = hdr.length;
  const oldRows = sh.getLastRow() - 1;
  if (oldRows > 0) sh.getRange(2, 1, oldRows, numCols).clearContent();
  if (winners.length) sh.getRange(2, 1, winners.length, numCols).setValues(winners);

  if (!quiet) toast_(`Reported Hours de-duplicated (PTIN+Program): kept ${winners.length}, removed ${body.length - winners.length}.`);
}
