function auditP00000000Rows() {
  const sh = mustGet_(SpreadsheetApp.getActive(), CFG.SHEET_MASTER);
  const vals = sh.getDataRange().getValues();
  const hdr = normalizeHeaderRow_(vals[0]);
  const mm  = mapHeaders_(hdr);
  let misses = 0;
  for (let r=1; r<vals.length; r++){
    const ptin = String(vals[r][mm.ptin]||'').trim().toUpperCase();
    const issue = String(vals[r][mm.masterIssueCol]||'').trim();
    if (ptin === 'P00000000' && issue !== 'PTIN does not exist') {
      misses++;
    }
  }
  toast_(`Rows with P00000000 missing sticky error: ${misses}`);
}