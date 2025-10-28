/** Safety shims so missing helpers don’t crash this file **/
if (typeof mapFreeTextToStandardIssue_ !== 'function') {
  function mapFreeTextToStandardIssue_(txt){
    const s = String(txt||'').trim().toLowerCase();
    if (!s) return '';
    if (s.includes('missing') && s.includes('ptin')) return 'Missing PTIN';
    if (s.includes('does not exist') || s.includes('invalid ptin') || s.includes('ptin invalid') || s.includes('all zero')) {
      return 'PTIN does not exist';
    }
    if (s.includes('mismatch') || s.includes('name mismatch') || s.includes('ptin & name') || s.includes('ptin/name')) {
      return 'PTIN & name do not match';
    }
    if (typeof CFG !== 'undefined' && CFG.REPORTING_ISSUE_CHOICES) {
      const hit = CFG.REPORTING_ISSUE_CHOICES.find(x => String(x).toLowerCase() === s);
      if (hit) return hit;
    }
    return 'Other';
  }
}
if (typeof syncMasterFromIssueSheet_ !== 'function') {
  function syncMasterFromIssueSheet_(){ /* no-op: legacy sheet removed */ }
}
if (typeof applyReportingIssueValidationAndFormatting_ !== 'function') {
  function applyReportingIssueValidationAndFormatting_(){ /* no-op */ }
}
if (typeof updateRosterValidityFromIssues_ !== 'function') {
  function updateRosterValidityFromIssues_(){ /* no-op */ }
}

function ingestSystemReportingIssues(quiet){
  const ss = SpreadsheetApp.getActive();
  const sys = ss.getSheetByName(CFG.SHEET_SYS_ISSUES);
  if (!sys) { if(!quiet) toast_(`Sheet "${CFG.SHEET_SYS_ISSUES}" not found.`, true); return; }

  const sVals = sys.getDataRange().getValues();
  if (sVals.length <= 1) { if(!quiet) toast_('No rows on System Reporting Issues.', true); return; }

  const shdr = sVals[0].map(s=>String(s||'').trim());
  const sIdx = (h)=> shdr.indexOf(h);
  const sF = sIdx('Attendee First Name');
  const sL = sIdx('Attendee Last Name');
  const sP = sIdx('PTIN');
  const sG = sIdx('Program Number');
  const sH = sIdx('CE Hours Awarded');
  const sC = sIdx('Program Completion Date');
  const sS = sIdx('Status');
  if ([sG,sS].some(i=>i<0)) { if(!quiet) toast_('System Reporting Issues missing required headers (Program Number, Status).', true); return; }

  // Collect sys entries
  const entries = [];
  for (let r=1;r<sVals.length;r++){
    const row = sVals[r];
    const first = String(row[sF]||'').trim();
    const last  = String(row[sL]||'').trim();
    const ptin  = formatPtinP0_(row[sP]||'');
    const prog  = String(row[sG]||'').toUpperCase().replace(/\s+/g,'').trim();
    const hours = row[sH];
    const comp  = row[sC];
    const issue = mapFreeTextToStandardIssue_(row[sS]||'');
    if (!prog || !issue) continue;
    entries.push({first,last,ptin,prog,hours,comp,issue});
  }
  if (!entries.length) { if(!quiet) toast_('No usable System Reporting Issues found.', true); return; }

  const master = ss.getSheetByName(CFG.SHEET_MASTER);
  const issues = ss.getSheetByName(CFG.SHEET_ISSUES);
  const clean  = ss.getSheetByName(CFG.SHEET_CLEAN);

  // ---- Master update
  if (master && master.getLastRow() > 1){
    const mVals = master.getDataRange().getValues();
    const mh = normalizeHeaderRow_(mVals[0]);
    const mm = mapHeaders_(mh);
    const body = mVals.slice(1);
    let touched = 0;

    const emailIx = mm.email;
    const firstIx = mm.firstName;
    const lastIx  = mm.lastName;

    function matchAndMark_(entry) {
      let any = 0;

      // A) PTIN + Program
      if (entry.ptin) {
        any += updateAll_((row)=> formatPtinP0_(row[mm.ptin]||'')===entry.ptin &&
                               String(row[mm.program]||'').toUpperCase().replace(/\s+/g,'').trim()===entry.prog);
      }
      // C) Name + Program fallback
      if (!any && firstIx!=null && lastIx!=null) {
        const fL = entry.first.toLowerCase(); const lL = entry.last.toLowerCase();
        any += updateAll_((row)=> {
          const p = String(row[mm.program]||'').toUpperCase().replace(/\s+/g,'').trim()===entry.prog;
          const f = String(row[firstIx]||'').trim().toLowerCase()===fL;
          const l = String(row[lastIx] ||'').trim().toLowerCase()===lL;
          return p && f && l;
        });
      }
      return any;

      function updateAll_(predicate){
        let hits=0;
        for (let i=0;i<body.length;i++){
          const row = body[i];
          if (!predicate(row)) continue;

          // Set the issue to the ingested one; do not clear “PTIN does not exist” unless it’s being overwritten with something else.
          if (mm.masterIssueCol!=null) row[mm.masterIssueCol] = entry.issue;

          // Un-report on a system issue
          if (mm.reportedCol!=null)    row[mm.reportedCol]    = false;
          if (mm.reportedAtCol!=null)  row[mm.reportedAtCol]  = '';
          hits++;
        }
        return hits;
      }
    }

    for (const e of entries) touched += matchAndMark_(e);

    if (touched) master.getRange(2,1,body.length,mVals[0].length).setValues(body);
  }

  // ---- Clean tagging so those rows stay out of export
  if (clean && clean.getLastRow() > 1){
    const cVals = clean.getDataRange().getValues();
    const ch = cVals[0].map(s=>String(s||'').trim());
    const ci = {
      f: ch.indexOf('Attendee First Name'),
      l: ch.indexOf('Attendee Last Name'),
      p: ch.indexOf('Attendee PTIN'),
      g: ch.indexOf('Program Number'),
      ri: ch.indexOf('Reporting Issue?')
    };
    if (ci.g>=0 && ci.ri>=0){
      const body = cVals.slice(1);
      let updates = 0;
      for (let i=0;i<body.length;i++){
        const row = body[i];
        const pt = formatPtinP0_(row[ci.p]||'');
        const pr = String(row[ci.g]||'').toUpperCase().replace(/\s+/g,'').trim();
        const fn = ci.f>=0 ? String(row[ci.f]||'').trim().toLowerCase() : '';
        const ln = ci.l>=0 ? String(row[ci.l]||'').trim().toLowerCase() : '';

        const hit = entries.find(e =>
          (e.prog===pr) && (
            (pt && e.ptin && e.ptin===pt) ||
            (fn && ln && e.first && e.last && e.first.toLowerCase()===fn && e.last.toLowerCase()===ln)
          )
        );
        if (hit) { row[ci.ri] = hit.issue; updates++; }
      }
      if (updates) clean.getRange(2,1,body.length, cVals[0].length).setValues(body);
    }
  }

  // ---- Reporting Issues sheet append (optional; keep if you still track it)
  if (issues){
    const toAppend = [];
    const m = ss.getSheetByName(CFG.SHEET_MASTER);
    let meta = new Map();
    if (m && m.getLastRow()>1){
      const mv = m.getDataRange().getValues();
      const mh = normalizeHeaderRow_(mv[0]);
      const mm = mapHeaders_(mh);
      meta = new Map();
      for (let r=1;r<mv.length;r++){
        const row = mv[r];
        const p = formatPtinP0_(row[mm.ptin]||'');
        const g = String(row[mm.program]||'').toUpperCase().replace(/\s+/g,'').trim();
        const em= mm.email!=null ? String(row[mm.email]||'').toLowerCase().trim() : '';
        const sc= mm.source!=null ? String(row[mm.source]||'').trim() : '';
        if (p && g) meta.set(p+'|'+g, {email: em, source: sc});
      }
    }
    for (const e of entries){
      const mmeta = meta.get((e.ptin||'')+'|'+e.prog) || {email:'',source:''};
      toAppend.push({
        'Attendee First Name': e.first,
        'Attendee Last Name':  e.last,
        'Attendee PTIN':       e.ptin,
        'Program Number':      e.prog,
        'CE Hours Awarded':    e.hours,
        'Program Completion Date': formatToMDY_(e.comp),
        'Email':               mmeta.email,
        'Source':              mmeta.source,
        'Reporting Issue':     e.issue,
        'Fixed?':              ''
      });
    }
    if (toAppend.length) appendIssuesRows_(ss, toAppend);
  }

  // Final harmonization (now safe even if helpers weren’t loaded elsewhere)
  syncMasterFromIssueSheet_();
  applyReportingIssueValidationAndFormatting_();
  updateRosterValidityFromIssues_();

  if (!quiet) toast_(`Ingested ${entries.length} System Reporting Issue row(s); Master & Issues updated.`);
}