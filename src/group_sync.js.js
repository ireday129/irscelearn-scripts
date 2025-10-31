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

  // Flexible aliases (case-insensitive) — broadened to avoid false "header mismatch"
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

  // Determine if the sheet has enough to work with:
  const haveIdentity = (map.first >= 0 && map.last >= 0);
  const haveKey      = (map.ptin >= 0 || map.email >= 0);          // allow either PTIN or Email as the match key
  const haveProgram  = (map.progName >= 0 || map.prog >= 0);       // allow Program Name OR Program Number

  const requiredMissing = [];
  if (!haveIdentity) {
    if (map.first < 0) requiredMissing.push('Attendee First Name');
    if (map.last  < 0) requiredMissing.push('Attendee Last Name');
  }
  if (!haveKey) requiredMissing.push('Attendee PTIN or Email');
  if (!haveProgram) requiredMissing.push('Program Name or Program Number');

  const ok = requiredMissing.length === 0;

  return { ok, missing: requiredMissing, ...map };
}
