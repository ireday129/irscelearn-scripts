/** STRICT-ISH GROUP SYNC (supports "Reporting" OR "Reporting Info") **/

function syncGroupSheets(quiet) {
  try {
    const ss = SpreadsheetApp.getActive();
    const master = mustGet_(ss, CFG.SHEET_MASTER);
    const cfg = readGroupConfigStrict_(); // { [groupId]: {name,url} }
    if (Object.keys(cfg).length === 0) { if (!quiet) toast_('Group Config has no rows.', true); return; }

    const mVals = master.getDataRange().getValues();
    if (mVals.length <= 1) { if(!quiet) toast_('Master is empty.', true); return; }
    const mHdr = normalizeHeaderRow_(mVals[0]);
    const mMap = mapHeaders_(mHdr);

    // START FIX: Override external mapHeaders_ failure for the 'group' column
    if (mMap.group == null) {
      const lowerHdr = mHdr.map(h => h.toLowerCase());
      // Explicitly look for the 'Group' header
      const groupIndex = lowerHdr.indexOf('group');
      if (groupIndex !== -1) {
        mMap.group = groupIndex;
      }
    }
    // END FIX

    const courseNameByNumber = readCoursesProgramMap_();

    const need = ['firstName','lastName','ptin','email','program','hours','completion','group','reportedCol','reportedAtCol','masterIssueCol'];
    const missing = need.filter(k => mMap[k] == null);
    if (missing.length) { if (!quiet) toast_('Master missing columns for Group sync: ' + missing.join(', '), true); return; }

    // Build group buckets from Master (Group column now contains Group ID)
    const byGroup = {};
    const body = mVals.slice(1);
    for (let r=0; r<body.length; r++) {
      const row = body[r];
      const gid = String(row[mMap.group] || '').trim(); // FIX: Using mMap.group instead of mMap.source
      if (!gid || !cfg[gid]) continue;

      const rec = {
        first: String(row[mMap.firstName]||'').trim(),
        last:  String(row[mMap.lastName]||'').trim(),
        ptin:  formatPtinP0_(row[mMap.ptin]||''),
        prog:  String(row[mMap.program]||'').toUpperCase().replace(/\s+/g,'').trim(),
        progName: null,
        hours: row[mMap.hours],
        email: String(row[mMap.email]||'').toLowerCase().trim(),
        comp:  formatToMDY_(row[mMap.completion]),
        issue: String(row[mMap.masterIssueCol]||'').trim(),
        reported: parseBool_(row[mMap.reportedCol]) ? true : false,
        reportedAt: row[mMap.reportedAtCol] instanceof Date ? row[mMap.reportedAtCol] : ''
      };
      rec.progName = courseNameByNumber.get(rec.prog) || rec.prog;
      if (!byGroup[gid]) byGroup[gid] = [];
      byGroup[gid].push(rec);
    }
    // De-dupe per group with fallbacks:
    // Priority key: (1) PTIN+Program, (2) Email+Program, (3) Name+Program
    // Winner = latest by (Reported At || Program Completion Date)
    Object.keys(byGroup).forEach(gid => {
      const arr = byGroup[gid];
      const pick = new Map();

      function bestTimestamp(rec) {
        if (rec && rec.reportedAt instanceof Date) return rec.reportedAt.getTime();
        // fallback to completion date (parsed if possible)
        const c = rec && rec.comp ? parseDate_(rec.comp) : null;
        return c instanceof Date ? c.getTime() : -1;
      }

      function makeKey(rec) {
        const prog = rec.prog || '';
        const ptin = rec.ptin || '';
        const email = rec.email || '';
        const first = (rec.first || '').trim().toLowerCase();
        const last  = (rec.last  || '').trim().toLowerCase();

        if (ptin && prog)  return `PTIN:${ptin}|${prog}`;
        if (email && prog) return `EMAIL:${email}|${prog}`;
        if (first && last && prog) return `NAME:${first} ${last}|${prog}`;
        // no stable dedupe key; return empty string to indicate "keep as-is"
        return '';
      }

      const out = [];
      for (let i = 0; i < arr.length; i++) {
        const rec = arr[i];
        const key = makeKey(rec);

        if (!key) {
          // Cannot dedupe reliably; include this row as-is
          out.push(rec);
          continue;
        }

        if (!pick.has(key)) {
          pick.set(key, i);
        } else {
          const prevIdx = pick.get(key);
          const prev = arr[prevIdx];
          if (bestTimestamp(rec) >= bestTimestamp(prev)) {
            pick.set(key, i);
          }
        }
      }

      // keep winners + all non-deduped
      for (const [, idx] of pick) out.push(arr[idx]);
      byGroup[gid] = out;
    });

    let groupsUpdated = 0, rowsUpdated = 0, groupsSkipped = 0;

    for (const gid of Object.keys(cfg)) {
      // If no rows for this GID, still clear their Reporting tab (body only) to reflect current truth
      const rows = byGroup[gid] || [];

      const {url, name} = cfg[gid];
      let ext;
      try { ext = SpreadsheetApp.openByUrl(url); }
      catch (e) { if (!quiet) toast_(`Group "${name || gid}": URL not accessible (share access?). Skipped.`, true); groupsSkipped++; continue; }

      const {sheet: sh, label: foundName} = getGroupTargetSheet_(ext);
      if (!sh) { if (!quiet) toast_(`Group "${name || gid}": missing "Reporting" and "Reporting Info" tabs. Skipped.`, true); groupsSkipped++; continue; }

      // Tolerant header mapping
      const map = mapGroupHeadersFlexible_(sh);
      if (!map.ok) {
        if (!quiet) toast_(`Group "${name || gid}" (${foundName}): header mismatch. Missing: ${map.missing.join(', ')}`, true);
        groupsSkipped++;
        continue;
      }

      // Clear body only, never touch headers
      if (sh.getLastRow() > 1) sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
      if (!rows.length) { groupsUpdated++; continue; }

      const data = rows.map(o => {
        const arr = new Array(sh.getLastColumn()).fill('');
        if (map.first>=0) arr[map.first] = o.first;
        if (map.last>=0)  arr[map.last]  = o.last;
        if (map.ptin>=0)  arr[map.ptin]  = o.ptin;
        if (map.progName>=0) arr[map.progName] = o.progName || o.prog;
        else if (map.prog>=0)  arr[map.prog]  = o.progName || o.prog;
        if (map.hours>=0) arr[map.hours] = o.hours;
        if (map.email>=0) arr[map.email] = o.email;
        if (map.comp>=0)  arr[map.comp]  = o.comp;
        if (map.issue>=0) arr[map.issue] = o.issue;
        // FIX: Write the actual boolean value (true/false) so it works with checkbox formatting
        if (map.reported>=0)   arr[map.reported]   = o.reported; 
        if (map.reportedAt>=0) arr[map.reportedAt] = o.reportedAt;
        return arr;
      });

      sh.getRange(2,1,data.length, sh.getLastColumn()).setValues(data);
      
      // FIX: Apply Checkbox Validation to the Reported? column
      if (map.reported >= 0 && data.length > 0) {
        const reportedRange = sh.getRange(2, map.reported + 1, data.length, 1);
        setCheckboxValidation_(reportedRange);
      }

      // Formats (safe)
      if (map.comp>=0) sh.getRange(2, map.comp+1, data.length, 1).setNumberFormat('mm/dd/yyyy');
      if (map.reportedAt>=0) sh.getRange(2, map.reportedAt+1, data.length, 1).setNumberFormat('mm/dd/yyyy hh:mm am/pm');

      groupsUpdated++;
      rowsUpdated += rows.length;
    }

    const msg = `Group sync: updated ${groupsUpdated} sheet(s), ${rowsUpdated} row(s); skipped ${groupsSkipped}.`;
    if (!quiet) toast_(msg);
    Logger.log(msg);
  } catch (err) {
    toast_('Group sync failed: ' + err.message, true);
    Logger.log(err.stack || err.message);
  }
}

/** Try to find either "Reporting" or "Reporting Info". Returns {sheet, label}. */
function getGroupTargetSheet_(extSpreadsheet) {
  const preferred = [CFG.GROUP_TARGET_SHEET, 'Reporting Info'];
  for (const name of preferred) {
    const sh = extSpreadsheet.getSheetByName(name);
    if (sh) return {sheet: sh, label: name};
  }
  return {sheet: null, label: ''};
}

/** Read Group Config: Group ID | Group Name | Spreadsheet URL */
function readGroupConfigStrict_() {
  const ss = SpreadsheetApp.getActive();
  const sh = mustGet_(ss, CFG.SHEET_GROUP_CONFIG);
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return {};

  const hdr = vals[0].map(s=>String(s||'').trim().toLowerCase());
  const iId   = hdr.indexOf('group id');
  const iName = hdr.indexOf('group name');
  const iUrl  = hdr.indexOf('spreadsheet url');
  if (iId<0 || iName<0 || iUrl<0) throw new Error('Group Config headers must be: Group ID, Group Name, Spreadsheet URL');

  const out = {};
  for (let r=1;r<vals.length;r++){
    const row = vals[r];
    const id  = String(row[iId]||'').trim();
    const nm  = String(row[iName]||'').trim();
    const url = String(row[iUrl]||'').trim();
    if (id && url) out[id] = {name: (nm || id), url};
  }
  return out;
}

/** Map group "Reporting"/"Reporting Info" headers with common variants/typos (case-insensitive). */
function mapGroupHeadersFlexible_(sheet){
  const hdr = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const lower = hdr.map(h=>h.toLowerCase());

  function idxOfAny(names){
    for (const n of names){
      const j = lower.indexOf(n.toLowerCase());
      if (j>=0) return j;
    }
    return -1;
  }

  // Known variants based on your spec (and real-world typos)
  const map = {
    first:     idxOfAny(['Attended First Name','Attendee First Name']),
    last:      idxOfAny(['Attendee Last Name']),
    ptin:      idxOfAny(['attendee PTIN','Attendee PTIN','PTIN']),
    prog:      idxOfAny(['Program Number']),
    progName:  idxOfAny(['Program Name']),
    hours:     idxOfAny(['CE hours Awards','CE Hours Awards','CE hours Awarded','CE Hours Awarded']),
    email:     idxOfAny(['email','Email']),
    comp:      idxOfAny(['program completion date','Program Completion Date']),
    issue:     idxOfAny(['Reporting Issue?']),
    reported:  idxOfAny(['Reported?']),
    reportedAt:idxOfAny(['Reported At'])
  };

  const missing = Object.keys(map).filter(k => map[k] < 0);
  return { ok: missing.length===0, missing, ...map };
}

/** Helper function to apply TRUE/FALSE validation, rendering a checkbox. */
function setCheckboxValidation_(range) {
  const rule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false) // Must be true/false only
      .build();
  range.setDataValidation(rule);
}

function readCoursesProgramMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Courses');
  const map = new Map();
  if (!sh || sh.getLastRow() < 2) return map;

  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(s => String(s || '').trim().toLowerCase());
  const iNum = hdr.indexOf('program number');
  const iName = hdr.indexOf('program name');
  if (iNum < 0 || iName < 0) return map;

  for (let r = 1; r < vals.length; r++) {
    const num = String(vals[r][iNum] || '').toUpperCase().replace(/\s+/g, '').trim();
    const name = String(vals[r][iName] || '').trim();
    if (num && name) map.set(num, name);
  }
  return map;
}
