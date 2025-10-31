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
/** Public entrypoint for the menu item: “Sync Group Sheets (strict)” */
function syncGroupSheets() {
  try {
    // Prefer a strict/primary implementation if present
    if (typeof syncGroupSheetsStrict === 'function') return syncGroupSheetsStrict();
    if (typeof doGroupSyncAll_ === 'function')       return doGroupSyncAll_();
    if (typeof runGroupSync === 'function')          return runGroupSync();
    if (typeof groupSyncMain === 'function')         return groupSyncMain();
    if (typeof syncGroupsFlexible === 'function')    return syncGroupsFlexible();

    // Fallback: helpful diagnostics so we know what's actually defined
    const fns = Object.keys(this)
      .filter(k => typeof this[k] === 'function' && /group.*sync|sync.*group|doGroup/i.test(k))
      .sort();
    toast_(
      'No concrete group sync function found for syncGroupSheets(). ' +
      (fns.length ? 'Candidates: ' + fns.join(', ') : 'No group-sync-like functions detected.'),
      true
    );
  } catch (err) {
    toast_('syncGroupSheets failed: ' + (err && err.message ? err.message : err), true);
    Logger.log(err && err.stack ? err.stack : err);
  }
}

/** Alias used by menus that reference syncGroupSheetsMenu */
function syncGroupSheetsMenu() {
  return syncGroupSheets();
}

/** Ensure global exports in Apps Script runtime (defensive) */
try {
  this.syncGroupSheets = this.syncGroupSheets || syncGroupSheets;
  this.syncGroupSheetsMenu = this.syncGroupSheetsMenu || syncGroupSheetsMenu;
} catch (e) {
  // no-op: some runtimes may not allow assigning to `this`
}

/**
 * === Strict Group Sync ===
 * Reads Master, filters rows by Group, and writes each group's rows to its linked sheet.
 * - Uses Courses sheet to convert Program Number -> Program Name.
 * - Accepts either Program Name or Program Number columns on target.
 * - Keeps Reporting Issue? rows (do NOT filter them out).
 * - Will clear & rewrite the data body only (keeps header row).
 *
 * Requires:
 *  - CFG.SHEET_MASTER (e.g., "Master")
 *  - A catalog sheet named "Groups" with columns: Group, Sheet URL  (case-insensitive)
 *  - A sheet named "Courses" with columns: Program Number, Program Name
 */
function syncGroupSheetsStrict() {
  const ss = SpreadsheetApp.getActive();

  // --- Load Master
  const master = mustGet_(ss, CFG.SHEET_MASTER);
  const mVals = master.getDataRange().getValues();
  if (mVals.length <= 1) { toast_('Master is empty; nothing to sync.', true); return; }
  const mHdr = normalizeHeaderRow_(mVals[0]);
  const mm   = mapHeaders_(mHdr);

  if (mm.group == null) { toast_('Master is missing "Group" column.', true); return; }

  // Build courses map: Program Number -> Program Name
  const coursesMap = loadCoursesMap_((ss));

  // Read group catalog (Group, Sheet URL)
  const groups = readGroupsCatalog_(ss);
  if (!groups.length) { toast_('No groups found in the "Groups" catalog.', true); return; }

  const mBody = mVals.slice(1);
  let totalSheets = 0, totalCells = 0, visited = 0;

  groups.forEach(entry => {
    visited++;
    const gName = String(entry.groupName || entry.group || '').trim();
    const gUrl  = String(entry.url || '').trim();
    const gId   = String(entry.groupId || '').trim();
    if (!gName || !gUrl) {
      Logger.log(`Group catalog entry missing name or url: ${JSON.stringify(entry)}`);
      return;
    }

    // Determine Master columns for group name / id
    const iGroupName = (mm.group != null) ? mm.group : findHeaderIndex_(mHdr, 'Group Name');
    const iGroupId   = findHeaderIndex_(mHdr, 'Group ID');

    // Filter Master rows for this group (match by name, or by ID if both sides have IDs)
    const rows = mBody.filter(r => {
      const name = iGroupName >= 0 ? String(r[iGroupName] || '').trim() : '';
      const id   = iGroupId   >= 0 ? String(r[iGroupId]   || '').trim() : '';
      return (name && name === gName) || (gId && id && id === gId);
    });

    if (!rows.length) {
      Logger.log(`Group "${gName}": no Master rows; skipping.`);
      return;
    }

    let targetSS;
    try {
      targetSS = openSpreadsheetByUrlOrId_(gUrl);
    } catch (e) {
      Logger.log(`Group "${gName}": cannot open target (${gUrl}): ${e.message}`);
      return;
    }

    const targetSheet = targetSS.getSheets()[0]; // convention: first sheet
    if (!targetSheet) {
      Logger.log(`Group "${gName}": no sheets found in target spreadsheet.`);
      return;
    }

    // Ensure target sheet headers are mapped; accept Program Name OR Program Number
    const gMap = mapGroupHeadersFlexible_(targetSheet);
    if (!gMap.ok) {
      Logger.log(`Group "${gName}": missing headers: ${gMap.missing.join(', ')}`);
      toast_(`Group "${gName}" missing required headers: ${gMap.missing.join(', ')}`, true);
      return;
    }

    // Build output rows aligned to target header order
    const out = [];
    const hdr = targetSheet.getRange(1,1,1,targetSheet.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
    const lower = hdr.map(h=>h.toLowerCase());

    // Helper to place a value by header label (case-insensitive)
    const placeByHeader = (arr, label, value) => {
      const i = lower.indexOf(String(label||'').toLowerCase());
      if (i >= 0) arr[i] = value;
    };

    // We’ll try to read fields from Master using its header map
    const wantDate = (v) => (v instanceof Date) ? v : (parseDate_(v) || v);

    rows.forEach(row => {
      const arr = new Array(hdr.length).fill('');

      const first = mm.firstName != null ? String(row[mm.firstName]||'').trim() : '';
      const last  = mm.lastName  != null ? String(row[mm.lastName] ||'').trim() : '';
      const ptin  = mm.ptin      != null ? formatPtinP0_(row[mm.ptin]||'')      : '';
      const email = mm.email     != null ? String(row[mm.email]    ||'').toLowerCase().trim() : '';
      const progN = mm.program   != null ? String(row[mm.program]  ||'').toUpperCase().replace(/\s+/g,'').trim() : '';
      const hrs   = mm.hours     != null ? row[mm.hours] : '';
      const comp  = mm.completion!= null ? wantDate(row[mm.completion]) : '';
      const issue = mm.masterIssueCol != null ? String(row[mm.masterIssueCol]||'').trim() : '';
      const reported    = mm.reportedCol    != null ? row[mm.reportedCol] : '';
      const reportedAt  = mm.reportedAtCol  != null ? wantDate(row[mm.reportedAtCol]) : '';

      // Convert Program Number -> Program Name via Courses map (fallback to Master.Program if not found)
      const progName = coursesMap.get(progN) || '';

      // Place values according to the target sheet headers
      placeByHeader(arr, 'Attendee First Name', first);
      placeByHeader(arr, 'Attendee Last Name',  last);
      placeByHeader(arr, 'Attendee PTIN',       ptin);
      placeByHeader(arr, 'Email',               email);

      // Prefer Program Name if the target has it; otherwise use Program Number
      if (gMap.progName >= 0) {
        placeByHeader(arr, 'Program Name', progName || progN);
      } else if (gMap.prog >= 0) {
        placeByHeader(arr, 'Program Number', progN);
      }

      // Hours / Completion
      placeByHeader(arr, 'CE Hours Awarded', hrs);
      placeByHeader(arr, 'CE Hours', hrs); // some sheets use this
      placeByHeader(arr, 'Program Completion Date', comp);

      // Issues / Reporting
      placeByHeader(arr, 'Reporting Issue?', issue);
      placeByHeader(arr, 'Reported?', truthy_(reported));
      placeByHeader(arr, 'Reported At', reportedAt);

      out.push(arr);
    });

    // Clear body & write
    const lastRow = targetSheet.getLastRow();
    const lastCol = targetSheet.getLastColumn();
    if (lastRow > 1) targetSheet.getRange(2,1,lastRow-1,lastCol).clearContent();

    if (out.length) {
      targetSheet.getRange(2,1,out.length, lastCol).setValues(out);
      // format date columns if present
      const iComp = lower.indexOf('program completion date');
      if (iComp >= 0) targetSheet.getRange(2, iComp+1, out.length, 1).setNumberFormat('mm/dd/yyyy');
      const iRepAt = lower.indexOf('reported at');
      if (iRepAt >= 0) targetSheet.getRange(2, iRepAt+1, out.length, 1).setNumberFormat('mm/dd/yyyy');
    }

    totalSheets++;
    totalCells += out.length * lastCol;
    Logger.log(`Group "${gName}" (${gId || 'no-id'}): wrote ${out.length} row(s) to ${targetSS.getName()}.`);
  });

  toast_(`Group sync done: visited ${visited}, wrote ${totalSheets} sheet(s), ~${totalCells} cells.`);
}

/** Open a spreadsheet by URL or Spreadsheet ID */
function openSpreadsheetByUrlOrId_(ref) {
  const s = String(ref||'').trim();
  if (!s) throw new Error('Empty spreadsheet reference.');
  if (/^https?:\/\//i.test(s)) return SpreadsheetApp.openByUrl(s);
  return SpreadsheetApp.openById(s);
}

/** Read "Courses" into Map: Program Number -> Program Name */
function loadCoursesMap_(ss) {
  const sh = ss.getSheetByName('Courses');
  const map = new Map();
  if (!sh || sh.getLastRow() < 2) return map;
  const vals = sh.getDataRange().getValues();
  const hdr  = normalizeHeaderRow_(vals[0]);
  const lower= hdr.map(h=>h.toLowerCase());
  const iNum = lower.indexOf('program number');
  const iNam = lower.indexOf('program name');
  if (iNum < 0 || iNam < 0) return map;
  for (let r=1;r<vals.length;r++){
    const num = String(vals[r][iNum]||'').toUpperCase().replace(/\s+/g,'').trim();
    const nam = String(vals[r][iNam]||'').trim();
    if (num) map.set(num, nam);
  }
  return map;
}

/** Read "Groups" or "Group Config" catalog: returns [{groupName, url, groupId}] */
function readGroupsCatalog_(ss) {
  // Prefer "Group Config" (with headers: Group ID, Group Name, Spreadsheet URL)
  let sh = ss.getSheetByName('Group Config');
  if (!sh || sh.getLastRow() < 2) {
    // Fallback to legacy "Groups" (headers: Group, Sheet URL)
    sh = ss.getSheetByName('Groups');
  }
  if (!sh || sh.getLastRow() < 2) return [];

  const vals = sh.getDataRange().getValues();
  const hdr  = normalizeHeaderRow_(vals[0]);
  const lower= hdr.map(h=>h.toLowerCase());

  // Flexible header candidates
  const idxOfAny = (names) => {
    for (const n of names) {
      const i = lower.indexOf(String(n).toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };

  const iId  = idxOfAny(['group id','id']);
  const iNam = idxOfAny(['group name','group']);
  const iUrl = idxOfAny(['spreadsheet url','sheet url','url']);

  if (iNam < 0 || iUrl < 0) {
    toast_('Group catalog missing "Group Name/Group" and/or "Spreadsheet URL/Sheet URL" headers.', true);
    return [];
  }

  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const groupName = String(vals[r][iNam] || '').trim();
    const url       = String(vals[r][iUrl] || '').trim();
    const groupId   = iId >= 0 ? String(vals[r][iId] || '').trim() : '';
    if (groupName && url) out.push({ groupName, url, groupId });
  }
  return out;
}

/** Export new function to global for menu handler */
try {
  this.syncGroupSheetsStrict = this.syncGroupSheetsStrict || syncGroupSheetsStrict;
} catch (e) {}