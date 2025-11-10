/** IRS CE TOOLS MAIN MENU **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('EarnTaxCE Tools');

  menu
    .addItem('Build Clean Upload (resumable)', 'buildCleanUpload')
    .addItem('Recheck Master for Issues', 'recheckMaster')
    .addSeparator()
    .addItem('Create Nightly Clean Trigger', 'createNightlyTrigger')
    .addItem('Remove Nightly Clean Trigger', 'removeNightlyTrigger')
    .addItem('Create Nightly Group Sync Trigger', 'createNightlyGroupSyncTrigger')
    .addItem('Remove Nightly Group Sync Trigger', 'removeNightlyGroupSyncTrigger')
    .addSeparator()
    .addSeparator()
    .addItem('Sync Group Sheets (strict)', 'syncGroupSheets')
    .addItem('Diagnose Group Sync', 'diagnoseGroupSync')
    .addSeparator()
    .addItem('Ingest System Reporting Issues', 'ingestSystemReportingIssues')
    .addSeparator()
    .addItem('Update Reporting Stats', 'updateReportingStatsMenu')
    .addItem('Sync Reported Hours \u2192 Master', 'syncReportedToMasterMenu') // ← new menu item
    .addSeparator()
    .addItem('Highlight Roster from Reported Hours', 'highlightRosterFromReportedHoursMenu')
    .addItem('Deduplicate Roster by Email', 'dedupeRosterByEmail')
    .addItem('Backfill Master from Roster', 'backfillMasterFromRosterCombined') // COMBINED ITEM
    .addSeparator()
    .addItem('Create/Repair Tabs (Clean & Issues only)', 'ensureAllTabs')
    .addSeparator()
    .addItem('Diagnostics (Log all systems)', 'runDiagnostics')
    .addSeparator()
    .addItem('Sanity Scan: Helpers Present', 'sanityScanHelpers');
  menu.addToUi();
}

/**
 * Sanity Scan: verifies required helpers and key sheets exist.
 * Results go to Logs and a short toast shows counts.
 */
function sanityScanHelpers() {
  const ss = SpreadsheetApp.getActive();

  // ---- declare the helpers you expect to exist
  const requiredHelpers = [
    'toast_',
    'mustGet_',
    'normalizeHeaderRow_',
    'mapHeaders_',
    'mapCleanHeaders_',
    'parseDate_',
    'formatToMDY_',
    'formatPtinP0_',
    'normalizeProgram_',
    'namesMatchFull_',
    'normalizeCompletionForUpload_',
    'writeCleanDataOnly_',
    'buildUnresolvedIssueIndex_',

    // workflow functions you call elsewhere (list only the ones you actually use)
    'buildCleanUpload',
    'recheckMaster',
    'markCleanAsReported',
    'updateProgramReportedTotals',
    'syncMasterWithReportedHours',
    'dedupeReportedHoursByPtinProgram',
    'highlightRosterFromReportedHours',
    'triggerRosterValidWebhookMaybe' // optional—remove if you renamed it
  ];

  // ---- key sheets you rely on (change these if your CFG is different)
  const requiredSheets = [
    (typeof CFG !== 'undefined' && CFG && CFG.SHEET_MASTER) || 'Master',
    (typeof CFG !== 'undefined' && CFG && CFG.SHEET_CLEAN) || 'Clean',
    (typeof CFG !== 'undefined' && CFG && CFG.SHEET_ROSTER) || 'Roster',
    'Reported Hours',            // ledger you’re using for stats/highlighting
    (typeof CFG !== 'undefined' && CFG && CFG.SHEET_SYS_ISSUES) || 'System Reporting Issues' // optional
  ];

  // ---- scan helpers
  const missingHelpers = [];
  const presentHelpers = [];
  requiredHelpers.forEach(fn => {
    const t = typeof globalThis[fn];
    if (t === 'function') presentHelpers.push(fn);
    else missingHelpers.push(fn);
  });

  // ---- scan sheets
  const missingSheets = [];
  const presentSheets = [];
  requiredSheets.forEach(name => {
    if (!name) return;
    const sh = ss.getSheetByName(String(name));
    if (sh) presentSheets.push(name);
    else missingSheets.push(name);
  });

  // ---- log a pretty report
  const lines = [];
  lines.push('--- IRS CE Tools • Sanity Scan ---');
  lines.push('Timestamp: ' + new Date());
  lines.push('');
  lines.push('Helpers present (' + presentHelpers.length + '): ' + (presentHelpers.length ? presentHelpers.join(', ') : '—'));
  lines.push('Helpers MISSING (' + missingHelpers.length + '): ' + (missingHelpers.length ? missingHelpers.join(', ') : '—'));
  lines.push('');
  lines.push('Sheets present (' + presentSheets.length + '): ' + (presentSheets.length ? presentSheets.join(', ') : '—'));
  lines.push('Sheets MISSING (' + missingSheets.length + '): ' + (missingSheets.length ? missingSheets.join(', ') : '—'));
  Logger.log(lines.join('\n'));

  // ---- quick toast summary
  toast_(
    'Sanity Scan complete • Helpers: ' + presentHelpers.length + ' ok, ' + missingHelpers.length + ' missing • Sheets: ' + presentSheets.length + ' ok, ' + missingSheets.length + ' missing',
    (missingHelpers.length || missingSheets.length)
  );
}

/** Optional unified diagnostic runner */
function runDiagnostics() {
  try {
    const logs = [];
    logs.push('--- IRS CE Diagnostics ---');
    logs.push('Timestamp: ' + new Date());

    logs.push('Sheets present:');
    const ss = SpreadsheetApp.getActive();
    ss.getSheets().forEach(s => logs.push('  • ' + s.getName()));

    logs.push('Triggers present:');
    ScriptApp.getProjectTriggers().forEach(t => {
      logs.push(`  • ${t.getHandlerFunction()} (${t.getEventType()})`);
    });

    Logger.log(logs.join('\n'));
    SpreadsheetApp.getActiveSpreadsheet().toast('Diagnostics complete. Check Execution Log.');
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Diagnostics failed: ' + e.message, 'IRS CE Tools', 5);
    Logger.log(e.stack || e.message);
  }
}

/** Menu: Update Reporting Stats (calls reporting_stats.js) */
function updateReportingStatsMenu() {
  try {
    updateProgramReportedTotals();  // must exist in reporting_stats.js
    toast_('Reporting Stats updated from Reported Hours.');
  } catch (e) {
    toast_('Failed to update Reporting Stats: ' + e.message, true);
    Logger.log(e.stack || e);
  }
}

/** Menu: Sync Reported Hours → Master (calls sync_reported_to_master.js.gs) */
function syncReportedToMasterMenu() {
  try {
    syncMasterWithReportedHours(); // must exist in sync_reported_to_master.js.gs
    toast_('Reported Hours → Master sync complete.');
  } catch (e) {
    toast_('Failed to sync Reported Hours → Master: ' + e.message, true);
    Logger.log(e.stack || e);
  }
}