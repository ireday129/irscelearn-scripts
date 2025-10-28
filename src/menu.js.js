/** IRS CE TOOLS MAIN MENU **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('IRS CE Tools');

  menu
    .addItem('Build Clean Upload (resumable)', 'buildCleanUpload')
    .addItem('Recheck Master for Issues', 'recheckMaster')
    .addSeparator()
    .addItem('Create Nightly Clean Trigger', 'createNightlyTrigger')
    .addItem('Remove Nightly Clean Trigger', 'removeNightlyTrigger')
    .addItem('Create Nightly Group Sync Trigger', 'createNightlyGroupSyncTrigger')
    .addItem('Remove Nightly Group Sync Trigger', 'removeNightlyGroupSyncTrigger')
    .addSeparator()
    .addItem('Mark Clean as Reported (resumable)', 'markCleanAsReported')
    .addItem('Export Clean as XLSX', 'exportCleanToXlsx')
    .addSeparator()
    .addItem('Sync Group Sheets (strict)', 'syncGroupSheets')
    .addItem('Diagnose Group Sync', 'diagnoseGroupSync')
    .addSeparator()
    .addItem('Ingest System Reporting Issues', 'ingestSystemReportingIssues')
    .addItem('Apply Reporting Fixes (manual)', 'applyReportingFixes')
    .addSeparator()
    .addItem('Update Reporting Stats', 'updateReportingStatsMenu')
    .addItem('Sync Reported Hours \u2192 Master', 'syncReportedToMasterMenu') // ← new menu item
    .addSeparator()
    .addItem('Deduplicate Roster by Email', 'dedupeRosterByEmail')
    .addItem('Backfill Master from Roster', 'backfillMasterFromRosterCombined') // COMBINED ITEM
    .addItem('Update Roster Validity from Issues', 'updateRosterValidityFromIssues_')
    .addSeparator()
    .addItem('Create/Repair Tabs (Clean & Issues only)', 'ensureAllTabs')
    .addSeparator()
    .addItem('Diagnostics (Log all systems)', 'runDiagnostics');

  menu.addToUi();
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