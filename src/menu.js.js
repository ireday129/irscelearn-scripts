/* eslint-env es6 */
/* eslint-env googleappsscript */

/* global SpreadsheetApp, ScriptApp, Logger */
/* global toast_ */
/* global updateProgramReportedTotals */
/* global syncMasterWithReportedHours */
/* global buildCleanUpload, recheckMaster */
/* global createNightlyTrigger, removeNightlyTrigger */
/* global createNightlyGroupSyncTrigger, removeNightlyGroupSyncTrigger */
/* global markCleanAsReported, exportCleanToXlsx */
/* global syncGroupSheets, diagnoseGroupSync */
/* global ingestSystemReportingIssues, applyReportingFixes */
/* global dedupeRosterByEmail, backfillMasterFromRosterCombined */
/* global updateRosterValidityFromIssues_, ensureAllTabs */
/* global highlightRosterFromReportedHoursMenu */

/** IRS CE TOOLS MAIN MENU **/
function onOpen(e) {
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
    .addItem('Sync Reported Hours -> Master', 'syncReportedToMasterMenu')
    .addSeparator()
    .addItem('Deduplicate Roster by Email', 'dedupeRosterByEmail')
    .addItem('Backfill Roster from Master (Reported only)', 'backfillRosterFromMasterReported_')
    .addItem('Update Roster Validity from Issues', 'updateRosterValidityFromIssues_')
    .addSeparator()
    .addItem('Create/Repair Tabs (Clean & Issues only)', 'ensureAllTabs')
    .addSeparator()
    .addItem('Diagnostics (Log all systems)', 'runDiagnostics')
    .addItem('Highlight Roster from Reported Hours', 'highlightRosterFromReportedHoursMenu');

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
      logs.push('  • ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')');
    });

    Logger.log(logs.join('\n'));
    SpreadsheetApp.getActiveSpreadsheet().toast('Diagnostics complete. Check Execution Log.');
  } catch (e2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Diagnostics failed: ' + e2.message, 'IRS CE Tools', 5);
    Logger.log(e2.stack || e2.message);
  }
}

/** Menu: Update Reporting Stats (calls reporting_stats.js) */
function updateReportingStatsMenu() {
  try {
    if (typeof updateProgramReportedTotals !== 'function') {
      throw new Error('updateProgramReportedTotals() is not loaded in this project.');
    }
    updateProgramReportedTotals();
    toast_('Reporting Stats updated from Reported Hours.');
  } catch (e3) {
    toast_('Failed to update Reporting Stats: ' + e3.message, true);
    Logger.log(e3.stack || e3);
  }
}

/** Menu: Sync Reported Hours → Master (calls sync_reported_to_master.js.gs) */
function syncReportedToMasterMenu() {
  try {
    if (typeof syncMasterWithReportedHours !== 'function') {
      throw new Error('syncMasterWithReportedHours() is not loaded in this project.');
    }
    syncMasterWithReportedHours(); // performs the upsert
    toast_('Reported Hours → Master sync complete.');
  } catch (e4) {
    toast_('Failed to sync Reported Hours → Master: ' + e4.message, true);
    Logger.log(e4.stack || e4);
  }
}