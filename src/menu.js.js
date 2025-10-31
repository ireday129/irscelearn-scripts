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
/* global dedupeRosterByEmail */
/* global updateRosterValidityFromIssues_, ensureAllTabs */
/* global backfillRosterFromMasterReported_ */
/* global highlightRosterFromReportedHours */

/**
 * IRS CE TOOLS MAIN MENU
 * Always registers handlers that exist in THIS file (wrappers),
 * so the menu never breaks even if core functions are missing.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('IRS CE Tools');

  menu
    .addItem('Build Clean Upload (resumable)', 'menu_buildCleanUpload')
    .addItem('Recheck Master for Issues', 'menu_recheckMaster')
    .addSeparator()
    .addItem('Create Nightly Clean Trigger', 'menu_createNightlyClean')
    .addItem('Remove Nightly Clean Trigger', 'menu_removeNightlyClean')
    .addItem('Create Nightly Group Sync Trigger', 'menu_createNightlyGroupSync')
    .addItem('Remove Nightly Group Sync Trigger', 'menu_removeNightlyGroupSync')
    .addSeparator()
    .addItem('Mark Clean as Reported (resumable)', 'menu_markCleanAsReported')
    .addItem('Export Clean as XLSX', 'menu_exportCleanToXlsx')
    .addSeparator()
    .addItem('Sync Group Sheets (strict)', 'menu_syncGroupSheets')
    .addItem('Diagnose Group Sync', 'menu_diagnoseGroupSync')
    .addSeparator()
    .addItem('Ingest System Reporting Issues', 'menu_ingestSystemReportingIssues')
    .addItem('Apply Reporting Fixes (manual)', 'menu_applyReportingFixes')
    .addSeparator()
    .addItem('Update Reporting Stats', 'menu_updateReportingStats')
    .addItem('Sync Reported Hours → Master', 'menu_syncReportedToMaster')
    .addSeparator()
    .addItem('Deduplicate Roster by Email', 'menu_dedupeRosterByEmail')
    .addItem('Backfill Roster from Master (Reported only)', 'menu_backfillRosterFromMasterReported')
    .addItem('Update Roster Validity from Issues', 'menu_updateRosterValidityFromIssues')
    .addSeparator()
    .addItem('Create/Repair Tabs (Clean & Issues only)', 'menu_ensureAllTabs')
    .addSeparator()
    .addItem('Diagnostics (Log all systems)', 'menu_runDiagnostics')
    .addItem('Highlight Roster from Reported Hours', 'menu_highlightRosterFromReportedHours');

  menu.addToUi();
}

/** --------------- SAFE WRAPPER UTIL --------------- */
function safeCall_(label, fn, args) {
  try {
    if (typeof fn !== 'function') {
      throw new Error(label + '() is not available in this project.');
    }
    const out = fn.apply(null, args || []);
    if (typeof toast_ === 'function') {
      toast_(label + ' ran.');
    } else {
      SpreadsheetApp.getActive().toast(label + ' ran.');
    }
    return out;
  } catch (err) {
    const msg = label + ' failed: ' + (err && err.message ? err.message : err);
    if (typeof toast_ === 'function') {
      toast_(msg, true);
    } else {
      SpreadsheetApp.getActive().toast(msg, 'IRS CE Tools', 5);
    }
    Logger.log(err && (err.stack || err));
    return null;
  }
}

/** --------------- MENU HANDLERS (WRAPPERS) --------------- */
function menu_buildCleanUpload() {
  return safeCall_('buildCleanUpload', (typeof buildCleanUpload === 'function') ? buildCleanUpload : null);
}
function menu_recheckMaster() {
  return safeCall_('recheckMaster', (typeof recheckMaster === 'function') ? recheckMaster : null);
}
function menu_createNightlyClean() {
  return safeCall_('createNightlyTrigger', (typeof createNightlyTrigger === 'function') ? createNightlyTrigger : null);
}
function menu_removeNightlyClean() {
  return safeCall_('removeNightlyTrigger', (typeof removeNightlyTrigger === 'function') ? removeNightlyTrigger : null);
}
function menu_createNightlyGroupSync() {
  return safeCall_('createNightlyGroupSyncTrigger', (typeof createNightlyGroupSyncTrigger === 'function') ? createNightlyGroupSyncTrigger : null);
}
function menu_removeNightlyGroupSync() {
  return safeCall_('removeNightlyGroupSyncTrigger', (typeof removeNightlyGroupSyncTrigger === 'function') ? removeNightlyGroupSyncTrigger : null);
}
function menu_markCleanAsReported() {
  return safeCall_('markCleanAsReported', (typeof markCleanAsReported === 'function') ? markCleanAsReported : null);
}
function menu_exportCleanToXlsx() {
  return safeCall_('exportCleanToXlsx', (typeof exportCleanToXlsx === 'function') ? exportCleanToXlsx : null);
}
function menu_syncGroupSheets() {
  // Support either a direct function or a menu alias provided elsewhere
  var fn = null;
  if (typeof syncGroupSheets === 'function') fn = syncGroupSheets;
  else if (typeof syncGroupSheetsMenu === 'function') fn = syncGroupSheetsMenu;
  return safeCall_('syncGroupSheets', fn);
}
function menu_diagnoseGroupSync() {
  return safeCall_('diagnoseGroupSync', (typeof diagnoseGroupSync === 'function') ? diagnoseGroupSync : null);
}
function menu_ingestSystemReportingIssues() {
  return safeCall_('ingestSystemReportingIssues', (typeof ingestSystemReportingIssues === 'function') ? ingestSystemReportingIssues : null);
}
function menu_applyReportingFixes() {
  return safeCall_('applyReportingFixes', (typeof applyReportingFixes === 'function') ? applyReportingFixes : null);
}
function menu_updateReportingStats() {
  return safeCall_('updateProgramReportedTotals', (typeof updateProgramReportedTotals === 'function') ? updateProgramReportedTotals : null);
}
function menu_syncReportedToMaster() {
  return safeCall_('syncMasterWithReportedHours', (typeof syncMasterWithReportedHours === 'function') ? syncMasterWithReportedHours : null);
}
function menu_dedupeRosterByEmail() {
  return safeCall_('dedupeRosterByEmail', (typeof dedupeRosterByEmail === 'function') ? dedupeRosterByEmail : null);
}
function menu_backfillRosterFromMasterReported() {
  return safeCall_(
    'backfillRosterFromMasterReported_',
    (typeof backfillRosterFromMasterReported_ === 'function') ? backfillRosterFromMasterReported_ : null
  );
}
function menu_updateRosterValidityFromIssues() {
  return safeCall_(
    'updateRosterValidityFromIssues_',
    (typeof updateRosterValidityFromIssues_ === 'function') ? updateRosterValidityFromIssues_ : null
  );
}
function menu_ensureAllTabs() {
  return safeCall_('ensureAllTabs', (typeof ensureAllTabs === 'function') ? ensureAllTabs : null);
}
function menu_runDiagnostics() {
  try {
    const logs = [];
    logs.push('--- IRS CE Diagnostics ---');
    logs.push('Timestamp: ' + new Date());
    const ss = SpreadsheetApp.getActive();
    logs.push('Sheets present:');
    ss.getSheets().forEach(function (s) { logs.push('  • ' + s.getName()); });
    logs.push('Triggers present:');
    ScriptApp.getProjectTriggers().forEach(function (t) {
      logs.push('  • ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')');
    });
    Logger.log(logs.join('\n'));
    SpreadsheetApp.getActiveSpreadsheet().toast('Diagnostics complete. Check Execution Log.');
  } catch (e2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Diagnostics failed: ' + e2.message, 'IRS CE Tools', 5);
    Logger.log(e2.stack || e2.message);
  }
}
function menu_highlightRosterFromReportedHours() {
  return safeCall_(
    'highlightRosterFromReportedHours',
    (typeof highlightRosterFromReportedHours === 'function') ? highlightRosterFromReportedHours : null
  );
}

/** --------------- EXPLICIT EXPORTS (defensive) --------------- */
// Make sure Apps Script can always discover the handlers
this.onOpen = onOpen;
this.menu_buildCleanUpload = menu_buildCleanUpload;
this.menu_recheckMaster = menu_recheckMaster;
this.menu_createNightlyClean = menu_createNightlyClean;
this.menu_removeNightlyClean = menu_removeNightlyClean;
this.menu_createNightlyGroupSync = menu_createNightlyGroupSync;
this.menu_removeNightlyGroupSync = menu_removeNightlyGroupSync;
this.menu_markCleanAsReported = menu_markCleanAsReported;
this.menu_exportCleanToXlsx = menu_exportCleanToXlsx;
this.menu_syncGroupSheets = menu_syncGroupSheets;
this.menu_diagnoseGroupSync = menu_diagnoseGroupSync;
this.menu_ingestSystemReportingIssues = menu_ingestSystemReportingIssues;
this.menu_applyReportingFixes = menu_applyReportingFixes;
this.menu_updateReportingStats = menu_updateReportingStats;
this.menu_syncReportedToMaster = menu_syncReportedToMaster;
this.menu_dedupeRosterByEmail = menu_dedupeRosterByEmail;
this.menu_backfillRosterFromMasterReported = menu_backfillRosterFromMasterReported;
this.menu_updateRosterValidityFromIssues = menu_updateRosterValidityFromIssues;
this.menu_ensureAllTabs = menu_ensureAllTabs;
this.menu_runDiagnostics = menu_runDiagnostics;
this.menu_highlightRosterFromReportedHours = menu_highlightRosterFromReportedHours;