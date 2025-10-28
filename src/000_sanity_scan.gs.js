/** 000_sanity_scan.gs
 * Lightweight runtime checks for undefined functions + load order
 * No external permissions required.
 */

function scanDefinedGlobals_() {
  // enumerate global function names
  var g = (typeof globalThis !== 'undefined' ? globalThis : this);
  var names = Object.keys(g).filter(function (k) {
    try { return typeof g[k] === 'function'; } catch (e) { return false; }
  }).sort();
  Logger.log('Defined global functions (' + names.length + '):\n' + names.join('\n'));
  SpreadsheetApp.getActive().toast('Defined globals listed in log.', 'IRS CE Tools', 5);
}

function scanRequiredFunctions_() {
  // Add any function you expect menu items or other files to call
  var REQUIRED = [
    // menu items (from your menu.js)
    'buildCleanUpload','recheckMaster','createNightlyTrigger','removeNightlyTrigger',
    'createNightlyGroupSyncTrigger','removeNightlyGroupSyncTrigger',
    'markCleanAsReported','exportCleanToXlsx','syncGroupSheets','diagnoseGroupSync',
    'ingestSystemReportingIssues','applyReportingFixes','updateReportingStatsMenu',
    'dedupeRosterByEmail','backfillMasterFromRosterCombined',
    'updateRosterValidityFromIssues_','ensureAllTabs','runDiagnostics',

    // reporting stats + reported hours
    'updateProgramReportedTotals','syncMasterWithReportedHours',
    'dedupeReportedHoursByPtinProgram_',

    // master/clean helpers you referenced recently
    'writeCleanDataOnly_','appendToClean_',
    'normalizeAndFlagMasterPtins_','processMasterEdits',
    'mapFreeTextToStandardIssue_','clearMasterIssuesFromFixedIssues_',
    'getIssuesSheet_','applyReportingIssueValidationAndFormatting_',
    'appendIssuesRows_','syncMasterFromIssueSheet_',

    // batch runner used by markCleanAsReported
    'runJob','stepMarkReported_','runMarkReportedBatch'
  ];

  var g = (typeof globalThis !== 'undefined' ? globalThis : this);
  var missing = REQUIRED.filter(function (name) {
    try { return typeof g[name] !== 'function'; } catch (e) { return true; }
  });

  if (missing.length) {
    Logger.log('❌ Missing functions (' + missing.length + '):\n' + missing.join('\n'));
    SpreadsheetApp.getActive().toast('Missing: ' + missing.length + ' function(s). Check Logs.', 'IRS CE Tools', 8);
  } else {
    Logger.log('✅ All REQUIRED functions are defined.');
    SpreadsheetApp.getActive().toast('All REQUIRED functions are defined.', 'IRS CE Tools', 5);
  }
}

/** Adds two menu items so you can run these from Sheets */
function onOpen_sanityMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sanity Scan')
    .addItem('List defined globals (log)', 'scanDefinedGlobals_')
    .addItem('Check REQUIRED functions', 'scanRequiredFunctions_')
    .addToUi();
}