/** CONFIG / CONSTANTS **/
const CFG = {
  SHEET_MASTER: 'Master',
  SHEET_CLEAN: 'Clean',
  SHEET_ISSUES: null, // Corrected: Sheet was deleted
  SHEET_ROSTER: 'Roster',
  SHEET_STATS:  'Reporting Stats',
  SHEET_SYS_ISSUES: 'System Reporting Issues',
  SHEET_GROUP_CONFIG: 'Group Config',   // headers: Group ID | Group Name | Spreadsheet URL
  GROUP_TARGET_SHEET: 'Reporting',      // exact tab name in group workbooks

  COL_HEADERS: {
    firstName: 'Attendee First Name',
    lastName:  'Attendee Last Name',
    ptin:      'Attendee PTIN',
    email:     'Email',
    program:   'Program Number',
    hours:     'CE Hours Awarded',
    completion:'Program Completion Date',
    group:    'Group', 
    masterIssueCol: 'Reporting Issue?',
    reportedCol:    'Reported?',
    updatedCol:     'Updated?',
    reportedAtCol:  'Reported At'
  },

  // ISSUE_HEADERS array has been removed as the Reporting Issue sheet was deleted.

  ROSTER_HEADERS: [
    'Attendee First Name','Attendee Last Name','Attendee PTIN','Email', 'Group' 
  ],

  SYS_ISSUE_HEADERS: [
    'Attendee First Name','Attendee Last Name','PTIN','Program Number',
    'CE Hours Awarded','Program Completion Date','Status'
  ],

  REPORTING_ISSUE_CHOICES: [
    'PTIN does not exist',
    'PTIN & name do not match',
    'Missing PTIN',
    'Other'
  ],

  TRIM_ALL_FIELDS: true,

  // Nightly rebuild time
  NIGHTLY_HOUR: 2,
  NIGHTLY_MINUTE: 0,

  // Stats
  AFTR_PROGRAM_2025: 'RTPMH-A-00004-25-S'
};
