/**
 * One-off utility: sync a "Read Me" sheet from the master workbook
 * into each group workbook, if missing.
 *
 * - Uses the "Read Me" sheet in the current (master) spreadsheet as the template.
 * - For each group in Group Config / Groups:
 *    - If the target spreadsheet has no "Read Me" sheet, copy it in.
 *    - Ensure "Read Me" appears immediately before the "Reporting" sheet.
 * - Does NOT modify any other sheets.
 */
function syncGroupReadMe() {
  const ss = SpreadsheetApp.getActive();

  // Template sheet in your main EarnTaxCE workbook
  const template = ss.getSheetByName('Read Me');
  if (!template) {
    toast_('syncGroupReadMe: No "Read Me" sheet found in the master workbook.', true);
    return;
  }

  // Use the same group catalog as group_sync: "Group Config" (preferred) or "Groups"
  const groups = readGroupsCatalog_(ss);
  if (!groups.length) {
    toast_('syncGroupReadMe: No groups found in "Group Config"/"Groups".', true);
    return;
  }

  let created = 0;
  let reordered = 0;
  let skipped = 0;
  let errors = 0;

  groups.forEach(entry => {
    const gName = String(entry.groupName || entry.group || '').trim();
    const gUrl  = String(entry.url || '').trim();

    if (!gName || !gUrl) {
      Logger.log('syncGroupReadMe: missing group name or URL in catalog entry: ' +
        JSON.stringify(entry));
      errors++;
      return;
    }

    let targetSS;
    try {
      targetSS = openSpreadsheetByUrlOrId_(gUrl);
    } catch (e) {
      Logger.log(`syncGroupReadMe: cannot open target for "${gName}" (${gUrl}): ${e.message}`);
      errors++;
      return;
    }

    if (!targetSS) {
      Logger.log(`syncGroupReadMe: target spreadsheet not found for "${gName}" (${gUrl}).`);
      errors++;
      return;
    }

    // Get / create Read Me sheet
    let readMe = targetSS.getSheetByName('Read Me');
    if (!readMe) {
      try {
        readMe = template.copyTo(targetSS);
        readMe.setName('Read Me');
        created++;
      } catch (e) {
        Logger.log(`syncGroupReadMe: failed to copy Read Me to "${gName}": ${e.message}`);
        errors++;
        return;
      }
    } else {
      skipped++;
    }

    // Try to position "Read Me" immediately before "Reporting" (if that sheet exists)
    try {
      const reporting = targetSS.getSheetByName('Reporting');
      if (reporting && readMe) {
        const sheets = targetSS.getSheets();
        const reportingIndex = sheets.indexOf(reporting); // 0-based
        const readMeIndex    = sheets.indexOf(readMe);    // 0-based

        if (reportingIndex >= 0 && readMeIndex >= 0 && readMeIndex !== reportingIndex - 1) {
          // Move Read Me to (reportingIndex) in 1-based indexing → just before Reporting
          targetSS.setActiveSheet(readMe);
          targetSS.moveActiveSheet(reportingIndex + 1);
          reordered++;
        }
      }
    } catch (e) {
      Logger.log(`syncGroupReadMe: failed to position Read Me in "${gName}": ${e.message}`);
      // Non-fatal; keep going
    }
  });

  toast_(
    `syncGroupReadMe done: created ${created}, reordered ${reordered}, ` +
    `skipped existing ${skipped}, errors ${errors}.`
  );
}

// Make sure Apps Script runtime can see this for menus/manual runs
try {
  this.syncGroupReadMe = this.syncGroupReadMe || syncGroupReadMe;
} catch (e) {
  // no-op
}