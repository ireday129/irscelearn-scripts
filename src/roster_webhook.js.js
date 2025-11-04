/** onEdit: mark Roster Valid? TRUE -> sync Master + Webhook **/
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_ROSTER) return;

    const map = mapRosterHeaders_(sh);
    if (!map || map.valid < 0) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Only react when editing Valid? column (map.valid is zero-based index)
    if (row >= 2 && col === (map.valid + 1)) {

      const newVal = e.value;
      if (!parseBool_(newVal)) return;

      // Grab row fields
      const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
      const first = String(vals[map.first] || '').trim();
      const last  = String(vals[map.last] || '').trim();
      const email = String(vals[map.email] || '').trim().toLowerCase();

      if (!email) {
        toast_('Valid marked TRUE but no Email present – webhook skipped.', true);
        return;
      }

      // ✅ Webhook payload
      const payload = {
        email: email,
        first_name: first,
        last_name: last
      };

      // ✅ POST the webhook
      try {
        const url = "https://irscelearn.com/wp-json/uap/v2/uap-5213-5214";
        const options = {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          payload: JSON.stringify(payload)
        };

        const response = UrlFetchApp.fetch(url, options);
        Logger.log("Webhook response: " + response.getContentText());

      } catch (err) {
        Logger.log("Webhook error: " + err.message);
        toast_("Webhook failed: " + err.message, true);
      }

      // ✅ Continue with your existing Master update behavior
      const ss = SpreadsheetApp.getActive();
      const master = mustGet_(ss, CFG.SHEET_MASTER);
      const mVals = master.getDataRange().getValues();
      if (mVals.length <= 1) return;

      const mHdr = normalizeHeaderRow_(mVals[0]);
      const mMap = mapHeaders_(mHdr);
      const body = mVals.slice(1);

      let updated = 0;
      for (let i=0;i<body.length;i++){
        const mrow = body[i];
        const em   = String(mrow[mMap.email]||'').trim().toLowerCase();
        if (em === email) {
          if (first) mrow[mMap.firstName] = first;
          if (last)  mrow[mMap.lastName]  = last;
          mrow[mMap.masterIssueCol] = '';
          updated++;
        }
      }
      if (updated) {
        master.getRange(2,1,body.length,mHdr.length).setValues(body);
        toast_(`Roster → Master synced + webhook sent for ${email}`);
      }
    }
  } catch (err) {
    toast_('onEdit error: ' + err.message, true);
  }
  /** Alias expected by sanity scan / menu */
function triggerRosterValidWebhookMaybe() {
  try {
    if (typeof postWebhookOnRosterValidMaybe === 'function') {
      return postWebhookOnRosterValidMaybe();
    }
    toast_('Webhook helper postWebhookOnRosterValidMaybe not found.', true);
  } catch (e) {
    toast_('Roster webhook error: ' + e.message, true);
  }
}
}
/**
 * Roster Valid? → Webhook
 * ------------------------------------------------------------
 * Fires ONLY when the Roster "Valid?" checkbox is set to TRUE.
 * This MUST be wired as an INSTALLABLE onEdit trigger:
 *   Triggers → Add Trigger → onEditRosterWebhook → From spreadsheet → On edit
 *
 * Requires helpers from utils: toast_, mapRosterHeaders_, parseBool_, mustGet_, normalizeHeaderRow_, mapHeaders_
 */

const ROSTER_VALID_WEBHOOK_URL = "https://irscelearn.com/wp-json/uap/v2/uap-5213-5214";

/**
 * Installable onEdit handler (NOT the simple onEdit)
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEditRosterWebhook(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== CFG.SHEET_ROSTER) return;

    const map = mapRosterHeaders_(sh);
    if (!map || map.valid < 0) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2) return;                               // skip header
    if (col !== (map.valid + 1)) return;               // only when Valid? column edited

    // Only fire when turned TRUE
    if (!parseBool_(e.value)) return;

    // Gather payload fields
    const vals  = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const first = String(vals[map.first] || '').trim();
    const last  = String(vals[map.last]  || '').trim();
    const email = String(vals[map.email] || '').toLowerCase().trim();
    const ptin  = map.ptin >= 0 ? formatPtinP0_(vals[map.ptin] || '') : '';

    if (!email) {
      toast_('Valid? marked TRUE but Email is blank → webhook skipped.', true);
      return;
    }

    // POST webhook
    sendRosterValidWebhook_(email, first, last, ptin);
    toast_(`Webhook queued for ${email}`);

    // Optional: reset row background to white when Valid? is checked
    try {
      const rowRange = sh.getRange(row, 1, 1, sh.getLastColumn());
      rowRange.setBackground(null); // clear (defaults back to white)
    } catch (bgErr) {
      Logger.log('Background reset failed: ' + bgErr.message);
    }

    // Optional: light sync to Master (clear issue text for same email)
    try {
      const ss     = SpreadsheetApp.getActive();
      const master = mustGet_(ss, CFG.SHEET_MASTER);
      const mVals  = master.getDataRange().getValues();
      if (mVals.length > 1) {
        const mHdr = normalizeHeaderRow_(mVals[0]); 
        const mm   = mapHeaders_(mHdr);
        const body = mVals.slice(1);
        let changed = 0;
        for (let i = 0; i < body.length; i++) {
          const r = body[i];
          const em = mm.email != null ? String(r[mm.email] || '').toLowerCase().trim() : '';
          if (em === email && mm.masterIssueCol != null) {
            if (String(r[mm.masterIssueCol] || '').trim() !== '') {
              r[mm.masterIssueCol] = '';
              changed++;
            }
          }
        }
        if (changed) {
          master.getRange(2, 1, body.length, mHdr.length).setValues(body);
        }
      }
    } catch (syncErr) {
      Logger.log('Master sync after webhook failed: ' + syncErr.message);
    }

  } catch (err) {
    toast_('onEditRosterWebhook error: ' + err.message, true);
    Logger.log(err && err.stack || err);
  }
}

/**
 * Helper to POST the webhook payload
 */
function sendRosterValidWebhook_(email, first, last, ptin) {
  const payload = {
    email: email,
    first_name: first,
    last_name: last,
    ptin: ptin || ''
  };
  const options = {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };
  const res = UrlFetchApp.fetch(ROSTER_VALID_WEBHOOK_URL, options);
  Logger.log("Webhook response: " + res.getResponseCode() + " " + res.getContentText());
}

/**
 * Manual tester: select a data row on Roster and run this function.
 * Useful for confirming the webhook without editing the checkbox.
 */
function triggerRosterValidWebhookMaybe() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CFG.SHEET_ROSTER);
    if (!sh) { toast_('Roster sheet not found.', true); return; }

    const map = mapRosterHeaders_(sh);
    if (!map) { toast_('Roster header map failed.', true); return; }

    const r = sh.getActiveCell() ? sh.getActiveCell().getRow() : 0;
    if (r < 2) { toast_('Select a data row on the Roster to test.', true); return; }

    const vals  = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
    const first = String(vals[map.first] || '').trim();
    const last  = String(vals[map.last]  || '').trim();
    const email = String(vals[map.email] || '').toLowerCase().trim();
    const ptin  = map.ptin >= 0 ? formatPtinP0_(vals[map.ptin] || '') : '';

    if (!email) { toast_('Selected row has no Email.', true); return; }
    sendRosterValidWebhook_(email, first, last, ptin);
    toast_(`Manual webhook sent for ${email}`);
  } catch (e) {
    toast_('Manual webhook failed: ' + e.message, true);
  }
}