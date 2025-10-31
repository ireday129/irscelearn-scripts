/**
 * SIMPLE onEdit webhook for Roster "Valid?" checkbox
 * Fires whenever the "Valid?" column is set to TRUE on the Roster sheet.
 * - No additional filters (email optional)
 * - No Master syncing or other side effects
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    var sh = e.range.getSheet();
    var rosterName = (typeof CFG !== 'undefined' && CFG.SHEET_ROSTER) ? String(CFG.SHEET_ROSTER) : 'Roster';
    if (String(sh.getName()).trim().toLowerCase() !== rosterName.trim().toLowerCase()) return;

    // Read header row and build a loose indexer
    var lastCol = sh.getLastColumn();
    if (lastCol < 1) return;

    var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (s) {
      return String(s || '').replace(/\uFEFF/g, '').trim();
    });
    var lower = hdr.map(function (h) { return h.toLowerCase().replace(/\s+/g, ' '); });

    function idx() {
      for (var a = 0; a < arguments.length; a++) {
        var target = String(arguments[a] || '').toLowerCase().replace(/\s+/g, ' ');
        var i = lower.indexOf(target);
        if (i >= 0) return i;
      }
      return -1;
    }

    // Column indexes (0-based)
    var iValid = idx('valid?', 'valid');
    if (iValid < 0) return; // can't identify the Valid? column

    // Only react when editing the Valid? column (row >=2)
    var row = e.range.getRow();
    var col = e.range.getColumn();
    if (row < 2 || col !== (iValid + 1)) return;

    // Only fire on TRUE/checked
    var newVal = (typeof e.value !== 'undefined') ? String(e.value).trim().toLowerCase() : '';
    var isTrue = (newVal === 'true' || newVal === '1' || newVal === 'yes' || newVal === 'y' || newVal === 'checked' || newVal === 'on' || newVal === '☑' || newVal === '✓' || newVal === '✔' );
    // For checkboxes, Apps Script uses 'TRUE'/'FALSE'
    if (newVal === 'true' || newVal === 'false') {
      // already handled by the lowercase checks above
    } else if (String(e.value).toUpperCase() === 'TRUE') {
      isTrue = true;
    }

    if (!isTrue) return;

    // Grab row values (email/first/last are optional for this webhook)
    var vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];

    var iFirst = idx('attendee first name', 'first name');
    var iLast  = idx('attendee last name', 'last name');
    var iEmail = idx('email');

    var payload = {
      email:  iEmail >= 0 ? String(vals[iEmail] || '').trim().toLowerCase() : '',
      first_name: iFirst >= 0 ? String(vals[iFirst] || '').trim() : '',
      last_name:  iLast  >= 0 ? String(vals[iLast]  || '').trim() : ''
    };

    // POST the webhook regardless of blanks
    var url = "https://irscelearn.com/wp-json/uap/v2/uap-5213-5214";
    var options = {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    };

    var resp = UrlFetchApp.fetch(url, options);
    var txt = resp ? resp.getContentText() : '';
    Logger.log("Roster Valid? webhook fired. Row " + row + ". Response: " + txt);
    if (typeof toast_ === 'function') {
      toast_("Webhook fired for row " + row + (payload.email ? (" (" + payload.email + ")") : ""), false);
    }

  } catch (err) {
    Logger.log("Roster Valid? webhook error: " + err.message);
    if (typeof toast_ === 'function') toast_("Webhook error: " + err.message, true);
  }
}