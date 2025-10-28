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
}