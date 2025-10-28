/** EXPORT CLEAN → XLSX (exclude Email & Reporting Issue?) **/

function exportCleanToXlsx() {
  const ss = SpreadsheetApp.getActive();
  const clean = ss.getSheetByName(CFG.SHEET_CLEAN);
  if (!clean) { toast_('Clean sheet not found.', true); return; }
  const data = clean.getDataRange().getValues();
  if (data.length <= 1) { toast_('Clean sheet is empty.', true); return; }

  const hdr = data[0].map(s=>String(s||'').trim());
  const exclude = new Set(['Email', 'Reporting Issue?']);
  const keepIdx = [], exportHdr = [];
  hdr.forEach((h, i) => { if (!exclude.has(h)) { keepIdx.push(i); exportHdr.push(h); } });
  const exportData = data.slice(1).map(row => keepIdx.map(i => row[i]));

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMddyyyy');
  const temp = SpreadsheetApp.create('TEMP_Clean_' + todayStr);
  const tempId = temp.getId();
  const tempSheet = temp.getSheets()[0];
  tempSheet.setName('Clean');

  tempSheet.getRange(1,1,1,exportHdr.length).setValues([exportHdr]);
  if (exportData.length) tempSheet.getRange(2,1,exportData.length, exportHdr.length).setValues(exportData);
  const tHdr = tempSheet.getRange(1,1,1, exportHdr.length).getValues()[0].map(s=>String(s||'').trim());
  const iC = tHdr.indexOf('Program Completion Date');
  if (iC >= 0 && exportData.length > 0) tempSheet.getRange(2, iC+1, exportData.length, 1).setNumberFormat('mm/dd/yyyy');
  SpreadsheetApp.flush();

  try {
    const filename = todayStr + '.xlsx';
    const blob = getXlsxBlob_(tempId, filename);
    const xlsx = DriveApp.createFile(blob);
    DriveApp.getFileById(tempId).setTrashed(true);
    showExportDialog_(xlsx.getUrl(), filename);
    toast_('Clean exported (Email & Reporting Issue? omitted): ' + xlsx.getName());
  } catch (err) {
    toast_('Export failed: ' + err.message, true);
    Logger.log(err.stack || err.message);
  }
}

function getXlsxBlob_(fileId, filename) {
  const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  try {
    if (typeof Drive!=='undefined' && Drive.Files && Drive.Files.export) {
      const resp = Drive.Files.export(fileId, mime);
      const blob = (resp && typeof resp.getBlob==='function') ? resp.getBlob() : resp;
      blob.setName(filename);
      return blob;
    }
  } catch(_){}
  const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
              '/export?mimeType=' + encodeURIComponent(mime);
  const resp = UrlFetchApp.fetch(url, {
    method:'get', headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}, muteHttpExceptions:true
  });
  if (resp.getResponseCode()<200 || resp.getResponseCode()>=300)
    throw new Error('Drive export failed: ' + resp.getContentText());
  const blob = resp.getBlob(); blob.setName(filename); return blob;
}

function showExportDialog_(url, filename) {
  const html = HtmlService.createHtmlOutput(`<!doctype html>
<html><head><meta charset="utf-8"><title>Clean Export Ready</title>
<style>body{font-family:Arial,sans-serif;padding:16px}a.btn{display:inline-block;padding:10px 14px;text-decoration:none;border-radius:6px;border:1px solid #1a73e8}.primary{background:#1a73e8;color:#fff}.row{margin-top:10px}.muted{color:#555;font-size:12px;margin-top:8px}button{padding:8px 12px;border-radius:6px;border:1px solid #ccc;background:#f6f6f6;cursor:pointer}input{width:100%;padding:6px 8px;font-size:12px}</style>
</head><body>
  <h2>Clean export ready</h2>
  <div class="row">
    <a class="btn primary" href="${url}" target="_blank" rel="noopener">Download ${filename}</a>
    <span style="display:inline-block;width:8px"></span>
    <button onclick="google.script.host.close()">Close</button>
  </div>
  <div class="row">
    <input id="dl" type="text" value="${url}" readonly />
    <div class="row">
      <button onclick="(function(){const el=document.getElementById('dl');el.select();el.setSelectionRange(0,99999);document.execCommand('copy')})()">Copy link</button>
      <span class="muted">Opens in Drive; you can also copy the link.</span>
    </div>
  </div>
</body></html>`).setWidth(420).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, 'Export Complete');
}