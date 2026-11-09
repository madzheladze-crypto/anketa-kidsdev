const SPREADSHEET_ID = '1a2GJjM9_pPfxPY08m0MJhjgH3r9PVJCVHve3qtzqSxE';
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Заявки');
    if (!sheet || d.consent !== true) throw new Error('Invalid submission');
    sheet.appendRow([new Date(), d.age||'',d.childName||'',d.interests||'',d.idea||'',d.experience||'',d.goals||'',d.format||'',d.parentName||'',d.phone||'',d.telegram||'',d.callTime||'',d.contact24h||'',d.question||'',d.consent?'Да':'Нет',d.ageOther||'',d.extra||'',d.source||'']);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  } finally {lock.releaseLock();}
}