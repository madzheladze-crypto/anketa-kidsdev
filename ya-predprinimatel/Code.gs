const SPREADSHEET_ID = '1a2GJjM9_pPfxPY08m0MJhjgH3r9PVJCVHve3qtzqSxE';
function doPost(e) {
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d=JSON.parse(e.postData.contents);
    const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Заявки');
    if(!sh || d["13"]!==true) throw new Error("Missing consent or sheet");
    const val=n=>Array.isArray(d[n])?d[n].join("; "):String(d[n]??"");
    sh.appendRow([new Date(),val(0),val(1),val(2),val(3),val(4),val(5),val(6),val(7),val(8),val(9),val(10),val(11),val(12),"Да",val("extra0"),Object.keys(d).filter(k=>k.startsWith("extra")&&k!=="extra0").map(k=>k+": "+d[k]).join("; "),val("source")]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON)}
  finally{lock.releaseLock()}
}