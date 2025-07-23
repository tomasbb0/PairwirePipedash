// ===== Code.gs =====

// ▶︎ This must be YOUR template file’s ID (from the URL you gave)
const TEMPLATE_FILE_ID = '11PoyhMdTB3yG-O3pBODImaTMmpRg68_5Q-txOlDHtfY';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Campaign Launcher')
    .addItem('Launch Bulk Campaign', 'openCampaignDialog')
    .addToUi();
}

function getCampaignList() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const leads = ss.getSheetByName('Leads');
  const last  = leads.getLastRow();
  if (last < 2) return [];
  return [...new Set(
    leads.getRange(2,2,last-1,1)
         .getValues().flat().filter(v=>v)
  )];
}

function openCampaignDialog() {
  const html = HtmlService.createTemplateFromFile('Dialog');
  html.campaigns = getCampaignList();
  SpreadsheetApp.getUi()
    .showModalDialog(html.evaluate().setWidth(350).setHeight(260), 'Select Campaign');
}

// 🔥 Single‐call runner
function runAllSteps(campaignId) {
  // 1️⃣ Copy the entire template SPREADSHEET
  const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
  const newFile      = templateFile.makeCopy(`${campaignId} | Bulk Emails Pairwire`);
  const newSs        = SpreadsheetApp.openById(newFile.getId());

  // 2️⃣ Write Campaign ID & Context
  const inst      = newSs.getSheetByName('Instructions');
  const campData  = SpreadsheetApp
                     .getActiveSpreadsheet()
                     .getSheetByName('Campaigns')
                     .getDataRange()
                     .getValues();
  const [hdr,...rows] = campData;
  const idIdx     = hdr.indexOf('Campaign ID');
  const ctxIdx    = hdr.indexOf('Campaign Context');
  const match     = rows.find(r=>r[idIdx]===campaignId) || [];
  inst.getRange('B6').setValue(campaignId);
  inst.getRange('B5').setValue(match[ctxIdx]||'');

  // 3️⃣ Filter Leads & populate RawLeadData
  const leads     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads');
  const last      = leads.getLastRow();
  const allData   = leads.getRange(2,1,last-1,19).getValues();
  const filtered  = allData
    .filter(r=>r[1]===campaignId)
    .map(r=>{ r.splice(1,1); return r; });
  const header    = leads.getRange(1,1,1,19).getValues()[0].filter((_,i)=>i!==1);
  const rawSheet  = newSs.getSheetByName('RawLeadData');
  rawSheet.clearContents();
  rawSheet.getRange(1,1,1,header.length).setValues([header]);
  if (filtered.length) {
    rawSheet.getRange(2,1,filtered.length,filtered[0].length)
            .setValues(filtered);
  }

  // 4️⃣ Return link & count
  return { 
    url: newSs.getUrl() + '#gid=' + rawSheet.getSheetId(), 
    count: filtered.length 
  };
}
